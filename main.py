from __future__ import annotations

import csv
import json
import queue
import re
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pyautogui

try:
    import pyperclip
except ImportError:
    pyperclip = None

from transformacoes import TRANSFORMACOES

BASE_DIR = Path(__file__).resolve().parent.parent
APP_DIR = BASE_DIR / 'app'
ROTEIROS_DIR = BASE_DIR / 'pyBossAll\\roteiros'
TABELAS_DIR = BASE_DIR / 'pyBossAll\\tabelas'

for pasta in (APP_DIR, ROTEIROS_DIR, TABELAS_DIR):
    pasta.mkdir(parents=True, exist_ok=True)

pyautogui.FAILSAFE = True


HELP_TEXT = '''COMANDOS DISPONÍVEIS NO ROTEIRO (JSON)

Formato esperado do arquivo:
[
  {"obs": "Descrição", "acao": "wait:0.3"},
  {"obs": "Descrição", "x": 100, "y": 200, "acao": "click"}
]

Comandos:
- wait:0.3
  Espera em segundos.

- press:enter
  Pressiona uma tecla.

- hotkey:ctrl+s
  Atalho. Separe as teclas por +.

- type:Texto aqui
  Digita/cola um texto.

- type:'500'
  Digita exatamente '500', incluindo as aspas simples.

- move
  Requer x e y no passo.

- click
  Requer x e y no passo.

- double_click
  Requer x e y no passo.

- right_click
  Requer x e y no passo.

- hold_click
  Pressiona e segura o mouse na coordenada.

- release_click
  Solta o mouse na coordenada.

- transform_copy:nome_funcao
  Copia o texto selecionado com Ctrl+C, aplica a função do arquivo transformacoes.py
  e deixa o resultado no clipboard.

- repeat:3|press:down
  Repete uma subação N vezes.

- repeat:i+1|press:tab
  Repete uma subação usando a variável i do loop externo.

- click + wait:0.5 + type:Exemplo
  Executa múltiplas ações no mesmo passo.

- reg_table:campo
  Usa o campo da tabela selecionada no dropdown.

- reg_table:minha_tabela.nome
  Usa explicitamente o campo nome da tabela minha_tabela.csv.

Regras da tabela CSV:
- Separador ';'
- Deve possuir cabeçalho.
- A coluna STATUS será criada automaticamente, se não existir.
- Registros com STATUS = OK serão ignorados em próximas execuções.
- Nesta implementação, o STATUS vira OK ao final de cada iteração executada com sucesso.

Observações:
- Delay inicial: aplicado uma única vez antes do início da execução.
- Delay global: aplicado entre os passos do roteiro.
- Repetições: número máximo de iterações.
- Se o roteiro usar tabela e acabarem os registros pendentes, a execução para.
- Pressione o canto superior esquerdo da tela para acionar o FAILSAFE do pyautogui.
'''


@dataclass
class TableCursor:
    name: str
    path: Path
    rows: list[dict[str, str]]
    fieldnames: list[str]
    pending_indexes: list[int]
    current_index: int | None = None

    @classmethod
    def load(cls, table_name: str) -> 'TableCursor':
        path = TABELAS_DIR / f'{table_name}.csv'
        if not path.exists():
            raise FileNotFoundError(f'Tabela não encontrada: {path.name}')

        with path.open('r', encoding='utf-8-sig', newline='') as f:
            sample = f.read(2048)
            f.seek(0)
            delimiter = ';' if sample.count(';') >= sample.count(',') else ','
            reader = csv.DictReader(f, delimiter=delimiter)
            rows = [dict(r) for r in reader]
            fieldnames = list(reader.fieldnames or [])

        if 'STATUS' not in fieldnames:
            fieldnames.append('STATUS')
            for row in rows:
                row['STATUS'] = row.get('STATUS', '')
        else:
            for row in rows:
                row['STATUS'] = (row.get('STATUS') or '').strip()

        pending_indexes = [
            idx for idx, row in enumerate(rows)
            if (row.get('STATUS') or '').strip().upper() != 'OK'
        ]
        return cls(table_name, path, rows, fieldnames, pending_indexes)

    def has_pending(self) -> bool:
        return len(self.pending_indexes) > 0

    def next_row(self) -> dict[str, str]:
        if not self.pending_indexes:
            raise RuntimeError(f'Tabela {self.name} sem registros pendentes.')
        self.current_index = self.pending_indexes.pop(0)
        return self.rows[self.current_index]

    def get_current_value(self, field: str) -> str:
        if self.current_index is None:
            self.next_row()
        assert self.current_index is not None
        row = self.rows[self.current_index]
        if field not in row:
            raise KeyError(f'Campo {field} não existe na tabela {self.name}.')
        return row.get(field, '') or ''

    def mark_current_ok(self) -> None:
        if self.current_index is None:
            return
        self.rows[self.current_index]['STATUS'] = 'OK'

    def reset_current(self) -> None:
        if self.current_index is not None:
            self.pending_indexes.insert(0, self.current_index)
            self.current_index = None

    def save(self) -> None:
        with self.path.open('w', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=self.fieldnames, delimiter=';')
            writer.writeheader()
            writer.writerows(self.rows)


class ScriptRunner:
    def __init__(
        self,
        log_callback: Callable[[str], None],
        is_stop_requested: Callable[[], bool],
        get_delay: Callable[[], float],
        get_start_delay: Callable[[], float],
    ) -> None:
        self.log = log_callback
        self.is_stop_requested = is_stop_requested
        self.get_delay = get_delay
        self.get_start_delay = get_start_delay

    def run(
        self,
        steps: list[dict[str, Any]],
        repetitions: int,
        selected_table: str | None,
    ) -> None:
        table_contexts: dict[str, TableCursor] = {}
        default_table = None
        if selected_table:
            default_table = TableCursor.load(selected_table)
            table_contexts[selected_table] = default_table
            self.log(f'Tabela selecionada: {selected_table} | pendentes: {len(default_table.pending_indexes)}')

        explicit_tables = self._extract_explicit_table_names(steps)
        for name in explicit_tables:
            if name not in table_contexts:
                table_contexts[name] = TableCursor.load(name)
                self.log(f'Tabela referenciada no roteiro: {name} | pendentes: {len(table_contexts[name].pending_indexes)}')

        uses_table = bool(default_table or explicit_tables)

        start_delay = self.get_start_delay()
        if start_delay > 0:
            self.log(f'Delay inicial: aguardando {start_delay} segundo(s) antes de iniciar.')
            time.sleep(start_delay)

        for i in range(repetitions):
            if self.is_stop_requested():
                self.log('Execução interrompida pelo usuário.')
                return

            if uses_table and not self._has_rows_available(steps, table_contexts, default_table):
                self.log('Sem registros pendentes suficientes para continuar. Execução finalizada.')
                return

            self.log(f'--- Iteração {i + 1}/{repetitions} ---')
            acquired_tables: set[str] = set()
            try:
                for step_idx, step in enumerate(steps, start=1):
                    if self.is_stop_requested():
                        self.log('Execução interrompida pelo usuário.')
                        return
                    obs = step.get('obs', '')
                    if obs:
                        self.log(f'Passo {step_idx}: {obs}')
                    self._execute_action_string(step, str(step.get('acao', '')).strip(), i, table_contexts, default_table, acquired_tables)
                    delay = self.get_delay()
                    if delay > 0:
                        time.sleep(delay)

                for table_name in acquired_tables:
                    table_contexts[table_name].mark_current_ok()
                    table_contexts[table_name].save()
                    table_contexts[table_name].current_index = None
                    self.log(f'Tabela {table_name}: registro marcado com STATUS=OK')
            except Exception as exc:
                for table_name in acquired_tables:
                    table_contexts[table_name].reset_current()
                raise RuntimeError(f'Erro na iteração {i + 1}: {exc}') from exc

        self.log('Execução concluída.')

    def _extract_explicit_table_names(self, steps: list[dict[str, Any]]) -> set[str]:
        names: set[str] = set()
        pattern = re.compile(r'reg_table:([A-Za-z0-9_]+)\.([A-Za-z0-9_]+)')
        for step in steps:
            action = str(step.get('acao', ''))
            for table_name, _field in pattern.findall(action):
                names.add(table_name)
        return names

    def _has_rows_available(
        self,
        steps: list[dict[str, Any]],
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
    ) -> bool:
        needed_tables = self._tables_needed_for_iteration(steps, default_table)
        for table_name in needed_tables:
            cursor = table_contexts[table_name]
            if not cursor.has_pending():
                return False
        return True

    def _tables_needed_for_iteration(self, steps: list[dict[str, Any]], default_table: TableCursor | None) -> set[str]:
        names: set[str] = set()
        pattern_explicit = re.compile(r'reg_table:([A-Za-z0-9_]+)\.([A-Za-z0-9_]+)')
        pattern_implicit = re.compile(r'reg_table:([A-Za-z0-9_]+)(?!\.)')
        for step in steps:
            action = str(step.get('acao', ''))
            explicit = pattern_explicit.findall(action)
            if explicit:
                names.update(t for t, _ in explicit)
            elif default_table and pattern_implicit.search(action):
                names.add(default_table.name)
        return names

    def _execute_action_string(
        self,
        step: dict[str, Any],
        action_string: str,
        i: int,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> None:
        if not action_string:
            return
        parts = [part.strip() for part in action_string.split(' + ') if part.strip()]
        for part in parts:
            self._execute_single_action(step, part, i, table_contexts, default_table, acquired_tables)

    def _execute_single_action(
        self,
        step: dict[str, Any],
        action: str,
        i: int,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> None:
        self.log(f'Ação: {action}')

        if action.startswith('repeat:'):
            expr, nested = action[len('repeat:'):].split('|', 1)
            count = int(self._safe_eval(expr.strip(), i))
            for _ in range(max(count, 0)):
                self._execute_single_action(step, nested.strip(), i, table_contexts, default_table, acquired_tables)
            return

        if action.startswith('wait:'):
            time.sleep(float(action.split(':', 1)[1]))
            return

        if action.startswith('press:'):
            key = action.split(':', 1)[1].strip()
            pyautogui.press(key)
            return

        if action.startswith('hotkey:'):
            keys = [k.strip() for k in action.split(':', 1)[1].split('+') if k.strip()]
            pyautogui.hotkey(*keys)
            return

        if action.startswith('type:'):
            # Mantém exatamente tudo o que vier após "type:", inclusive aspas simples ou duplas.
            raw_text = action[len('type:'):]
            text = self._replace_placeholders(raw_text, i, table_contexts, default_table, acquired_tables)
            if pyperclip:
                pyperclip.copy(text)
                pyautogui.hotkey('ctrl', 'v')
            else:
                pyautogui.write(text)
            return

        if action.startswith('reg_table:'):
            text = self._resolve_reg_table(action, table_contexts, default_table, acquired_tables)
            if pyperclip:
                pyperclip.copy(text)
                pyautogui.hotkey('ctrl', 'v')
            else:
                pyautogui.write(text)
            return

        if action.startswith('transform_copy:'):
            func_name = action.split(':', 1)[1].strip()
            if pyperclip is None:
                raise RuntimeError('pyperclip não está instalado. transform_copy requer clipboard.')
            if func_name not in TRANSFORMACOES:
                raise ValueError(f'Função de transformação não encontrada: {func_name}')
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            original = pyperclip.paste()
            transformed = TRANSFORMACOES[func_name](original)
            pyperclip.copy(transformed)
            self.log(f'transform_copy aplicado com {func_name}')
            return

        x = step.get('x')
        y = step.get('y')
        if action == 'move':
            self._require_xy(x, y, action)
            pyautogui.moveTo(int(x), int(y))
            return
        if action == 'click':
            self._require_xy(x, y, action)
            pyautogui.click(int(x), int(y))
            return
        if action == 'double_click':
            self._require_xy(x, y, action)
            pyautogui.doubleClick(int(x), int(y))
            return
        if action == 'right_click':
            self._require_xy(x, y, action)
            pyautogui.rightClick(int(x), int(y))
            return
        if action == 'hold_click':
            self._require_xy(x, y, action)
            pyautogui.mouseDown(int(x), int(y))
            return
        if action == 'release_click':
            self._require_xy(x, y, action)
            pyautogui.mouseUp(int(x), int(y))
            return

        raise ValueError(f'Ação não suportada: {action}')

    def _resolve_reg_table(
        self,
        action: str,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> str:
        payload = action.split(':', 1)[1].strip()
        if '.' in payload:
            table_name, field = payload.split('.', 1)
            cursor = table_contexts.get(table_name)
            if cursor is None:
                raise FileNotFoundError(f'Tabela {table_name} não está carregada.')
        else:
            if default_table is None:
                raise ValueError('reg_table:campo exige uma tabela selecionada no dropdown.')
            cursor = default_table
            field = payload

        if cursor.current_index is None:
            cursor.next_row()
            acquired_tables.add(cursor.name)
            self.log(f'Tabela {cursor.name}: usando novo registro pendente.')
        return cursor.get_current_value(field)

    def _replace_placeholders(
        self,
        text: str,
        i: int,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> str:
        text = text.replace('{i}', str(i))

        def repl(match: re.Match[str]) -> str:
            token = match.group(1)
            if token.startswith('reg_table:'):
                return self._resolve_reg_table(token, table_contexts, default_table, acquired_tables)
            return match.group(0)

        return re.sub(r'\{(reg_table:[^{}]+)\}', repl, text)

    @staticmethod
    def _safe_eval(expr: str, i: int) -> int:
        allowed = {'i': i}
        return int(eval(expr, {'__builtins__': {}}, allowed))

    @staticmethod
    def _require_xy(x: Any, y: Any, action: str) -> None:
        if x is None or y is None:
            raise ValueError(f'Ação {action} exige x e y no passo do roteiro.')


class TableEditorWindow(tk.Toplevel):
    def __init__(self, master: 'AutomationApp') -> None:
        super().__init__(master)
        self.app = master
        self.title('Tabelas')
        self.geometry('900x560')
        self._build()
        self.refresh_list()

    def _build(self) -> None:
        left = ttk.Frame(self)
        left.pack(side='left', fill='y', padx=8, pady=8)
        right = ttk.Frame(self)
        right.pack(side='right', fill='both', expand=True, padx=8, pady=8)

        self.listbox = tk.Listbox(left, width=30)
        self.listbox.pack(fill='y', expand=True)
        self.listbox.bind('<<ListboxSelect>>', lambda _e: self.load_selected())

        btns = ttk.Frame(left)
        btns.pack(fill='x', pady=(8, 0))
        ttk.Button(btns, text='Novo', command=self.new_file).pack(fill='x', pady=2)
        ttk.Button(btns, text='Importar CSV', command=self.import_csv).pack(fill='x', pady=2)
        ttk.Button(btns, text='Salvar', command=self.save_file).pack(fill='x', pady=2)
        ttk.Button(btns, text='Excluir', command=self.delete_file).pack(fill='x', pady=2)
        ttk.Button(btns, text='Atualizar lista', command=self.refresh_list).pack(fill='x', pady=2)

        top = ttk.Frame(right)
        top.pack(fill='x')
        ttk.Label(top, text='Nome do arquivo (sem .csv):').pack(side='left')
        self.name_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.name_var, width=40).pack(side='left', padx=8)

        self.editor = tk.Text(right, wrap='none', undo=True)
        self.editor.pack(fill='both', expand=True, pady=(8, 0))
        self.editor.insert('1.0', 'ID;NOME;STATUS\n1;Exemplo;\n')

    def refresh_list(self) -> None:
        self.listbox.delete(0, tk.END)
        for path in sorted(TABELAS_DIR.glob('*.csv')):
            self.listbox.insert(tk.END, path.stem)

    def load_selected(self) -> None:
        selection = self.listbox.curselection()
        if not selection:
            return
        name = self.listbox.get(selection[0])
        path = TABELAS_DIR / f'{name}.csv'
        self.name_var.set(name)
        self.editor.delete('1.0', tk.END)
        self.editor.insert('1.0', path.read_text(encoding='utf-8'))

    def new_file(self) -> None:
        self.name_var.set('nova_tabela')
        self.editor.delete('1.0', tk.END)
        self.editor.insert('1.0', 'ID;NOME;STATUS\n')

    def import_csv(self) -> None:
        file = filedialog.askopenfilename(filetypes=[('CSV', '*.csv'), ('Todos', '*.*')])
        if not file:
            return
        path = Path(file)
        self.name_var.set(path.stem)
        self.editor.delete('1.0', tk.END)
        self.editor.insert('1.0', path.read_text(encoding='utf-8-sig'))

    def save_file(self) -> None:
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror('Erro', 'Informe o nome do arquivo.')
            return
        path = TABELAS_DIR / f'{name}.csv'
        content = self.editor.get('1.0', tk.END).strip() + '\n'
        path.write_text(content, encoding='utf-8')
        self.app.refresh_table_dropdown()
        self.refresh_list()
        messagebox.showinfo('OK', f'Tabela salva: {path.name}')

    def delete_file(self) -> None:
        name = self.name_var.get().strip()
        if not name:
            return
        path = TABELAS_DIR / f'{name}.csv'
        if path.exists() and messagebox.askyesno('Confirmar', f'Excluir {path.name}?'):
            path.unlink()
            self.app.refresh_table_dropdown()
            self.refresh_list()
            self.new_file()


class ScriptEditorWindow(tk.Toplevel):
    def __init__(self, master: 'AutomationApp') -> None:
        super().__init__(master)
        self.app = master
        self.title('Roteiros')
        self.geometry('980x620')
        self._build()
        self.refresh_list()

    def _build(self) -> None:
        left = ttk.Frame(self)
        left.pack(side='left', fill='y', padx=8, pady=8)
        right = ttk.Frame(self)
        right.pack(side='right', fill='both', expand=True, padx=8, pady=8)

        self.listbox = tk.Listbox(left, width=30)
        self.listbox.pack(fill='y', expand=True)
        self.listbox.bind('<<ListboxSelect>>', lambda _e: self.load_selected())

        btns = ttk.Frame(left)
        btns.pack(fill='x', pady=(8, 0))
        ttk.Button(btns, text='Novo', command=self.new_file).pack(fill='x', pady=2)
        ttk.Button(btns, text='Salvar', command=self.save_file).pack(fill='x', pady=2)
        ttk.Button(btns, text='Excluir', command=self.delete_file).pack(fill='x', pady=2)
        ttk.Button(btns, text='Carregar na tela principal', command=self.load_into_main).pack(fill='x', pady=2)
        ttk.Button(btns, text='Atualizar lista', command=self.refresh_list).pack(fill='x', pady=2)

        top = ttk.Frame(right)
        top.pack(fill='x')
        ttk.Label(top, text='Nome do arquivo (sem .json):').pack(side='left')
        self.name_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.name_var, width=40).pack(side='left', padx=8)

        self.editor = tk.Text(right, wrap='none', undo=True)
        self.editor.pack(fill='both', expand=True, pady=(8, 0))
        self.new_file()

    def _default_script(self) -> str:
        sample = [
            {'obs': 'Espera simples', 'acao': 'wait:0.3'},
            {'obs': 'Seleciona tudo', 'acao': 'hotkey:ctrl+a'},
            {'obs': 'Digita nome a partir da tabela selecionada', 'acao': 'reg_table:NOME'},
        ]
        return json.dumps(sample, indent=2, ensure_ascii=False)

    def refresh_list(self) -> None:
        self.listbox.delete(0, tk.END)
        for path in sorted(ROTEIROS_DIR.glob('*.json')):
            self.listbox.insert(tk.END, path.stem)
        self.app.refresh_script_dropdown()

    def load_selected(self) -> None:
        selection = self.listbox.curselection()
        if not selection:
            return
        name = self.listbox.get(selection[0])
        path = ROTEIROS_DIR / f'{name}.json'
        self.name_var.set(name)
        self.editor.delete('1.0', tk.END)
        self.editor.insert('1.0', path.read_text(encoding='utf-8'))

    def new_file(self) -> None:
        self.name_var.set('novo_roteiro')
        self.editor.delete('1.0', tk.END)
        self.editor.insert('1.0', self._default_script())

    def save_file(self) -> None:
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror('Erro', 'Informe o nome do arquivo.')
            return
        content = self.editor.get('1.0', tk.END).strip()
        try:
            parsed = json.loads(content)
            if not isinstance(parsed, list):
                raise ValueError('O roteiro deve ser uma lista JSON.')
        except Exception as exc:
            messagebox.showerror('Erro de JSON', str(exc))
            return
        path = ROTEIROS_DIR / f'{name}.json'
        path.write_text(json.dumps(parsed, indent=2, ensure_ascii=False), encoding='utf-8')
        self.app.refresh_script_dropdown()
        self.refresh_list()
        messagebox.showinfo('OK', f'Roteiro salvo: {path.name}')

    def delete_file(self) -> None:
        name = self.name_var.get().strip()
        if not name:
            return
        path = ROTEIROS_DIR / f'{name}.json'
        if path.exists() and messagebox.askyesno('Confirmar', f'Excluir {path.name}?'):
            path.unlink()
            self.app.refresh_script_dropdown()
            self.refresh_list()
            self.new_file()

    def load_into_main(self) -> None:
        self.app.script_text.delete('1.0', tk.END)
        self.app.script_text.insert('1.0', self.editor.get('1.0', tk.END))
        self.app.script_name_var.set(self.name_var.get().strip())
        messagebox.showinfo('OK', 'Roteiro carregado na tela principal.')


class AutomationApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title('Automação de Teclado e Mouse')
        self.geometry('1220x780')

        self.script_name_var = tk.StringVar()
        self.selected_script_var = tk.StringVar()
        self.selected_table_var = tk.StringVar()
        self.start_delay_var = tk.StringVar(value='0')
        self.delay_var = tk.StringVar(value='0.3')
        self.repetitions_var = tk.StringVar(value='1')
        self.mouse_position_var = tk.StringVar(value='Mouse: X=0 Y=0')

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.stop_event = threading.Event()
        self.worker_thread: threading.Thread | None = None

        self._build_ui()
        self.refresh_script_dropdown()
        self.refresh_table_dropdown()
        self.after(150, self.process_log_queue)
        self.after(100, self.update_mouse_position)

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=10)
        top.pack(fill='x')

        ttk.Label(top, text='Roteiro salvo:').grid(row=0, column=0, sticky='w', padx=4, pady=4)
        self.script_combo = ttk.Combobox(top, textvariable=self.selected_script_var, state='readonly', width=30)
        self.script_combo.grid(row=0, column=1, sticky='we', padx=4, pady=4)
        ttk.Button(top, text='Carregar', command=self.load_selected_script).grid(row=0, column=2, padx=4, pady=4)
        ttk.Button(top, text='Salvar como roteiro', command=self.save_script_from_main).grid(row=0, column=3, padx=4, pady=4)
        ttk.Button(top, text='Roteiros', command=self.open_script_editor).grid(row=0, column=4, padx=4, pady=4)
        ttk.Button(top, text='Tabelas', command=self.open_table_editor).grid(row=0, column=5, padx=4, pady=4)
        ttk.Button(top, text='Help', command=self.open_help).grid(row=0, column=6, padx=4, pady=4)

        ttk.Label(top, text='Nome do roteiro atual:').grid(row=1, column=0, sticky='w', padx=4, pady=4)
        ttk.Entry(top, textvariable=self.script_name_var, width=32).grid(row=1, column=1, sticky='we', padx=4, pady=4)

        ttk.Label(top, text='Tabela:').grid(row=1, column=2, sticky='e', padx=4, pady=4)
        self.table_combo = ttk.Combobox(top, textvariable=self.selected_table_var, state='readonly', width=28)
        self.table_combo.grid(row=1, column=3, sticky='we', padx=4, pady=4)

        ttk.Label(top, text='Delay inicial (s):').grid(row=1, column=4, sticky='e', padx=4, pady=4)
        ttk.Entry(top, textvariable=self.start_delay_var, width=10).grid(row=1, column=5, sticky='w', padx=4, pady=4)

        ttk.Label(top, text='Delay entre passos (s):').grid(row=1, column=6, sticky='e', padx=4, pady=4)
        ttk.Entry(top, textvariable=self.delay_var, width=10).grid(row=1, column=7, sticky='w', padx=4, pady=4)

        ttk.Label(top, text='Repetições:').grid(row=1, column=8, sticky='e', padx=4, pady=4)
        ttk.Entry(top, textvariable=self.repetitions_var, width=10).grid(row=1, column=9, sticky='w', padx=4, pady=4)

        ttk.Label(top, textvariable=self.mouse_position_var).grid(row=2, column=0, columnspan=10, sticky='w', padx=4, pady=(2, 4))

        for c in range(10):
            top.columnconfigure(c, weight=1)

        center = ttk.Panedwindow(self, orient='vertical')
        center.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        script_frame = ttk.Labelframe(center, text='Roteiro para execução')
        console_frame = ttk.Labelframe(center, text='Console')
        center.add(script_frame, weight=3)
        center.add(console_frame, weight=2)

        self.script_text = tk.Text(script_frame, wrap='none', undo=True, height=20)
        self.script_text.pack(fill='both', expand=True, padx=6, pady=6)
        self.script_text.insert('1.0', json.dumps([
            {'obs': 'Exemplo', 'acao': 'wait:0.3'},
            {'obs': 'Exemplo com tabela selecionada', 'acao': 'reg_table:NOME'},
        ], indent=2, ensure_ascii=False))

        controls = ttk.Frame(self, padding=(10, 0, 10, 10))
        controls.pack(fill='x')
        ttk.Button(controls, text='Executar', command=self.start_execution).pack(side='left', padx=4)
        ttk.Button(controls, text='Parar', command=self.stop_execution).pack(side='left', padx=4)
        ttk.Button(controls, text='Limpar console', command=self.clear_console).pack(side='left', padx=4)

        self.console_text = tk.Text(console_frame, wrap='word', state='disabled', height=12)
        self.console_text.pack(fill='both', expand=True, padx=6, pady=6)

    def refresh_script_dropdown(self) -> None:
        values = [''] + [path.stem for path in sorted(ROTEIROS_DIR.glob('*.json'))]
        self.script_combo['values'] = values
        if self.selected_script_var.get() not in values:
            self.selected_script_var.set('')

    def refresh_table_dropdown(self) -> None:
        values = [''] + [path.stem for path in sorted(TABELAS_DIR.glob('*.csv'))]
        self.table_combo['values'] = values
        if self.selected_table_var.get() not in values:
            self.selected_table_var.set('')

    def open_help(self) -> None:
        win = tk.Toplevel(self)
        win.title('Help - Comandos do Roteiro')
        win.geometry('900x680')
        txt = tk.Text(win, wrap='word')
        txt.pack(fill='both', expand=True)
        txt.insert('1.0', HELP_TEXT)
        txt.config(state='disabled')

    def open_table_editor(self) -> None:
        TableEditorWindow(self)

    def open_script_editor(self) -> None:
        ScriptEditorWindow(self)

    def load_selected_script(self) -> None:
        name = self.selected_script_var.get().strip()
        if not name:
            return
        path = ROTEIROS_DIR / f'{name}.json'
        if not path.exists():
            messagebox.showerror('Erro', 'Roteiro não encontrado.')
            return
        self.script_name_var.set(name)
        self.script_text.delete('1.0', tk.END)
        self.script_text.insert('1.0', path.read_text(encoding='utf-8'))
        self.log(f'Roteiro carregado: {name}')

    def save_script_from_main(self) -> None:
        name = self.script_name_var.get().strip() or 'roteiro_principal'
        content = self.script_text.get('1.0', tk.END).strip()
        try:
            parsed = json.loads(content)
            if not isinstance(parsed, list):
                raise ValueError('O roteiro deve ser uma lista JSON.')
        except Exception as exc:
            messagebox.showerror('Erro de JSON', str(exc))
            return
        path = ROTEIROS_DIR / f'{name}.json'
        path.write_text(json.dumps(parsed, indent=2, ensure_ascii=False), encoding='utf-8')
        self.refresh_script_dropdown()
        self.selected_script_var.set(name)
        self.log(f'Roteiro salvo: {path.name}')
        messagebox.showinfo('OK', f'Roteiro salvo: {path.name}')

    def start_execution(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning('Aviso', 'Já existe uma execução em andamento.')
            return

        try:
            steps = json.loads(self.script_text.get('1.0', tk.END).strip())
            if not isinstance(steps, list):
                raise ValueError('O roteiro deve ser uma lista JSON.')
            repetitions = int(self.repetitions_var.get())
            float(self.start_delay_var.get())
            float(self.delay_var.get())
        except Exception as exc:
            messagebox.showerror('Erro', str(exc))
            return

        self.stop_event.clear()
        selected_table = self.selected_table_var.get().strip() or None
        runner = ScriptRunner(
            self.log,
            self.stop_event.is_set,
            lambda: float(self.delay_var.get() or 0),
            lambda: float(self.start_delay_var.get() or 0),
        )

        def target() -> None:
            try:
                self.log('Execução iniciada.')
                runner.run(steps, repetitions, selected_table)
            except Exception as exc:
                self.log(f'ERRO: {exc}')
            finally:
                self.log('Thread finalizada.')

        self.worker_thread = threading.Thread(target=target, daemon=True)
        self.worker_thread.start()

    def stop_execution(self) -> None:
        self.stop_event.set()
        self.log('Solicitação de parada registrada.')

    def clear_console(self) -> None:
        self.console_text.config(state='normal')
        self.console_text.delete('1.0', tk.END)
        self.console_text.config(state='disabled')

    def log(self, message: str) -> None:
        timestamp = time.strftime('%H:%M:%S')
        self.log_queue.put(f'[{timestamp}] {message}')

    def process_log_queue(self) -> None:
        while not self.log_queue.empty():
            message = self.log_queue.get()
            self.console_text.config(state='normal')
            self.console_text.insert(tk.END, message + '\n')
            self.console_text.see(tk.END)
            self.console_text.config(state='disabled')
        self.after(150, self.process_log_queue)

    def update_mouse_position(self) -> None:
        try:
            x, y = pyautogui.position()
            self.mouse_position_var.set(f'Mouse: X={x} Y={y}')
        except Exception:
            self.mouse_position_var.set('Mouse: X=? Y=?')
        self.after(100, self.update_mouse_position)


if __name__ == '__main__':
    app = AutomationApp()
    app.mainloop()