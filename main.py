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
APP_DIR = BASE_DIR / "app"
ROTEIROS_DIR = BASE_DIR / "pyBossAll\\roteiros"
TABELAS_DIR = BASE_DIR / "pyBossAll\\tabelas"

for pasta in (APP_DIR, ROTEIROS_DIR, TABELAS_DIR):
    pasta.mkdir(parents=True, exist_ok=True)

pyautogui.FAILSAFE = True


HELP_TEXT = """COMANDOS DISPONÍVEIS NO ROTEIRO (JSON)

============================================================
1) ESTRUTURA GERAL
============================================================

O roteiro deve ser uma LISTA JSON.

Exemplo:
[
  {
    "info": "Mover mouse",
    "esperar": {"antes": 0.3, "depois": 0.2},
    "mouse": {"x": 250, "y": 620, "acao": "mover"}
  },
  {
    "info": "Clique esquerdo",
    "mouse": {"x": 250, "y": 620, "acao": "clicar_esquerdo"}
  },
  {
    "info": "Digita texto fixo",
    "teclado": {"digitar": "Teste"}
  }
]

============================================================
2) ITENS POSSÍVEIS EM CADA OBJETO DO ROTEIRO
============================================================

Cada item pode ter:
- "info"
- "esperar"
- "mouse"
- "teclado"
- "repetir"

Regras:
- Cada item deve ter somente 1 ação principal:
  OU "mouse"
  OU "teclado"
- Não use "mouse" e "teclado" juntos no mesmo item.
- "repetir" repete o item atual.
- "esperar" pode ser usado antes e/ou depois do item.
- "info" é apenas descritivo e aparece no console.

============================================================
3) DETALHE DE CADA ITEM
============================================================

3.1) info
------------------------------------------------------------

Formato:
"info": "Texto informacional"

Exemplo:
{
  "info": "Abrir busca",
  "teclado": {"atalho": "ctrl+f"}
}

============================================================
3.2) esperar
------------------------------------------------------------

Formato:
"esperar": {
  "antes": 1,
  "depois": 1
}

Subitens possíveis:
- antes  -> tempo em segundos antes da ação
- depois -> tempo em segundos depois da ação

Exemplos:
{
  "info": "Aguardar antes",
  "esperar": {"antes": 1},
  "teclado": {"pressionar": "enter"}
}

{
  "info": "Aguardar depois",
  "esperar": {"depois": 1.5},
  "teclado": {"pressionar": "tab"}
}

{
  "info": "Aguardar antes e depois",
  "esperar": {"antes": 0.5, "depois": 0.5},
  "mouse": {"x": 100, "y": 200, "acao": "clicar_esquerdo"}
}

============================================================
3.3) repetir
------------------------------------------------------------

Formato:
"repetir": 3

ou

"repetir": "3"

ou expressão com i:
"repetir": "i+1"

Observação:
- i é o índice da iteração externa, começando em 0.
- Se estiver na 1ª iteração, i = 0.
- Se estiver na 2ª iteração, i = 1.

Exemplos:
{
  "info": "Pressionar TAB 3 vezes",
  "repetir": 3,
  "teclado": {"pressionar": "tab"}
}

{
  "info": "Repete conforme iteração externa",
  "repetir": "i+1",
  "teclado": {"pressionar": "down"}
}

============================================================
4) ITEM MOUSE
============================================================

Formato:
"mouse": {
  "x": 150,
  "y": 150,
  "acao": "clicar_direito"
}

Subitens possíveis:
- x
- y
- acao

============================================================
4.1) mouse.x
------------------------------------------------------------

Coordenada X da tela.

Formato:
"x": 150

Também pode ser texto com placeholder:
"x": "{i}"

============================================================
4.2) mouse.y
------------------------------------------------------------

Coordenada Y da tela.

Formato:
"y": 150

Também pode ser texto com placeholder:
"y": "{i}"

============================================================
4.3) mouse.acao
------------------------------------------------------------

Valores permitidos:
- "mover"
- "clicar_esquerdo"
- "clicar_direito"
- "clicar_duplo"
- "clicar_segurar"
- "clicar_soltar"

Exemplos:
{
  "info": "Mover mouse",
  "mouse": {"x": 100, "y": 200, "acao": "mover"}
}

{
  "info": "Clique esquerdo",
  "mouse": {"x": 100, "y": 200, "acao": "clicar_esquerdo"}
}

{
  "info": "Clique direito",
  "mouse": {"x": 100, "y": 200, "acao": "clicar_direito"}
}

{
  "info": "Duplo clique",
  "mouse": {"x": 100, "y": 200, "acao": "clicar_duplo"}
}

{
  "info": "Segurar clique",
  "mouse": {"x": 100, "y": 200, "acao": "clicar_segurar"}
}

{
  "info": "Soltar clique",
  "mouse": {"x": 100, "y": 200, "acao": "clicar_soltar"}
}

Observação:
- Todas as ações de mouse exigem x e y.

============================================================
5) ITEM TECLADO
============================================================

Formato:
"teclado": {
  "digitar": "texto"
}

O objeto "teclado" deve ter UMA chave principal por item.

Subitens possíveis:
- "digitar"
- "atalho"
- "pressionar"
- "campo_tabela"
- "funcao_py"

============================================================
5.1) teclado.digitar
------------------------------------------------------------

Digita ou cola um texto no campo atual.

Formato:
"teclado": {"digitar": "meu texto"}

Exemplo:
{
  "info": "Digita texto fixo",
  "teclado": {"digitar": "ZS4_VCI_A003"}
}

Também aceita placeholder:
{
  "info": "Digita índice da iteração",
  "teclado": {"digitar": "Item_{i}"}
}

============================================================
5.2) teclado.atalho
------------------------------------------------------------

Executa hotkey com pyautogui.hotkey.

Formato:
"teclado": {"atalho": "ctrl+s"}

Exemplos:
{
  "info": "Salvar",
  "teclado": {"atalho": "ctrl+s"}
}

{
  "info": "Localizar",
  "teclado": {"atalho": "ctrl+f"}
}

{
  "info": "Selecionar tudo",
  "teclado": {"atalho": "ctrl+a"}
}

============================================================
5.3) teclado.pressionar
------------------------------------------------------------

Pressiona uma tecla específica com pyautogui.press.

Formato:
"teclado": {"pressionar": "enter"}

Teclas comuns:
- enter
- tab
- up
- down
- left
- right
- esc
- backspace
- delete
- home
- end
- pageup
- pagedown
- space

Exemplos:
{
  "info": "Confirmar",
  "teclado": {"pressionar": "enter"}
}

{
  "info": "Próximo campo",
  "teclado": {"pressionar": "tab"}
}

{
  "info": "Seta para cima",
  "teclado": {"pressionar": "up"}
}

{
  "info": "Seta para baixo",
  "teclado": {"pressionar": "down"}
}

{
  "info": "Seta para direita",
  "teclado": {"pressionar": "right"}
}

{
  "info": "Seta para esquerda",
  "teclado": {"pressionar": "left"}
}

============================================================
5.4) teclado.campo_tabela
------------------------------------------------------------

Busca o próximo registro pendente em uma tabela CSV e usa o valor do campo informado.

Formatos:
a) usando a tabela selecionada no dropdown:
"teclado": {"campo_tabela": "NOME"}

b) usando tabela específica:
"teclado": {"campo_tabela": "clientes.EMAIL"}

Regras:
1. Sem ponto:
   - usa a tabela selecionada no dropdown principal
   - exemplo: "NOME"

2. Com ponto:
   - primeiro nome = nome da tabela sem .csv
   - segundo nome = coluna
   - exemplo: "clientes.EMAIL"

3. Toda vez que um campo_tabela for usado:
   - o app captura o próximo registro pendente daquela tabela
   - o registro fica "reservado" na iteração atual
   - ao final da iteração bem-sucedida, STATUS = OK
   - registros com STATUS=OK não são reutilizados

4. O app pode usar mais de uma tabela no mesmo roteiro.

Exemplos:
{
  "info": "Usa nome da tabela padrão",
  "teclado": {"campo_tabela": "NOME"}
}

{
  "info": "Usa email da tabela clientes.csv",
  "teclado": {"campo_tabela": "clientes.EMAIL"}
}

============================================================
5.5) teclado.funcao_py
------------------------------------------------------------

Usa uma função do arquivo transformacoes.py.

Formato:
"teclado": {"funcao_py": "NOME_DA_FUNCAO"}

Comportamento:
- copia o texto selecionado
- envia esse texto como input da função
- recebe o retorno
- cola o retorno no campo selecionado

Exemplo:
{
  "info": "Transformar texto selecionado",
  "teclado": {"funcao_py": "trocar_prefixo_tabela_rf"}
}

Requisitos:
- pyperclip instalado
- função existente em transformacoes.py
- função registrada em TRANSFORMACOES

============================================================
6) PLACEHOLDERS
============================================================

Suportado:
- {i}

Exemplo:
{
  "info": "Digita índice",
  "teclado": {"digitar": "Registro_{i}"}
}

============================================================
7) EXEMPLO COMPLETO
============================================================

[
  {
    "info": "Mover até o campo",
    "esperar": {"antes": 0.3, "depois": 0.2},
    "mouse": {"x": 250, "y": 620, "acao": "mover"}
  },
  {
    "info": "Clicar no campo",
    "mouse": {"x": 250, "y": 620, "acao": "clicar_esquerdo"}
  },
  {
    "info": "Digitar valor da tabela padrão",
    "teclado": {"campo_tabela": "NOME"}
  },
  {
    "info": "Pressionar TAB 2 vezes",
    "repetir": 2,
    "teclado": {"pressionar": "tab"}
  },
  {
    "info": "Salvar",
    "teclado": {"atalho": "ctrl+s"}
  }
]

============================================================
8) REGRAS DAS TABELAS CSV
============================================================

- Separador preferencial: ;
- Deve possuir cabeçalho
- Se STATUS não existir, o app cria
- STATUS=OK significa registro já consumido
- Registros com STATUS=OK não são reutilizados
- O OK é gravado ao final da iteração executada com sucesso
- Pode haver uso de múltiplas tabelas no mesmo roteiro

============================================================
9) CONSOLE
============================================================

- O console é limpo automaticamente antes de cada execução
- A linha de iteração aparece em negrito:
  --- Iteração X/Y ---
- Cada item do JSON é logado em uma única linha
- Ao final da iteração o tempo total aparece em vermelho

============================================================
10) SEGURANÇA
============================================================

- O FAILSAFE do pyautogui está ativo
- Mova o mouse para o canto superior esquerdo para interromper por FAILSAFE
"""


@dataclass
class TableCursor:
    name: str
    path: Path
    rows: list[dict[str, str]]
    fieldnames: list[str]
    pending_indexes: list[int]
    current_index: int | None = None

    @classmethod
    def load(cls, table_name: str) -> "TableCursor":
        path = TABELAS_DIR / f"{table_name}.csv"
        if not path.exists():
            raise FileNotFoundError(f"Tabela não encontrada: {path.name}")

        with path.open("r", encoding="utf-8-sig", newline="") as f:
            sample = f.read(2048)
            f.seek(0)
            delimiter = ";" if sample.count(";") >= sample.count(",") else ","
            reader = csv.DictReader(f, delimiter=delimiter)
            rows = [dict(r) for r in reader]
            fieldnames = list(reader.fieldnames or [])

        if "STATUS" not in fieldnames:
            fieldnames.append("STATUS")
            for row in rows:
                row["STATUS"] = row.get("STATUS", "")
        else:
            for row in rows:
                row["STATUS"] = (row.get("STATUS") or "").strip()

        pending_indexes = [
            idx for idx, row in enumerate(rows)
            if (row.get("STATUS") or "").strip().upper() != "OK"
        ]
        return cls(table_name, path, rows, fieldnames, pending_indexes)

    def has_pending(self) -> bool:
        return len(self.pending_indexes) > 0

    def next_row(self) -> dict[str, str]:
        if not self.pending_indexes:
            raise RuntimeError(f"Tabela {self.name} sem registros pendentes.")
        self.current_index = self.pending_indexes.pop(0)
        return self.rows[self.current_index]

    def get_current_value(self, field: str) -> str:
        if self.current_index is None:
            self.next_row()
        assert self.current_index is not None
        row = self.rows[self.current_index]
        if field not in row:
            raise KeyError(f"Campo {field} não existe na tabela {self.name}.")
        return row.get(field, "") or ""

    def mark_current_ok(self) -> None:
        if self.current_index is None:
            return
        self.rows[self.current_index]["STATUS"] = "OK"

    def reset_current(self) -> None:
        if self.current_index is not None:
            self.pending_indexes.insert(0, self.current_index)
            self.current_index = None

    def save(self) -> None:
        with self.path.open("w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=self.fieldnames, delimiter=";")
            writer.writeheader()
            writer.writerows(self.rows)


class ScriptRunner:
    def __init__(
        self,
        log_callback: Callable[[str, str], None],
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
            self.log(
                f"Tabela selecionada: {selected_table} | pendentes: {len(default_table.pending_indexes)}",
                "info",
            )

        explicit_tables = self._extract_explicit_table_names(steps)
        for name in explicit_tables:
            if name not in table_contexts:
                table_contexts[name] = TableCursor.load(name)
                self.log(
                    f"Tabela referenciada no roteiro: {name} | pendentes: {len(table_contexts[name].pending_indexes)}",
                    "info",
                )

        uses_table = bool(default_table or explicit_tables)

        start_delay = self.get_start_delay()
        if start_delay > 0:
            self.log(
                f"Delay inicial: aguardando {start_delay} segundo(s) antes de iniciar.",
                "info",
            )
            time.sleep(start_delay)

        for i in range(repetitions):
            if self.is_stop_requested():
                self.log("Execução interrompida pelo usuário.", "warning")
                return

            if uses_table and not self._has_rows_available(steps, table_contexts, default_table):
                self.log(
                    "Sem registros pendentes suficientes para continuar. Execução finalizada.",
                    "warning",
                )
                return

            iteration_start = time.perf_counter()
            self.log(f"--- Iteração {i + 1}/{repetitions} ---", "iteration")
            acquired_tables: set[str] = set()

            try:
                for step_idx, step in enumerate(steps, start=1):
                    if self.is_stop_requested():
                        self.log("Execução interrompida pelo usuário.", "warning")
                        return

                    self._execute_step(
                        step=step,
                        step_idx=step_idx,
                        i=i,
                        table_contexts=table_contexts,
                        default_table=default_table,
                        acquired_tables=acquired_tables,
                    )

                    delay = self.get_delay()
                    if delay > 0 and step_idx < len(steps):
                        time.sleep(delay)

                for table_name in acquired_tables:
                    table_contexts[table_name].mark_current_ok()
                    table_contexts[table_name].save()
                    table_contexts[table_name].current_index = None
                    self.log(f"Tabela {table_name}: registro marcado com STATUS=OK", "success")

            except Exception as exc:
                for table_name in acquired_tables:
                    table_contexts[table_name].reset_current()
                raise RuntimeError(f"Erro na iteração {i + 1}: {exc}") from exc
            finally:
                elapsed = time.perf_counter() - iteration_start
                self.log(f"Tempo total da iteração: {elapsed:.2f}s", "elapsed")

        self.log("Execução concluída.", "success")

    def _execute_step(
        self,
        step: dict[str, Any],
        step_idx: int,
        i: int,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> None:
        info = str(step.get("info") or step.get("obs") or f"Item {step_idx}")
        before_wait, after_wait = self._parse_wait(step.get("esperar"))
        repeat_count = max(self._parse_repeat(step.get("repetir"), i), 1)
        action_kind, action_value = self._get_action_payload(step)

        summary = self._build_step_summary(
            info=info,
            action_kind=action_kind,
            action_value=action_value,
            before_wait=before_wait,
            after_wait=after_wait,
            repeat_count=repeat_count,
        )
        self.log(summary, "step")

        for _ in range(repeat_count):
            if before_wait > 0:
                time.sleep(before_wait)

            if action_kind == "mouse":
                self._execute_mouse(
                    payload=action_value,
                    i=i,
                    table_contexts=table_contexts,
                    default_table=default_table,
                    acquired_tables=acquired_tables,
                )
            elif action_kind == "teclado":
                self._execute_keyboard(
                    payload=action_value,
                    i=i,
                    table_contexts=table_contexts,
                    default_table=default_table,
                    acquired_tables=acquired_tables,
                )

            if after_wait > 0:
                time.sleep(after_wait)

    def _get_action_payload(self, step: dict[str, Any]) -> tuple[str | None, Any | None]:
        has_mouse = isinstance(step.get("mouse"), dict) and bool(step.get("mouse"))
        has_keyboard = isinstance(step.get("teclado"), dict) and bool(step.get("teclado"))

        if has_mouse and has_keyboard:
            raise ValueError("Cada item pode ter somente 1 ação principal: mouse OU teclado.")
        if has_mouse:
            return "mouse", step.get("mouse")
        if has_keyboard:
            return "teclado", step.get("teclado")
        return None, None

    def _build_step_summary(
        self,
        info: str,
        action_kind: str | None,
        action_value: Any | None,
        before_wait: float,
        after_wait: float,
        repeat_count: int,
    ) -> str:
        parts = [f"INFO={info}"]

        if before_wait > 0:
            parts.append(f"ANTES={before_wait}s")

        if action_kind == "mouse" and isinstance(action_value, dict):
            x = action_value.get("x")
            y = action_value.get("y")
            acao = action_value.get("acao")
            parts.append(f"MOUSE=x:{x}, y:{y}, acao:{acao}")

        elif action_kind == "teclado" and isinstance(action_value, dict):
            key, value = next(iter(action_value.items()))
            parts.append(f"TECLADO={key}:{value}")

        if repeat_count > 1:
            parts.append(f"REPETIR={repeat_count}")

        if after_wait > 0:
            parts.append(f"DEPOIS={after_wait}s")

        return " | ".join(parts)

    def _extract_explicit_table_names(self, steps: list[dict[str, Any]]) -> set[str]:
        names: set[str] = set()
        for step in steps:
            teclado = step.get("teclado")
            if not isinstance(teclado, dict):
                continue
            campo_tabela = teclado.get("campo_tabela")
            if not campo_tabela:
                continue
            raw = str(campo_tabela).strip()
            if "." in raw:
                table_name, _field = raw.split(".", 1)
                names.add(table_name.strip())
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

    def _tables_needed_for_iteration(
        self,
        steps: list[dict[str, Any]],
        default_table: TableCursor | None,
    ) -> set[str]:
        names: set[str] = set()

        for step in steps:
            teclado = step.get("teclado")
            if not isinstance(teclado, dict):
                continue

            campo_tabela = teclado.get("campo_tabela")
            if not campo_tabela:
                continue

            raw = str(campo_tabela).strip()
            if "." in raw:
                table_name, _field = raw.split(".", 1)
                names.add(table_name.strip())
            elif default_table:
                names.add(default_table.name)

        return names

    def _parse_wait(self, wait_value: Any) -> tuple[float, float]:
        before = 0.0
        after = 0.0

        if wait_value in (None, ""):
            return before, after

        if isinstance(wait_value, dict):
            if "antes" in wait_value:
                before = float(wait_value["antes"])
            if "depois" in wait_value:
                after = float(wait_value["depois"])
            return before, after

        raise ValueError('Campo "esperar" inválido. Use {"antes": n, "depois": n}.')

    def _parse_repeat(self, repeat_value: Any, i: int) -> int:
        if repeat_value in (None, ""):
            return 1
        if isinstance(repeat_value, int):
            return repeat_value
        return int(self._safe_eval(str(repeat_value).strip(), i))

    def _execute_mouse(
        self,
        payload: dict[str, Any],
        i: int,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> None:
        x_raw = payload.get("x")
        y_raw = payload.get("y")
        action = str(payload.get("acao", "")).strip().lower()

        if action == "":
            raise ValueError('mouse.acao é obrigatório.')

        x = self._parse_xy_value(x_raw, i, table_contexts, default_table, acquired_tables)
        y = self._parse_xy_value(y_raw, i, table_contexts, default_table, acquired_tables)

        self._require_xy(x, y, action)

        if action == "mover":
            pyautogui.moveTo(x, y)
        elif action == "clicar_esquerdo":
            pyautogui.click(x, y, button="left")
        elif action == "clicar_direito":
            pyautogui.click(x, y, button="right")
        elif action == "clicar_duplo":
            pyautogui.doubleClick(x, y, button="left")
        elif action == "clicar_segurar":
            pyautogui.mouseDown(x, y, button="left")
        elif action == "clicar_soltar":
            pyautogui.mouseUp(x, y, button="left")
        else:
            raise ValueError(f"Ação de mouse inválida: {action}")

    def _execute_keyboard(
        self,
        payload: dict[str, Any],
        i: int,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> None:
        if len(payload) != 1:
            raise ValueError("O objeto teclado deve conter somente 1 subitem por passo.")

        key, value = next(iter(payload.items()))
        key = str(key).strip()

        if key == "digitar":
            text = self._replace_placeholders(str(value), i, table_contexts, default_table, acquired_tables)
            self._write_text(text)

        elif key == "atalho":
            hotkeys = [part.strip().lower() for part in str(value).split("+") if part.strip()]
            if not hotkeys:
                raise ValueError("Atalho inválido.")
            pyautogui.hotkey(*hotkeys)

        elif key == "pressionar":
            pyautogui.press(str(value).strip().lower())

        elif key == "campo_tabela":
            text = self._resolve_field_reference(
                reference=str(value).strip(),
                table_contexts=table_contexts,
                default_table=default_table,
                acquired_tables=acquired_tables,
            )
            self._write_text(text)

        elif key == "funcao_py":
            self._execute_transform_function(str(value).strip())

        else:
            raise ValueError(
                f"Subitem de teclado inválido: {key}. "
                "Use digitar, atalho, pressionar, campo_tabela ou funcao_py."
            )

    def _execute_transform_function(self, function_name: str) -> None:
        if pyperclip is None:
            raise RuntimeError("pyperclip não está instalado. Instale para usar funcao_py.")

        func = TRANSFORMACOES.get(function_name)
        if func is None:
            raise KeyError(f"Função não encontrada em TRANSFORMACOES: {function_name}")

        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.15)
        original = pyperclip.paste()
        transformed = func(original)
        pyperclip.copy(str(transformed))
        pyautogui.hotkey("ctrl", "v")

    def _write_text(self, text: str) -> None:
        if pyperclip is not None:
            pyperclip.copy(text)
            pyautogui.hotkey("ctrl", "v")
        else:
            pyautogui.write(text, interval=0)

    def _parse_xy_value(
        self,
        value: Any,
        i: int,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> int:
        if value is None:
            raise ValueError("x e y são obrigatórios para ações de mouse.")

        if isinstance(value, (int, float)):
            return int(value)

        text = self._replace_placeholders(str(value), i, table_contexts, default_table, acquired_tables)
        return int(float(text))

    def _resolve_field_reference(
        self,
        reference: str,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> str:
        raw = reference.strip()

        if "." in raw:
            table_name, field = raw.split(".", 1)
            table_name = table_name.strip()
            field = field.strip()
            if table_name not in table_contexts:
                table_contexts[table_name] = TableCursor.load(table_name)
            cursor = table_contexts[table_name]
        else:
            if default_table is None:
                raise RuntimeError(
                    "campo_tabela sem nome de tabela exige uma tabela selecionada no dropdown."
                )
            field = raw
            cursor = default_table

        if cursor.current_index is None:
            cursor.next_row()
            acquired_tables.add(cursor.name)
            self.log(f"Tabela {cursor.name}: usando novo registro pendente.", "info")

        return cursor.get_current_value(field)

    def _replace_placeholders(
        self,
        text: str,
        i: int,
        table_contexts: dict[str, TableCursor],
        default_table: TableCursor | None,
        acquired_tables: set[str],
    ) -> str:
        text = text.replace("{i}", str(i))
        return text

    @staticmethod
    def _safe_eval(expr: str, i: int) -> int:
        allowed = {"i": i}
        return int(eval(expr, {"__builtins__": {}}, allowed))

    @staticmethod
    def _require_xy(x: Any, y: Any, action: str) -> None:
        if x is None or y is None:
            raise ValueError(f'Ação "{action}" exige x e y no passo do roteiro.')


class TableEditorWindow(tk.Toplevel):
    def __init__(self, master: "AutomationApp") -> None:
        super().__init__(master)
        self.app = master
        self.title("CRUD de Tabelas (CSV)")
        self.geometry("900x560")
        self._build()
        self.refresh_list()

    def _build(self) -> None:
        left = ttk.Frame(self)
        left.pack(side="left", fill="y", padx=8, pady=8)
        right = ttk.Frame(self)
        right.pack(side="right", fill="both", expand=True, padx=8, pady=8)

        self.listbox = tk.Listbox(left, width=30)
        self.listbox.pack(fill="y", expand=True)
        self.listbox.bind("<<ListboxSelect>>", lambda _e: self.load_selected())

        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Novo", command=self.new_file).pack(fill="x", pady=2)
        ttk.Button(btns, text="Importar CSV", command=self.import_csv).pack(fill="x", pady=2)
        ttk.Button(btns, text="Salvar", command=self.save_file).pack(fill="x", pady=2)
        ttk.Button(btns, text="Excluir", command=self.delete_file).pack(fill="x", pady=2)
        ttk.Button(btns, text="Atualizar lista", command=self.refresh_list).pack(fill="x", pady=2)

        top = ttk.Frame(right)
        top.pack(fill="x")
        ttk.Label(top, text="Nome do arquivo (sem .csv):").pack(side="left")
        self.name_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.name_var, width=40).pack(side="left", padx=8)

        self.editor = tk.Text(right, wrap="none", undo=True)
        self.editor.pack(fill="both", expand=True, pady=(8, 0))
        self.editor.insert("1.0", "ID;NOME;STATUS\n1;Exemplo;\n")

    def refresh_list(self) -> None:
        self.listbox.delete(0, tk.END)
        for path in sorted(TABELAS_DIR.glob("*.csv")):
            self.listbox.insert(tk.END, path.stem)

    def load_selected(self) -> None:
        selection = self.listbox.curselection()
        if not selection:
            return
        name = self.listbox.get(selection[0])
        path = TABELAS_DIR / f"{name}.csv"
        self.name_var.set(name)
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", path.read_text(encoding="utf-8"))

    def new_file(self) -> None:
        self.name_var.set("nova_tabela")
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", "ID;NOME;STATUS\n")

    def import_csv(self) -> None:
        file = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("Todos", "*.*")])
        if not file:
            return
        path = Path(file)
        self.name_var.set(path.stem)
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", path.read_text(encoding="utf-8-sig"))

    def save_file(self) -> None:
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Erro", "Informe o nome do arquivo.")
            return
        path = TABELAS_DIR / f"{name}.csv"
        content = self.editor.get("1.0", tk.END).strip() + "\n"
        path.write_text(content, encoding="utf-8")
        self.app.refresh_table_dropdown()
        self.refresh_list()
        messagebox.showinfo("OK", f"Tabela salva: {path.name}")

    def delete_file(self) -> None:
        name = self.name_var.get().strip()
        if not name:
            return
        path = TABELAS_DIR / f"{name}.csv"
        if path.exists() and messagebox.askyesno("Confirmar", f"Excluir {path.name}?"):
            path.unlink()
            self.app.refresh_table_dropdown()
            self.refresh_list()
            self.new_file()


class ScriptEditorWindow(tk.Toplevel):
    def __init__(self, master: "AutomationApp") -> None:
        super().__init__(master)
        self.app = master
        self.title("CRUD de Roteiros (JSON)")
        self.geometry("980x620")
        self._build()
        self.refresh_list()

    def _build(self) -> None:
        left = ttk.Frame(self)
        left.pack(side="left", fill="y", padx=8, pady=8)
        right = ttk.Frame(self)
        right.pack(side="right", fill="both", expand=True, padx=8, pady=8)

        self.listbox = tk.Listbox(left, width=30)
        self.listbox.pack(fill="y", expand=True)
        self.listbox.bind("<<ListboxSelect>>", lambda _e: self.load_selected())

        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Novo", command=self.new_file).pack(fill="x", pady=2)
        ttk.Button(btns, text="Salvar", command=self.save_file).pack(fill="x", pady=2)
        ttk.Button(btns, text="Excluir", command=self.delete_file).pack(fill="x", pady=2)
        ttk.Button(btns, text="Carregar na tela principal", command=self.load_into_main).pack(fill="x", pady=2)
        ttk.Button(btns, text="Atualizar lista", command=self.refresh_list).pack(fill="x", pady=2)

        top = ttk.Frame(right)
        top.pack(fill="x")
        ttk.Label(top, text="Nome do arquivo (sem .json):").pack(side="left")
        self.name_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.name_var, width=40).pack(side="left", padx=8)

        self.editor = tk.Text(right, wrap="none", undo=True)
        self.editor.pack(fill="both", expand=True, pady=(8, 0))
        self.new_file()

    def _default_script(self) -> str:
        sample = [
            {
                "info": "Mover mouse",
                "mouse": {"x": 250, "y": 620, "acao": "mover"},
            },
            {
                "info": "Clique esquerdo",
                "mouse": {"x": 250, "y": 620, "acao": "clicar_esquerdo"},
            },
            {
                "info": "Digitar da tabela padrão",
                "teclado": {"campo_tabela": "NOME"},
            },
            {
                "info": "Pressionar ENTER",
                "teclado": {"pressionar": "enter"},
            },
        ]
        return json.dumps(sample, ensure_ascii=False)

    def refresh_list(self) -> None:
        self.listbox.delete(0, tk.END)
        for path in sorted(ROTEIROS_DIR.glob("*.json")):
            self.listbox.insert(tk.END, path.stem)
        self.app.refresh_script_dropdown()

    def load_selected(self) -> None:
        selection = self.listbox.curselection()
        if not selection:
            return
        name = self.listbox.get(selection[0])
        path = ROTEIROS_DIR / f"{name}.json"
        self.name_var.set(name)
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", path.read_text(encoding="utf-8"))

    def new_file(self) -> None:
        self.name_var.set("novo_roteiro")
        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", self._default_script())

    def save_file(self) -> None:
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Erro", "Informe o nome do arquivo.")
            return

        content = self.editor.get("1.0", tk.END).strip()
        try:
            parsed = json.loads(content)
            if not isinstance(parsed, list):
                raise ValueError("O roteiro deve ser uma lista JSON.")
        except Exception as exc:
            messagebox.showerror("Erro de JSON", str(exc))
            return

        path = ROTEIROS_DIR / f"{name}.json"
        path.write_text(content, encoding="utf-8")
        self.app.refresh_script_dropdown()
        self.refresh_list()
        messagebox.showinfo("OK", f"Roteiro salvo: {path.name}")

    def delete_file(self) -> None:
        name = self.name_var.get().strip()
        if not name:
            return
        path = ROTEIROS_DIR / f"{name}.json"
        if path.exists() and messagebox.askyesno("Confirmar", f"Excluir {path.name}?"):
            path.unlink()
            self.app.refresh_script_dropdown()
            self.refresh_list()
            self.new_file()

    def load_into_main(self) -> None:
        self.app.script_text.delete("1.0", tk.END)
        self.app.script_text.insert("1.0", self.editor.get("1.0", tk.END))
        self.app.script_name_var.set(self.name_var.get().strip())
        self.app.apply_json_highlight()
        messagebox.showinfo("OK", "Roteiro carregado na tela principal.")


class AutomationApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Automação de Teclado e Mouse")
        self.geometry("1220x780")

        self.script_name_var = tk.StringVar()
        self.selected_script_var = tk.StringVar()
        self.selected_table_var = tk.StringVar()
        self.start_delay_var = tk.StringVar(value="0")
        self.delay_var = tk.StringVar(value="0.3")
        self.repetitions_var = tk.StringVar(value="1")
        self.mouse_position_var = tk.StringVar(value="Mouse: X=0 Y=0")

        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.stop_event = threading.Event()
        self.worker_thread: threading.Thread | None = None

        self._build_ui()
        self.refresh_script_dropdown()
        self.refresh_table_dropdown()
        self.after(150, self.process_log_queue)
        self.after(100, self.update_mouse_position)

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Roteiro salvo:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.script_combo = ttk.Combobox(top, textvariable=self.selected_script_var, state="readonly", width=30)
        self.script_combo.grid(row=0, column=1, sticky="we", padx=4, pady=4)
        ttk.Button(top, text="Carregar", command=self.load_selected_script).grid(row=0, column=2, padx=4, pady=4)
        ttk.Button(top, text="Salvar como roteiro", command=self.save_script_from_main).grid(row=0, column=3, padx=4, pady=4)
        ttk.Button(top, text="CRUD Roteiros", command=self.open_script_editor).grid(row=0, column=4, padx=4, pady=4)
        ttk.Button(top, text="CRUD Tabelas", command=self.open_table_editor).grid(row=0, column=5, padx=4, pady=4)
        ttk.Button(top, text="Help", command=self.open_help).grid(row=0, column=6, padx=4, pady=4)

        ttk.Label(top, text="Nome do roteiro atual:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(top, textvariable=self.script_name_var, width=32).grid(row=1, column=1, sticky="we", padx=4, pady=4)

        ttk.Label(top, text="Tabela:").grid(row=1, column=2, sticky="e", padx=4, pady=4)
        self.table_combo = ttk.Combobox(top, textvariable=self.selected_table_var, state="readonly", width=28)
        self.table_combo.grid(row=1, column=3, sticky="we", padx=4, pady=4)

        ttk.Label(top, text="Delay inicial (s):").grid(row=1, column=4, sticky="e", padx=4, pady=4)
        ttk.Entry(top, textvariable=self.start_delay_var, width=10).grid(row=1, column=5, sticky="w", padx=4, pady=4)

        ttk.Label(top, text="Delay entre passos (s):").grid(row=1, column=6, sticky="e", padx=4, pady=4)
        ttk.Entry(top, textvariable=self.delay_var, width=10).grid(row=1, column=7, sticky="w", padx=4, pady=4)

        ttk.Label(top, text="Repetições:").grid(row=1, column=8, sticky="e", padx=4, pady=4)
        ttk.Entry(top, textvariable=self.repetitions_var, width=10).grid(row=1, column=9, sticky="w", padx=4, pady=4)

        ttk.Label(top, textvariable=self.mouse_position_var).grid(
            row=2, column=0, columnspan=10, sticky="w", padx=4, pady=(2, 4)
        )

        for c in range(10):
            top.columnconfigure(c, weight=1)

        center = ttk.Panedwindow(self, orient="vertical")
        center.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        script_frame = ttk.Labelframe(center, text="Roteiro para execução")
        console_frame = ttk.Labelframe(center, text="Console")
        center.add(script_frame, weight=3)
        center.add(console_frame, weight=2)

        self.script_text = tk.Text(
            script_frame,
            wrap="none",
            undo=True,
            height=20,
            bg="#0f172a",
            fg="#e2e8f0",
            insertbackground="#f8fafc",
            selectbackground="#334155",
            padx=10,
            pady=10,
            font=("Consolas", 10),
        )
        self.script_text.pack(fill="both", expand=True, padx=6, pady=6)
        self._configure_script_tags()
        self.script_text.insert(
            "1.0",
            json.dumps(
                [
                    {
                        "info": "Exemplo mouse",
                        "mouse": {"x": 250, "y": 620, "acao": "clicar_esquerdo"},
                    },
                    {
                        "info": "Exemplo com tabela selecionada",
                        "teclado": {"campo_tabela": "NOME"},
                    },
                    {
                        "info": "Exemplo pressionar ENTER",
                        "teclado": {"pressionar": "enter"},
                    },
                ],
                ensure_ascii=False,
            ),
        )
        self.apply_json_highlight()
        self.script_text.bind("<KeyRelease>", lambda _e: self.apply_json_highlight())
        self.script_text.bind("<<Paste>>", lambda _e: self.after(10, self.apply_json_highlight))

        controls = ttk.Frame(self, padding=(10, 0, 10, 10))
        controls.pack(fill="x")
        ttk.Button(controls, text="Executar", command=self.start_execution).pack(side="left", padx=4)
        ttk.Button(controls, text="Parar", command=self.stop_execution).pack(side="left", padx=4)
        ttk.Button(controls, text="Limpar console", command=self.clear_console).pack(side="left", padx=4)

        self.console_text = tk.Text(
            console_frame,
            wrap="word",
            state="disabled",
            height=12,
            bg="#111827",
            fg="#e5e7eb",
            insertbackground="#ffffff",
            padx=8,
            pady=8,
            font=("Consolas", 10),
        )
        self.console_text.pack(fill="both", expand=True, padx=6, pady=6)
        self._configure_console_tags()

    def _configure_script_tags(self) -> None:
        self.script_text.tag_configure("json_key", foreground="#93c5fd")
        self.script_text.tag_configure("json_string", foreground="#86efac")
        self.script_text.tag_configure("json_number", foreground="#fca5a5")
        self.script_text.tag_configure("json_boolean", foreground="#f9a8d4")
        self.script_text.tag_configure("json_brace", foreground="#cbd5e1")

    def _configure_console_tags(self) -> None:
        self.console_text.tag_configure("default", foreground="#e5e7eb")
        self.console_text.tag_configure("info", foreground="#93c5fd")
        self.console_text.tag_configure("success", foreground="#86efac")
        self.console_text.tag_configure("warning", foreground="#fde68a")
        self.console_text.tag_configure("error", foreground="#fca5a5")
        self.console_text.tag_configure("step", foreground="#e5e7eb")
        self.console_text.tag_configure("iteration", foreground="#ffffff", font=("Consolas", 10, "bold"))
        self.console_text.tag_configure("elapsed", foreground="#ff6b6b")

    def apply_json_highlight(self) -> None:
        text_widget = self.script_text
        content = text_widget.get("1.0", "end-1c")

        for tag in ("json_key", "json_string", "json_number", "json_boolean", "json_brace"):
            text_widget.tag_remove(tag, "1.0", tk.END)

        for match in re.finditer(r'"([^"\\]*(?:\\.[^"\\]*)*)"\s*:', content):
            start = f"1.0+{match.start()}c"
            end = f"1.0+{match.end() - 1}c"
            text_widget.tag_add("json_key", start, end)

        for match in re.finditer(r':\s*"([^"\\]*(?:\\.[^"\\]*)*)"', content):
            colon_pos = match.group(0).find('"')
            start = f"1.0+{match.start() + colon_pos}c"
            end = f"1.0+{match.end()}c"
            text_widget.tag_add("json_string", start, end)

        for match in re.finditer(r"\b-?\d+(?:\.\d+)?\b", content):
            start = f"1.0+{match.start()}c"
            end = f"1.0+{match.end()}c"
            text_widget.tag_add("json_number", start, end)

        for match in re.finditer(r"\b(true|false|null)\b", content):
            start = f"1.0+{match.start()}c"
            end = f"1.0+{match.end()}c"
            text_widget.tag_add("json_boolean", start, end)

        for match in re.finditer(r"[\{\}\[\]]", content):
            start = f"1.0+{match.start()}c"
            end = f"1.0+{match.end()}c"
            text_widget.tag_add("json_brace", start, end)

    def refresh_script_dropdown(self) -> None:
        values = [""] + [path.stem for path in sorted(ROTEIROS_DIR.glob("*.json"))]
        self.script_combo["values"] = values
        if self.selected_script_var.get() not in values:
            self.selected_script_var.set("")

    def refresh_table_dropdown(self) -> None:
        values = [""] + [path.stem for path in sorted(TABELAS_DIR.glob("*.csv"))]
        self.table_combo["values"] = values
        if self.selected_table_var.get() not in values:
            self.selected_table_var.set("")

    def open_help(self) -> None:
        win = tk.Toplevel(self)
        win.title("Help - Comandos do Roteiro")
        win.geometry("980x720")
        txt = tk.Text(win, wrap="word")
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", HELP_TEXT)
        txt.config(state="disabled")

    def open_table_editor(self) -> None:
        TableEditorWindow(self)

    def open_script_editor(self) -> None:
        ScriptEditorWindow(self)

    def load_selected_script(self) -> None:
        name = self.selected_script_var.get().strip()
        if not name:
            return
        path = ROTEIROS_DIR / f"{name}.json"
        if not path.exists():
            messagebox.showerror("Erro", "Roteiro não encontrado.")
            return
        self.script_name_var.set(name)
        self.script_text.delete("1.0", tk.END)
        self.script_text.insert("1.0", path.read_text(encoding="utf-8"))
        self.apply_json_highlight()
        self.log(f"Roteiro carregado: {name}", "info")

    def save_script_from_main(self) -> None:
        name = self.script_name_var.get().strip() or "roteiro_principal"
        content = self.script_text.get("1.0", tk.END).strip()
        try:
            parsed = json.loads(content)
            if not isinstance(parsed, list):
                raise ValueError("O roteiro deve ser uma lista JSON.")
        except Exception as exc:
            messagebox.showerror("Erro de JSON", str(exc))
            return
        path = ROTEIROS_DIR / f"{name}.json"
        path.write_text(content, encoding="utf-8")
        self.refresh_script_dropdown()
        self.selected_script_var.set(name)
        self.log(f"Roteiro salvo: {path.name}", "success")
        messagebox.showinfo("OK", f"Roteiro salvo: {path.name}")

    def start_execution(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Aviso", "Já existe uma execução em andamento.")
            return

        try:
            steps = json.loads(self.script_text.get("1.0", tk.END).strip())
            if not isinstance(steps, list):
                raise ValueError("O roteiro deve ser uma lista JSON.")
            repetitions = int(self.repetitions_var.get())
            float(self.start_delay_var.get())
            float(self.delay_var.get())
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
            return

        self.clear_console()
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
                self.log("Execução iniciada.", "info")
                runner.run(steps, repetitions, selected_table)
            except Exception as exc:
                self.log(f"ERRO: {exc}", "error")
            finally:
                self.log("Thread finalizada.", "info")

        self.worker_thread = threading.Thread(target=target, daemon=True)
        self.worker_thread.start()

    def stop_execution(self) -> None:
        self.stop_event.set()
        self.log("Solicitação de parada registrada.", "warning")

    def clear_console(self) -> None:
        self.console_text.config(state="normal")
        self.console_text.delete("1.0", tk.END)
        self.console_text.config(state="disabled")

    def log(self, message: str, tag: str = "default") -> None:
        timestamp = time.strftime("%H:%M:%S")
        self.log_queue.put((f"[{timestamp}] {message}", tag))

    def process_log_queue(self) -> None:
        while not self.log_queue.empty():
            message, tag = self.log_queue.get()
            self.console_text.config(state="normal")
            self.console_text.insert(tk.END, message + "\n", tag)
            self.console_text.see(tk.END)
            self.console_text.config(state="disabled")
        self.after(150, self.process_log_queue)

    def update_mouse_position(self) -> None:
        try:
            x, y = pyautogui.position()
            self.mouse_position_var.set(f"Mouse: X={x} Y={y}")
        except Exception:
            self.mouse_position_var.set("Mouse: X=? Y=?")
        self.after(100, self.update_mouse_position)


if __name__ == "__main__":
    app = AutomationApp()
    app.mainloop()