from __future__ import annotations


def manter(texto: str) -> str:
    return texto


def upper(texto: str) -> str:
    return texto.upper()


def lower(texto: str) -> str:
    return texto.lower()


def strip(texto: str) -> str:
    return texto.strip()


def somente_digitos(texto: str) -> str:
    return ''.join(ch for ch in texto if ch.isdigit())


def remover_espacos(texto: str) -> str:
    return ' '.join(texto.split())


def capitalizar(texto: str) -> str:
    return texto.title()

def trocar_prefixo_tabela_rf(texto: str) -> str:
    """
    Exemplo:
    ZS4_VCI_A003 -> S4H_TB_A003
    """
    if texto is None:
        return ""

    texto = str(texto).strip()

    prefixo_antigo = "ZS4_VCI_"
    prefixo_novo = "S4H_TB_"

    if texto.startswith(prefixo_antigo):
        return prefixo_novo + texto[len(prefixo_antigo):]

    return texto

TRANSFORMACOES = {
    'manter': manter,
    'upper': upper,
    'lower': lower,
    'strip': strip,
    'somente_digitos': somente_digitos,
    'remover_espacos': remover_espacos,
    'capitalizar': capitalizar,
    'trocar_prefixo_tabela_rf': trocar_prefixo_tabela_rf,
}
