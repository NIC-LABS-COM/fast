"""
requests_parser.py — Parser do arquivo .txt de requests do SAP.

Converte o conteudo bruto (linhas "REQUEST ; DESCRICAO") em uma lista
de dicionarios no formato:
  [{"request": "A4HK900006", "descriptions": ["LS"]}, ...]
"""


def parse_requests_txt(raw_content: str) -> list[dict]:
    """
    Faz o parsing do conteudo do arquivo de requests.

    O conteudo pode vir com '\\n' literal (codificado pelo VBS) ou
    com quebras de linha reais.

    Args:
        raw_content: Texto bruto vindo do stdout do VBS.

    Returns:
        Lista de dicts com 'request' e 'descriptions'.
    """
    # O VBS codifica quebras de linha como literal \\n
    text = raw_content.replace("\\n", "\n")

    results: list[dict] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue

        if ";" not in stripped:
            continue

        parts = stripped.split(";", maxsplit=1)
        request = parts[0].strip()
        description = parts[1].strip() if len(parts) > 1 else ""

        if not request:
            continue

        results.append({
            "request": request,
            "descriptions": [description] if description else [],
        })

    return results
