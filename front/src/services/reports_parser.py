"""
reports_parser.py — Parser do arquivo .txt de reports do SAP (SE38).

Converte o conteudo bruto (linhas "FILENAME ; CATEGORY ; PACKAGE ; FGROUP")
em uma lista de dicionarios no formato:
  [{"fileName": "Z_REPORT_1", "category": "PROGRAM", "packageName": "Z_PKG", "functionGroup": ""}, ...]
"""

VALID_CATEGORIES = {"PROGRAM", "FUNCTION_MODULE", "CLASS"}


def parse_reports_txt(raw_content: str) -> list[dict]:
    """
    Faz o parsing do conteudo do arquivo de reports.

    Formato esperado por linha (separador ;):
      pos 0: fileName
      pos 1: category (PROGRAM | FUNCTION_MODULE | CLASS)
      pos 2: packageName
      pos 3: functionGroup (opcional)

    Args:
        raw_content: Texto bruto vindo do stdout do VBS.

    Returns:
        Lista de dicts com fileName, category, packageName, functionGroup.
    """
    text = raw_content.replace("\\n", "\n")

    results: list[dict] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue

        if ";" not in stripped:
            continue

        parts = [p.strip() for p in stripped.split(";")]

        file_name = parts[0] if len(parts) > 0 else ""
        if not file_name:
            continue

        category = parts[1] if len(parts) > 1 else ""
        package_name = parts[2] if len(parts) > 2 else ""
        function_group = parts[3] if len(parts) > 3 else ""

        results.append({
            "fileName": file_name,
            "category": category,
            "packageName": package_name,
            "functionGroup": function_group,
        })

    return results
