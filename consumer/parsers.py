"""Parsers de TXT gerados pelos scripts VBS."""


def parse_requests_txt(raw_content: str) -> list[dict]:
    """Converte 'REQUEST ; DESCRICAO' ou apenas 'REQUEST' em lista de dicts."""
    text = raw_content.replace("\\n", "\n")
    results: list[dict] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if ";" in stripped:
            parts = stripped.split(";", maxsplit=1)
            request = parts[0].strip()
            description = parts[1].strip() if len(parts) > 1 else ""
        else:
            request = stripped
            description = ""
        if not request:
            continue
        results.append({
            "request": request,
            "descriptions": [description] if description else [],
        })
    return results


def parse_reports_txt(raw_content: str) -> list[dict]:
    """Converte 'FILENAME ; CATEGORY ; PACKAGE ; FGROUP' ou apenas 'FILENAME' em lista de dicts."""
    text = raw_content.replace("\\n", "\n")
    results: list[dict] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if ";" in stripped:
            parts = [p.strip() for p in stripped.split(";")]
        else:
            parts = [stripped]
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
