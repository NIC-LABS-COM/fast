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


def parse_packages_txt(raw_content: str) -> list[dict]:
    """Converte saida do buscaPacotes.vbs em lista de dicts com 'packageName'."""
    text = raw_content.replace("\\n", "\n")
    results: list[dict] = []
    for line in text.splitlines():
        package_name = line.strip()
        if not package_name:
            continue
        results.append({"packageName": package_name})
    return results


def parse_versions_metadata_txt(raw_content: str) -> list[dict]:
    """Converte JSON gerado pelo report Z_GET_VERSION_METADATA em lista de dicts.

    O report ABAP usa /ui2/cl_json=>serialize com pretty_name=camel_case,
    gerando JSON direto no arquivo. Campos esperados:
    id, title, updated, request, fileName, category.
    """
    import json

    text = raw_content.replace("\\n", "\n").strip()
    if not text:
        return []

    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return []

    if not isinstance(data, list):
        data = [data]

    results: list[dict] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        results.append({
            "id":       str(item.get("id", "")),
            "title":    str(item.get("title", "")),
            "updated":  str(item.get("updated", "")),
            "request":  str(item.get("request", "")),
            "fileName": str(item.get("fileName", item.get("file_name", ""))),
            "category": str(item.get("category", "")),
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
