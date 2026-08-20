from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any
from openpyxl import load_workbook


def _replace_tokens(value: Any, variables: dict) -> Any:
    if not isinstance(value, str):
        return value
    out = value
    for key, val in variables.items():
        out = out.replace("{{" + str(key) + "}}", "" if val is None else str(val))
    return out


def _copy_row_style(ws, source_row: int, target_row: int, max_col: int):
    for col in range(1, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)
    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def _find_marker(ws, marker: str):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == marker:
                return cell.row, cell.column
    return None


def generate(template_path: Path, config: dict, payload: dict) -> BytesIO:
    wb = load_workbook(template_path)
    sheet_name = config.get("sheet") or wb.sheetnames[0]
    ws = wb[sheet_name]

    variables = payload.get("variables", {})
    for row in ws.iter_rows():
        for cell in row:
            cell.value = _replace_tokens(cell.value, variables)

    sections = config.get("sections", {}) or {}
    # Processa de baixo para cima para que inserções não invalidem marcadores superiores.
    found = []
    for section_name, section_cfg in sections.items():
        marker = section_cfg.get("marker", "{{" + section_name.upper() + "}}")
        pos = _find_marker(ws, marker)
        if pos:
            found.append((pos[0], pos[1], section_name, section_cfg, marker))

    for marker_row, marker_col, section_name, section_cfg, marker in sorted(found, reverse=True):
        records = payload.get("sections", {}).get(section_name, []) or []
        template_row = int(section_cfg.get("template_row", marker_row + 1))
        start_row = int(section_cfg.get("start_row", template_row))
        columns = section_cfg.get("columns", {}) or {}

        # Remove o texto marcador sem alterar o restante do layout.
        ws.cell(marker_row, marker_col).value = None

        if not records:
            if section_cfg.get("delete_template_row_when_empty", False):
                ws.delete_rows(template_row, 1)
            continue

        # Uma linha já existe como linha-modelo; insere apenas as adicionais.
        if len(records) > 1:
            ws.insert_rows(start_row + 1, amount=len(records) - 1)
            for i in range(1, len(records)):
                _copy_row_style(ws, template_row, start_row + i, ws.max_column)

        # Também garante estilo na primeira linha caso start_row seja diferente.
        if start_row != template_row:
            _copy_row_style(ws, template_row, start_row, ws.max_column)

        for idx, record in enumerate(records):
            row_num = start_row + idx
            for field, col in columns.items():
                ws.cell(row_num, int(col)).value = record.get(field)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out
