from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


def _replace_tokens(value: Any, variables: dict) -> Any:
    """
    Substitui tokens no formato {{variavel}} pelos valores recebidos
    em payload["variables"].
    """
    if not isinstance(value, str):
        return value

    out = value

    for key, val in variables.items():
        token = "{{" + str(key) + "}}"
        replacement = "" if val is None else str(val)
        out = out.replace(token, replacement)

    return out


def _copy_row_style(
    ws,
    source_row: int,
    target_row: int,
    max_col: int,
):
    """
    Copia a formatação de uma linha-modelo para outra linha.

    Células MergedCell são ignoradas porque são somente leitura.
    """
    for col in range(1, max_col + 1):

        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)

        # Proteção para células mescladas.
        if isinstance(src, MergedCell):
            continue

        if isinstance(dst, MergedCell):
            continue

        if src.has_style:
            dst._style = copy(src._style)

        if src.number_format:
            dst.number_format = src.number_format

        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)

    # Copia altura da linha.
    source_height = ws.row_dimensions[source_row].height

    if source_height is not None:
        ws.row_dimensions[target_row].height = source_height


def _find_marker(ws, marker: str):
    """
    Procura um marcador dentro da planilha.

    Exemplo:
    {{PARTIDAS}}
    {{OCIOSAS}}
    {{RECARGAS}}
    """

    for row in ws.iter_rows():

        for cell in row:

            # Células internas de um merge não possuem valor editável.
            if isinstance(cell, MergedCell):
                continue

            if cell.value == marker:
                return cell.row, cell.column

    return None


def _replace_variables(ws, variables: dict):
    """
    Substitui variáveis simples existentes no template.

    Exemplo:
    {{linha}}
    {{data}}
    {{titulo}}
    """

    if not variables:
        return

    for row in ws.iter_rows():

        for cell in row:

            # IMPORTANTE:
            # somente a célula superior esquerda de um merge é editável.
            if isinstance(cell, MergedCell):
                continue

            if isinstance(cell.value, str):
                cell.value = _replace_tokens(
                    cell.value,
                    variables,
                )


def generate(
    template_path: Path,
    config: dict,
    payload: dict,
) -> BytesIO:
    """
    Gera um relatório Excel utilizando:

    - template XLSX
    - config.yaml
    - payload JSON

    Retorna o arquivo gerado em memória (BytesIO).
    """

    # ---------------------------------------------------------
    # 1. Valida template
    # ---------------------------------------------------------

    template_path = Path(template_path)

    if not template_path.exists():
        raise FileNotFoundError(
            f"Template não encontrado: {template_path}"
        )

    # ---------------------------------------------------------
    # 2. Abre workbook
    # ---------------------------------------------------------

    wb = load_workbook(template_path)

    # ---------------------------------------------------------
    # 3. Seleciona planilha
    # ---------------------------------------------------------

    sheet_name = config.get("sheet") or wb.sheetnames[0]

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"A planilha '{sheet_name}' não existe no template. "
            f"Planilhas disponíveis: {wb.sheetnames}"
        )

    ws = wb[sheet_name]

    # ---------------------------------------------------------
    # 4. Variáveis simples
    # ---------------------------------------------------------

    variables = payload.get("variables", {}) or {}

    _replace_variables(
        ws,
        variables,
    )

    # ---------------------------------------------------------
    # 5. Seções dinâmicas
    # ---------------------------------------------------------

    sections = config.get("sections", {}) or {}

    found = []

    # Primeiro encontra todos os marcadores antes de alterar
    # a estrutura da planilha.
    for section_name, section_cfg in sections.items():

        marker = section_cfg.get(
            "marker",
            "{{" + section_name.upper() + "}}",
        )

        pos = _find_marker(
            ws,
            marker,
        )

        if pos:

            found.append(
                (
                    pos[0],
                    pos[1],
                    section_name,
                    section_cfg,
                    marker,
                )
            )

    # ---------------------------------------------------------
    # 6. Processa seções de baixo para cima
    # ---------------------------------------------------------
    #
    # Isso é importante porque inserir linhas em uma seção
    # inferior não altera a posição das seções superiores.
    # ---------------------------------------------------------

    for (
        marker_row,
        marker_col,
        section_name,
        section_cfg,
        marker,
    ) in sorted(found, key=lambda item: item[0], reverse=True):

        records = (
            payload
            .get("sections", {})
            .get(section_name, [])
            or []
        )

        template_row = int(
            section_cfg.get(
                "template_row",
                marker_row + 1,
            )
        )

        start_row = int(
            section_cfg.get(
                "start_row",
                template_row,
            )
        )

        columns = (
            section_cfg.get("columns", {})
            or {}
        )

        # -----------------------------------------------------
        # Remove marcador
        # -----------------------------------------------------

        marker_cell = ws.cell(
            marker_row,
            marker_col,
        )

        if not isinstance(marker_cell, MergedCell):
            marker_cell.value = None

        # -----------------------------------------------------
        # Se não houver registros
        # -----------------------------------------------------

        if not records:

            if section_cfg.get(
                "delete_template_row_when_empty",
                False,
            ):
                ws.delete_rows(
                    template_row,
                    1,
                )

            continue

        # -----------------------------------------------------
        # Insere linhas adicionais
        # -----------------------------------------------------
        #
        # Já existe uma linha-modelo.
        #
        # 1 registro  -> 0 novas linhas
        # 2 registros -> 1 nova linha
        # 10 registros -> 9 novas linhas
        # -----------------------------------------------------

        if len(records) > 1:

            ws.insert_rows(
                start_row + 1,
                amount=len(records) - 1,
            )

            # Replica o estilo da linha modelo.
            for i in range(
                1,
                len(records),
            ):

                target_row = start_row + i

                _copy_row_style(
                    ws,
                    template_row,
                    target_row,
                    ws.max_column,
                )

        # -----------------------------------------------------
        # Caso a linha inicial seja diferente da linha-modelo
        # -----------------------------------------------------

        if start_row != template_row:

            _copy_row_style(
                ws,
                template_row,
                start_row,
                ws.max_column,
            )

        # -----------------------------------------------------
        # Preenche registros
        # -----------------------------------------------------

        for idx, record in enumerate(records):

            row_num = start_row + idx

            for field, col in columns.items():

                col_num = int(col)

                target_cell = ws.cell(
                    row_num,
                    col_num,
                )

                # Não tenta escrever em célula interna
                # pertencente a um merge.
                if isinstance(target_cell, MergedCell):
                    continue

                target_cell.value = record.get(field)

    # ---------------------------------------------------------
    # 7. Salva arquivo em memória
    # ---------------------------------------------------------

    out = BytesIO()

    wb.save(out)

    out.seek(0)

    return out
