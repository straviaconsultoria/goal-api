from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange


# ============================================================
# TOKENS
# ============================================================

def _replace_tokens(value: Any, variables: dict) -> Any:
    """
    Substitui tokens no formato {{CAMPO}} pelos valores recebidos
    em payload["variables"].

    Exemplo:
        {{LINHA}} -> 082
    """

    if not isinstance(value, str):
        return value

    out = value

    for key, val in variables.items():
        token = "{{" + str(key) + "}}"
        replacement = "" if val is None else str(val)

        out = out.replace(
            token,
            replacement,
        )

    return out


# ============================================================
# CÓPIA DE ESTILO
# ============================================================

def _copy_row_style(
    ws,
    source_row: int,
    target_row: int,
    max_col: int,
):
    """
    Copia toda a formatação visual de uma linha-modelo.

    Utilizado para preservar:
    - zebrado
    - bordas
    - preenchimento
    - fonte
    - alinhamento
    - formatos numéricos
    - proteção
    - altura da linha
    """

    for col in range(1, max_col + 1):

        src = ws.cell(
            source_row,
            col,
        )

        dst = ws.cell(
            target_row,
            col,
        )

        if isinstance(src, MergedCell):
            continue

        if isinstance(dst, MergedCell):
            continue

        if src.has_style:
            dst._style = copy(src._style)

        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)

        if src.number_format:
            dst.number_format = src.number_format

    source_height = ws.row_dimensions[source_row].height

    if source_height is not None:
        ws.row_dimensions[target_row].height = source_height


# ============================================================
# MARCADORES
# ============================================================

def _find_marker(
    ws,
    marker: str,
):
    """
    Procura um marcador técnico dentro da planilha.

    Exemplo:
        {{REGISTROS_API}}
    """

    for row in ws.iter_rows():

        for cell in row:

            if isinstance(cell, MergedCell):
                continue

            if cell.value == marker:

                return (
                    cell.row,
                    cell.column,
                )

    return None


# ============================================================
# VARIÁVEIS DO TEMPLATE
# ============================================================

def _replace_variables(
    ws,
    variables: dict,
):
    """
    Substitui variáveis simples dentro das células.

    Exemplo:
        {{LINHA}}
        {{ITINERARIO}}
        {{TIPO_DIA}}
        {{ROUTE_1}}
    """

    if not variables:
        return

    for row in ws.iter_rows():

        for cell in row:

            if isinstance(cell, MergedCell):
                continue

            if isinstance(cell.value, str):

                cell.value = _replace_tokens(
                    cell.value,
                    variables,
                )


# ============================================================
# NOME DA ABA
# ============================================================

def _replace_sheet_title(
    ws,
    variables: dict,
):
    """
    Permite utilizar tokens também no nome da aba.

    Exemplo:
        {{LINHA}} -> 082
    """

    new_title = _replace_tokens(
        ws.title,
        variables,
    )

    if new_title == ws.title:
        return

    # Caracteres não permitidos pelo Excel em nomes de abas.
    invalid_chars = [
        "\\",
        "/",
        "*",
        "?",
        ":",
        "[",
        "]",
    ]

    for char in invalid_chars:
        new_title = new_title.replace(
            char,
            "-",
        )

    # Excel limita nomes de abas a 31 caracteres.
    new_title = new_title[:31].strip()

    if not new_title:
        new_title = "Relatorio"

    ws.title = new_title


# ============================================================
# ÁREA DE IMPRESSÃO
# ============================================================

def _update_print_area(
    ws,
    inserted_rows: int,
    insertion_start_row: int,
):
    """
    Expande a área de impressão quando novas linhas são inseridas.

    Exemplo original:
        A1:U15

    Se forem adicionadas 56 linhas:
        A1:U71
    """

    if inserted_rows <= 0:
        return

    print_area = ws.print_area

    if not print_area:
        return

    ranges = []

    try:
        ranges = list(print_area.ranges)
    except Exception:
        return

    new_ranges = []

    for cell_range in ranges:

        min_col = cell_range.min_col
        min_row = cell_range.min_row
        max_col = cell_range.max_col
        max_row = cell_range.max_row

        if max_row >= insertion_start_row:
            max_row += inserted_rows

        start_cell = (
            f"{get_column_letter(min_col)}"
            f"{min_row}"
        )

        end_cell = (
            f"{get_column_letter(max_col)}"
            f"{max_row}"
        )

        new_ranges.append(
            f"{start_cell}:{end_cell}"
        )

    if new_ranges:
        ws.print_area = ",".join(new_ranges)


# ============================================================
# GERAÇÃO DO RELATÓRIO
# ============================================================

def generate(
    template_path: Path,
    config: dict,
    payload: dict,
) -> BytesIO:
    """
    Motor genérico de geração de relatórios Excel.

    Suporta:

    - template XLSX
    - variáveis {{TOKEN}}
    - seções dinâmicas
    - múltiplas linhas-modelo
    - zebrado
    - quantidade variável de registros
    - rodapé deslocável
    - células mescladas
    - nome de aba dinâmico
    - atualização da área de impressão
    """

    # ========================================================
    # 1. VALIDA TEMPLATE
    # ========================================================

    template_path = Path(
        template_path
    )

    if not template_path.exists():

        raise FileNotFoundError(
            f"Template não encontrado: "
            f"{template_path}"
        )

    # ========================================================
    # 2. ABRE WORKBOOK
    # ========================================================

    wb = load_workbook(
        template_path
    )

    # ========================================================
    # 3. LOCALIZA A PLANILHA
    # ========================================================

    configured_sheet = config.get(
        "sheet"
    )

    if configured_sheet:

        if configured_sheet not in wb.sheetnames:

            raise ValueError(
                f"A planilha '{configured_sheet}' "
                f"não existe no template. "
                f"Planilhas disponíveis: "
                f"{wb.sheetnames}"
            )

        ws = wb[
            configured_sheet
        ]

    else:

        ws = wb[
            wb.sheetnames[0]
        ]

    # ========================================================
    # 4. VARIÁVEIS
    # ========================================================

    variables = (
        payload.get(
            "variables",
            {}
        )
        or {}
    )

    # ========================================================
    # 5. SEÇÕES DINÂMICAS
    # ========================================================
    #
    # IMPORTANTE:
    # Localizamos os marcadores ANTES de substituir tokens.
    # ========================================================

    sections = (
        config.get(
            "sections",
            {}
        )
        or {}
    )

    found_sections = []

    for (
        section_name,
        section_cfg,
    ) in sections.items():

        marker = section_cfg.get(
            "marker",
            "{{"
            + section_name.upper()
            + "}}",
        )

        position = _find_marker(
            ws,
            marker,
        )

        if position is None:

            raise ValueError(
                f"Marcador '{marker}' "
                f"da seção '{section_name}' "
                f"não encontrado no template."
            )

        marker_row = position[0]
        marker_col = position[1]

        found_sections.append(
            (
                marker_row,
                marker_col,
                section_name,
                section_cfg,
                marker,
            )
        )

    # ========================================================
    # 6. PROCESSA SEÇÕES DE BAIXO PARA CIMA
    # ========================================================

    for (
        marker_row,
        marker_col,
        section_name,
        section_cfg,
        marker,
    ) in sorted(
        found_sections,
        key=lambda item: item[0],
        reverse=True,
    ):

        records = (
            payload
            .get(
                "sections",
                {}
            )
            .get(
                section_name,
                []
            )
            or []
        )

        # ----------------------------------------------------
        # LINHAS-MODELO
        # ----------------------------------------------------

        template_rows = (
            section_cfg.get(
                "template_rows"
            )
            or []
        )

        # Compatibilidade com configs antigos.
        if not template_rows:

            old_template_row = (
                section_cfg.get(
                    "template_row"
                )
            )

            if old_template_row:

                template_rows = [
                    int(
                        old_template_row
                    )
                ]

        if not template_rows:

            template_rows = [
                marker_row
            ]

        template_rows = [
            int(row)
            for row
            in template_rows
        ]

        # ----------------------------------------------------
        # PRIMEIRA LINHA DOS DADOS
        # ----------------------------------------------------

        start_row = int(
            section_cfg.get(
                "start_row",
                template_rows[0],
            )
        )

        # ----------------------------------------------------
        # COLUNAS
        # ----------------------------------------------------

        columns = (
            section_cfg.get(
                "columns",
                {}
            )
            or {}
        )

        # ----------------------------------------------------
        # REMOVE MARCADOR
        # ----------------------------------------------------

        marker_cell = ws.cell(
            marker_row,
            marker_col,
        )

        if not isinstance(
            marker_cell,
            MergedCell,
        ):

            marker_cell.value = None

        # ----------------------------------------------------
        # QUANTIDADE DE LINHAS-MODELO
        # ----------------------------------------------------

        model_count = len(
            template_rows
        )

        record_count = len(
            records
        )

        # ----------------------------------------------------
        # SEM REGISTROS
        # ----------------------------------------------------

        if record_count == 0:

            if section_cfg.get(
                "clear_template_rows_when_empty",
                True,
            ):

                for row_num in template_rows:

                    for col in columns.values():

                        target_cell = ws.cell(
                            row_num,
                            int(col),
                        )

                        if not isinstance(
                            target_cell,
                            MergedCell,
                        ):

                            target_cell.value = None

            continue

        # ----------------------------------------------------
        # INSERE LINHAS ADICIONAIS
        # ----------------------------------------------------
        #
        # Temos, no OSO:
        #
        # linha 10 = modelo A
        # linha 11 = modelo B
        #
        # Portanto já existem 2 posições.
        #
        # 2 registros  -> insere 0
        # 3 registros  -> insere 1
        # 58 registros -> insere 56
        # ----------------------------------------------------

        additional_rows = max(
            0,
            record_count - model_count,
        )

        if additional_rows > 0:

            insertion_row = (
                start_row
                + model_count
            )

            ws.insert_rows(
                insertion_row,
                amount=additional_rows,
            )

            # Atualiza área de impressão.
            _update_print_area(
                ws,
                additional_rows,
                insertion_row,
            )

        # ----------------------------------------------------
        # APLICA ESTILO ZEBRADO
        # ----------------------------------------------------
        #
        # Registro 1 -> modelo 10
        # Registro 2 -> modelo 11
        # Registro 3 -> modelo 10
        # Registro 4 -> modelo 11
        # ...
        # ----------------------------------------------------

        for index in range(
            record_count
        ):

            target_row = (
                start_row
                + index
            )

            template_index = (
                index
                % model_count
            )

            source_row = (
                template_rows[
                    template_index
                ]
            )

            # As linhas originais 10 e 11
            # já possuem a formatação correta.
            #
            # Só precisamos copiar estilo
            # quando chegarmos às novas linhas.
            if target_row not in template_rows:

                _copy_row_style(
                    ws,
                    source_row,
                    target_row,
                    ws.max_column,
                )

        # ----------------------------------------------------
        # PREENCHE OS REGISTROS
        # ----------------------------------------------------

        for (
            index,
            record,
        ) in enumerate(
            records
        ):

            row_num = (
                start_row
                + index
            )

            for (
                field,
                col,
            ) in columns.items():

                col_num = int(
                    col
                )

                target_cell = ws.cell(
                    row_num,
                    col_num,
                )

                if isinstance(
                    target_cell,
                    MergedCell,
                ):
                    continue

                value = record.get(
                    field
                )

                target_cell.value = value

    # ========================================================
    # 7. SUBSTITUI VARIÁVEIS SIMPLES
    # ========================================================

    _replace_variables(
        ws,
        variables,
    )

    # ========================================================
    # 8. ALTERA NOME DA ABA
    # ========================================================

    _replace_sheet_title(
        ws,
        variables,
    )

    # ========================================================
    # 9. SALVA EM MEMÓRIA
    # ========================================================

    out = BytesIO()

    wb.save(
        out
    )

    out.seek(0)

    return out
