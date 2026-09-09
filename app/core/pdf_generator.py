from io import BytesIO
from pathlib import Path
from typing import Any

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


# ============================================================
# CORES
# ============================================================

COLOR_BLUE_DARK = colors.HexColor("#1F4E78")
COLOR_BLUE = colors.HexColor("#5B9BD5")
COLOR_BLUE_LIGHT = colors.HexColor("#D9EAF7")
COLOR_BLUE_LIGHTER = colors.HexColor("#EAF3F8")

COLOR_GRAY = colors.HexColor("#D9E1F2")
COLOR_GRAY_LIGHT = colors.HexColor("#F2F2F2")
COLOR_BORDER = colors.HexColor("#A6A6A6")

COLOR_ENTRADA = colors.HexColor("#C4D79B")
COLOR_FECHAMENTO = colors.HexColor("#9FC5E8")
COLOR_ESPECIAL = colors.HexColor("#FFFF66")
COLOR_RECARGA = colors.HexColor("#FABF8F")

COLOR_WHITE = colors.white
COLOR_BLACK = colors.black


# ============================================================
# ESTILOS DE TEXTO
# ============================================================

STYLE_TITLE = ParagraphStyle(
    name="OSO_Title",
    fontName="Helvetica-Bold",
    fontSize=11,
    leading=13,
    textColor=COLOR_BLACK,
    alignment=TA_LEFT,
)

STYLE_SUBTITLE = ParagraphStyle(
    name="OSO_Subtitle",
    fontName="Helvetica-Bold",
    fontSize=8,
    leading=10,
    textColor=COLOR_BLACK,
    alignment=TA_LEFT,
)

STYLE_HEADER = ParagraphStyle(
    name="OSO_Header",
    fontName="Helvetica-Bold",
    fontSize=5.2,
    leading=6,
    textColor=COLOR_BLACK,
    alignment=TA_CENTER,
)

STYLE_HEADER_WHITE = ParagraphStyle(
    name="OSO_Header_White",
    fontName="Helvetica-Bold",
    fontSize=6,
    leading=7,
    textColor=COLOR_WHITE,
    alignment=TA_CENTER,
)

STYLE_CELL = ParagraphStyle(
    name="OSO_Cell",
    fontName="Helvetica",
    fontSize=4.8,
    leading=5.6,
    textColor=COLOR_BLACK,
    alignment=TA_CENTER,
)

STYLE_CELL_LEFT = ParagraphStyle(
    name="OSO_Cell_Left",
    fontName="Helvetica",
    fontSize=4.6,
    leading=5.4,
    textColor=COLOR_BLACK,
    alignment=TA_LEFT,
)


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def _safe_value(value: Any) -> str:
    if value is None:
        return ""

    return str(value)


def _paragraph(
    value: Any,
    style=STYLE_CELL,
):
    text = _safe_value(value)

    # Evita interpretação acidental de caracteres HTML
    text = (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )

    return Paragraph(
        text,
        style,
    )


def _get_displacement_color(
    displacement_type: Any,
):
    if displacement_type is None:
        return None

    displacement_type = (
        str(displacement_type)
        .strip()
        .upper()
    )

    if displacement_type == "ENTRADA":
        return COLOR_ENTRADA

    if displacement_type == "FECHAMENTO":
        return COLOR_FECHAMENTO

    if displacement_type == "ESPECIAL":
        return COLOR_ESPECIAL

    if displacement_type == "DESLOCAMENTO_RECARGA":
        return COLOR_RECARGA

    return None


# ============================================================
# CABEÇALHO DO DOCUMENTO
# ============================================================

def _build_document_header(
    tenant_dir: Path,
    variables: dict,
):

    linha = _safe_value(
        variables.get("LINHA")
    )

    itinerario = _safe_value(
        variables.get("ITINERARIO")
    )

    titulo_oso = _safe_value(
        variables.get("TITULO_OSO")
    )

    tipo_dia = _safe_value(
        variables.get("TIPO_DIA")
    )

    logo_path = (
        tenant_dir
        / "assets"
        / "logo.png"
    )

    # ========================================================
    # BLOCO ESQUERDO
    # ========================================================

    left_content = []

    linha_text = linha

    if itinerario:
        linha_text = (
            f"{linha} - {itinerario}"
            if linha
            else itinerario
        )

    left_content.append(
        _paragraph(
            linha_text,
            STYLE_TITLE,
        )
    )

    if titulo_oso:
        left_content.append(
            _paragraph(
                titulo_oso,
                STYLE_SUBTITLE,
            )
        )

    left_content.append(
        _paragraph(
            "TABELA RESUMO OPERAÇÃO DA LINHA",
            STYLE_SUBTITLE,
        )
    )

    if tipo_dia:
        left_content.append(
            _paragraph(
                tipo_dia,
                STYLE_SUBTITLE,
            )
        )

    # ========================================================
    # LOGO
    # ========================================================

    logo = ""

    if logo_path.exists():

        logo = Image(
            str(logo_path),
        )

        # Mantém proporção da imagem.
        max_width = 70 * mm
        max_height = 15 * mm

        original_width = float(
            logo.imageWidth
        )

        original_height = float(
            logo.imageHeight
        )

        scale = min(
            max_width / original_width,
            max_height / original_height,
        )

        logo.drawWidth = (
            original_width
            * scale
        )

        logo.drawHeight = (
            original_height
            * scale
        )

    header = Table(
        [
            [
                left_content,
                logo,
            ]
        ],
        colWidths=[
            190 * mm,
            85 * mm,
        ],
    )

    header.setStyle(
        TableStyle(
            [
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP",
                ),
                (
                    "ALIGN",
                    (1, 0),
                    (1, 0),
                    "RIGHT",
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),
            ]
        )
    )

    return header


# ============================================================
# CABEÇALHO DA TABELA
# ============================================================

def _build_table_header(
    variables: dict,
):

    route_1 = _safe_value(
        variables.get("ROUTE_1")
    )

    route_2 = _safe_value(
        variables.get("ROUTE_2")
    )

    # ========================================================
    # LINHA 1
    #
    # PARTIDAS = A:N
    # PLANO DE RECARGA = O:U
    # ========================================================

    row_1 = [
        _paragraph(
            "PARTIDAS",
            STYLE_HEADER_WHITE,
        ),
    ]

    row_1.extend(
        [""] * 13
    )

    row_1.append(
        _paragraph(
            "PLANO DE RECARGA",
            STYLE_HEADER_WHITE,
        )
    )

    row_1.extend(
        [""] * 6
    )

    # ========================================================
    # LINHA 2
    #
    # Sentido 1 = A:G
    # Sentido 2 = H:N
    # Deslocamento = O:Q
    # Recarga = R:U
    # ========================================================

    row_2 = [
        _paragraph(
            route_1 or "SENTIDO 1",
            STYLE_HEADER,
        ),
    ]

    row_2.extend(
        [""] * 6
    )

    row_2.append(
        _paragraph(
            route_2 or "SENTIDO 2",
            STYLE_HEADER,
        )
    )

    row_2.extend(
        [""] * 6
    )

    row_2.append(
        _paragraph(
            "DESLOCAMENTO RECARGA",
            STYLE_HEADER,
        )
    )

    row_2.extend(
        [""] * 2
    )

    row_2.append(
        _paragraph(
            "RECARGA PROGRAMADA",
            STYLE_HEADER,
        )
    )

    row_2.extend(
        [""] * 3
    )

    # ========================================================
    # LINHA 3
    # ========================================================

    headers = [
        "Bloco",
        "Itinerário",
        "Horário<br/>Início",
        "Tempo<br/>Percurso",
        "Parada<br/>Operacional",
        "Horário<br/>Fim",
        "Headway",

        "Bloco",
        "Itinerário",
        "Horário<br/>Início",
        "Tempo<br/>Percurso",
        "Parada<br/>Operacional",
        "Horário<br/>Fim",
        "Headway",

        "Horário<br/>Início",
        "Horário<br/>Fim",
        "SOC<br/>Anterior",

        "Horário<br/>Início",
        "Horário<br/>Fim",
        "SOC<br/>Após",
        "Tempo Total<br/>Recarga",
    ]

    row_3 = [
        Paragraph(
            header,
            STYLE_HEADER,
        )
        for header in headers
    ]

    return [
        row_1,
        row_2,
        row_3,
    ]


# ============================================================
# LINHAS DE DADOS
# ============================================================

def _build_data_rows(
    records: list,
):

    fields = [
        "1_BLOCO",
        "1_ITINERARIO",
        "1_HORARIO_INICIO",
        "1_TEMPO_PERCURSO",
        "1_PARADA_OPERACIONAL",
        "1_HORARIO_FIM",
        "1_HEADWAY",

        "2_BLOCO",
        "2_ITINERARIO",
        "2_HORARIO_INICIO",
        "2_TEMPO_PERCURSO",
        "2_PARADA_OPERACIONAL",
        "2_HORARIO_FIM",
        "2_HEADWAY",

        "3_HORARIO_INICIO",
        "3_HORARIO_FIM",
        "3_SOC_ANTERIOR_RECARGA",

        "4_HORARIO_INICIO",
        "4_HORARIO_FIM",
        "4_SOC_APOS_RECARGA",
        "4_TEMPO_TOTAL_RECARGA",
    ]

    rows = []

    for record in records:

        row = []

        for index, field in enumerate(
            fields
        ):

            # Itinerários ficam alinhados
            # à esquerda para melhorar leitura.
            if index in (1, 8):

                row.append(
                    _paragraph(
                        record.get(field),
                        STYLE_CELL_LEFT,
                    )
                )

            else:

                row.append(
                    _paragraph(
                        record.get(field),
                        STYLE_CELL,
                    )
                )

        rows.append(row)

    return rows


# ============================================================
# LARGURA DAS COLUNAS
# ============================================================

def _get_column_widths():
    """
    O OSO possui 21 colunas.

    Os itinerários (B e I) recebem mais espaço.
    As demais colunas são compactas.

    Soma aproximada compatível com A4 landscape
    utilizando margens reduzidas.
    """

    return [
        13 * mm,     # A  1_BLOCO
        31 * mm,     # B  1_ITINERARIO
        11 * mm,     # C  INICIO
        11 * mm,     # D  PERCURSO
        12 * mm,     # E  PARADA
        11 * mm,     # F  FIM
        10 * mm,     # G  HEADWAY

        13 * mm,     # H  2_BLOCO
        31 * mm,     # I  2_ITINERARIO
        11 * mm,     # J  INICIO
        11 * mm,     # K  PERCURSO
        12 * mm,     # L  PARADA
        11 * mm,     # M  FIM
        10 * mm,     # N  HEADWAY

        10 * mm,     # O  INICIO DESLOC.
        10 * mm,     # P  FIM DESLOC.
        10 * mm,     # Q  SOC ANTERIOR

        10 * mm,     # R  INICIO RECARGA
        10 * mm,     # S  FIM RECARGA
        10 * mm,     # T  SOC APÓS
        12 * mm,     # U  TEMPO RECARGA
    ]


# ============================================================
# ESTILO DA TABELA
# ============================================================

def _build_table_style(
    records: list,
):

    commands = [
        # ----------------------------------------------------
        # MERGES - CABEÇALHO PRINCIPAL
        # ----------------------------------------------------

        (
            "SPAN",
            (0, 0),
            (13, 0),
        ),

        (
            "SPAN",
            (14, 0),
            (20, 0),
        ),

        # ----------------------------------------------------
        # MERGES - SUBGRUPOS
        # ----------------------------------------------------

        (
            "SPAN",
            (0, 1),
            (6, 1),
        ),

        (
            "SPAN",
            (7, 1),
            (13, 1),
        ),

        (
            "SPAN",
            (14, 1),
            (16, 1),
        ),

        (
            "SPAN",
            (17, 1),
            (20, 1),
        ),

        # ----------------------------------------------------
        # CORES DO CABEÇALHO
        # ----------------------------------------------------

        (
            "BACKGROUND",
            (0, 0),
            (13, 0),
            COLOR_BLUE_DARK,
        ),

        (
            "BACKGROUND",
            (14, 0),
            (20, 0),
            COLOR_BLUE_DARK,
        ),

        (
            "BACKGROUND",
            (0, 1),
            (6, 1),
            COLOR_BLUE_LIGHT,
        ),

        (
            "BACKGROUND",
            (7, 1),
            (13, 1),
            COLOR_BLUE_LIGHT,
        ),

        (
            "BACKGROUND",
            (14, 1),
            (16, 1),
            COLOR_GRAY,
        ),

        (
            "BACKGROUND",
            (17, 1),
            (20, 1),
            COLOR_GRAY,
        ),

        (
            "BACKGROUND",
            (0, 2),
            (20, 2),
            COLOR_BLUE_LIGHTER,
        ),

        # ----------------------------------------------------
        # GERAL
        # ----------------------------------------------------

        (
            "VALIGN",
            (0, 0),
            (-1, -1),
            "MIDDLE",
        ),

        (
            "ALIGN",
            (0, 0),
            (-1, -1),
            "CENTER",
        ),

        (
            "LEFTPADDING",
            (0, 0),
            (-1, -1),
            1.5,
        ),

        (
            "RIGHTPADDING",
            (0, 0),
            (-1, -1),
            1.5,
        ),

        (
            "TOPPADDING",
            (0, 0),
            (-1, -1),
            2,
        ),

        (
            "BOTTOMPADDING",
            (0, 0),
            (-1, -1),
            2,
        ),

        (
            "GRID",
            (0, 0),
            (-1, -1),
            0.25,
            COLOR_BORDER,
        ),

        # ----------------------------------------------------
        # DIVISÕES DOS QUADRANTES
        # ----------------------------------------------------

        (
            "LINEBEFORE",
            (7, 0),
            (7, -1),
            0.8,
            COLOR_BLACK,
        ),

        (
            "LINEBEFORE",
            (14, 0),
            (14, -1),
            0.8,
            COLOR_BLACK,
        ),

        (
            "LINEBEFORE",
            (17, 0),
            (17, -1),
            0.8,
            COLOR_BLACK,
        ),

        (
            "LINEAFTER",
            (20, 0),
            (20, -1),
            0.8,
            COLOR_BLACK,
        ),
    ]

    # ========================================================
    # ZEBRADO
    #
    # As três primeiras linhas são cabeçalho.
    # ========================================================

    for index in range(
        len(records)
    ):

        table_row = (
            index
            + 3
        )

        if index % 2 == 1:

            commands.append(
                (
                    "BACKGROUND",
                    (0, table_row),
                    (20, table_row),
                    COLOR_GRAY_LIGHT,
                )
            )

    # ========================================================
    # CORES DOS DESLOCAMENTOS
    # ========================================================

    for index, record in enumerate(
        records
    ):

        table_row = (
            index
            + 3
        )

        # ----------------------------------------------------
        # QUADRANTE 1 - A:G
        # ----------------------------------------------------

        fill = _get_displacement_color(
            record.get(
                "1_TIPO_DESLOCAMENTO"
            )
        )

        if fill is not None:

            commands.append(
                (
                    "BACKGROUND",
                    (0, table_row),
                    (6, table_row),
                    fill,
                )
            )

        # ----------------------------------------------------
        # QUADRANTE 2 - H:N
        # ----------------------------------------------------

        fill = _get_displacement_color(
            record.get(
                "2_TIPO_DESLOCAMENTO"
            )
        )

        if fill is not None:

            commands.append(
                (
                    "BACKGROUND",
                    (7, table_row),
                    (13, table_row),
                    fill,
                )
            )

        # ----------------------------------------------------
        # QUADRANTE 3 - O:Q
        # ----------------------------------------------------

        displacement_3 = record.get(
            "3_TIPO_DESLOCAMENTO"
        )

        fill = _get_displacement_color(
            displacement_3
        )

        if fill is not None:

            commands.append(
                (
                    "BACKGROUND",
                    (14, table_row),
                    (16, table_row),
                    fill,
                )
            )

        # ----------------------------------------------------
        # QUADRANTE 4 - R:U
        # ----------------------------------------------------

        fill = _get_displacement_color(
            record.get(
                "4_TIPO_DESLOCAMENTO"
            )
        )

        if fill is not None:

            commands.append(
                (
                    "BACKGROUND",
                    (17, table_row),
                    (20, table_row),
                    fill,
                )
            )

        # ----------------------------------------------------
        # REGRA ESPECIAL:
        #
        # DESLOCAMENTO_RECARGA no Q3
        # colore H:U, reproduzindo a regra atual do Excel.
        # ----------------------------------------------------

        if (
            displacement_3 is not None
            and str(
                displacement_3
            ).strip().upper()
            == "DESLOCAMENTO_RECARGA"
        ):

            commands.append(
                (
                    "BACKGROUND",
                    (7, table_row),
                    (20, table_row),
                    COLOR_RECARGA,
                )
            )

    return TableStyle(
        commands
    )


# ============================================================
# RODAPÉ / NÚMERO DE PÁGINA
# ============================================================

def _draw_page_footer(
    canvas,
    doc,
):

    canvas.saveState()

    page_width, _ = landscape(
        A4
    )

    page_number = (
        canvas.getPageNumber()
    )

    text = (
        f"Página {page_number}"
    )

    canvas.setFont(
        "Helvetica",
        6,
    )

    canvas.setFillColor(
        colors.HexColor(
            "#666666"
        )
    )

    text_width = stringWidth(
        text,
        "Helvetica",
        6,
    )

    canvas.drawString(
        page_width
        - doc.rightMargin
        - text_width,
        5 * mm,
        text,
    )

    canvas.restoreState()


# ============================================================
# GERADOR PDF
# ============================================================

def generate_pdf(
    tenant: str,
    report: str,
    payload: dict,
) -> BytesIO:

    # ========================================================
    # VALIDAÇÃO
    # ========================================================

    if not isinstance(
        payload,
        dict,
    ):
        raise ValueError(
            "Payload do PDF deve ser um objeto JSON."
        )

    variables = (
        payload.get(
            "variables",
            {}
        )
        or {}
    )

    sections = (
        payload.get(
            "sections",
            {}
        )
        or {}
    )

    records = (
        sections.get(
            "registros_api",
            []
        )
        or []
    )

    if not records:
        raise ValueError(
            "Nenhum registro encontrado em "
            "sections.registros_api para geração do PDF."
        )

    # ========================================================
    # CAMINHOS
    # ========================================================

    base_dir = (
        Path(__file__)
        .resolve()
        .parents[2]
    )

    tenant_dir = (
        base_dir
        / "tenants"
        / tenant
    )

    if not tenant_dir.exists():
        raise FileNotFoundError(
            f"Tenant não encontrado: {tenant}"
        )

    # ========================================================
    # OUTPUT
    # ========================================================

    output = BytesIO()

    # ========================================================
    # DOCUMENTO
    # ========================================================

    page_width, page_height = (
        landscape(A4)
    )

    left_margin = 5 * mm
    right_margin = 5 * mm
    top_margin = 5 * mm
    bottom_margin = 9 * mm

    doc = BaseDocTemplate(
        output,
        pagesize=landscape(A4),
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
        title=(
            f"OSO - "
            f"{_safe_value(variables.get('LINHA'))}"
        ),
        author="GoalBus",
    )

    frame = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        doc.width,
        doc.height,
        id="oso_frame",
        leftPadding=0,
        rightPadding=0,
        topPadding=0,
        bottomPadding=0,
    )

    page_template = PageTemplate(
        id="oso_page",
        frames=[
            frame
        ],
        onPage=_draw_page_footer,
    )

    doc.addPageTemplates(
        [
            page_template
        ]
    )

    # ========================================================
    # ELEMENTOS
    # ========================================================

    elements = []

    # --------------------------------------------------------
    # CABEÇALHO
    # --------------------------------------------------------

    elements.append(
        _build_document_header(
            tenant_dir,
            variables,
        )
    )

    elements.append(
        Spacer(
            1,
            3 * mm,
        )
    )

    # --------------------------------------------------------
    # TABELA
    # --------------------------------------------------------

    table_data = []

    table_data.extend(
        _build_table_header(
            variables
        )
    )

    table_data.extend(
        _build_data_rows(
            records
        )
    )

    column_widths = (
        _get_column_widths()
    )

    total_width = sum(
        column_widths
    )

    available_width = (
        page_width
        - left_margin
        - right_margin
    )

    # Segurança:
    # se as larguras ultrapassarem a área útil,
    # reduz proporcionalmente.
    if total_width > available_width:

        scale = (
            available_width
            / total_width
        )

        column_widths = [
            width * scale
            for width
            in column_widths
        ]

    table = Table(
        table_data,
        colWidths=column_widths,

        # Repete as três linhas de cabeçalho
        # automaticamente em cada página.
        repeatRows=3,

        hAlign="LEFT",
    )

    table.setStyle(
        _build_table_style(
            records
        )
    )

    elements.append(
        table
    )

    # ========================================================
    # BUILD
    # ========================================================

    doc.build(
        elements
    )

    output.seek(0)

    return output
