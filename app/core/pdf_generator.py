from io import BytesIO
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
)


# ============================================================
# GERADOR PDF
# ============================================================

def generate_pdf(
    tenant: str,
    report: str,
    payload: dict,
) -> BytesIO:

    output = BytesIO()

    # ========================================================
    # DOCUMENTO
    # ========================================================

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=8 * mm,
        leftMargin=8 * mm,
        topMargin=8 * mm,
        bottomMargin=8 * mm,
        title=f"{report} - {tenant}",
    )

    elements = []

    styles = getSampleStyleSheet()

    # ========================================================
    # CAMINHOS
    # ========================================================

    base_dir = Path(__file__).resolve().parents[2]

    tenant_dir = (
        base_dir
        / "tenants"
        / tenant
    )

    logo_path = (
        tenant_dir
        / "assets"
        / "logo.png"
    )

    # ========================================================
    # VARIÁVEIS
    # ========================================================

    variables = (
        payload.get(
            "variables",
            {}
        )
        or {}
    )

    linha = variables.get(
        "LINHA",
        ""
    )

    # ========================================================
    # CABEÇALHO
    # ========================================================

    header_data = []

    title = Paragraph(
        f"<b>OSO - Tabela Resumo Operação da Linha {linha}</b>",
        styles["Heading2"],
    )

    if logo_path.exists():

        logo = Image(
            str(logo_path),
            width=70 * mm,
            height=13 * mm,
        )

        header_data.append(
            [
                title,
                logo,
            ]
        )

    else:

        header_data.append(
            [
                title,
                "",
            ]
        )

    header = Table(
        header_data,
        colWidths=[
            180 * mm,
            90 * mm,
        ],
    )

    header.setStyle(
        TableStyle(
            [
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),
                (
                    "ALIGN",
                    (1, 0),
                    (1, 0),
                    "RIGHT",
                ),
            ]
        )
    )

    elements.append(header)

    elements.append(
        Spacer(
            1,
            5 * mm,
        )
    )

    # ========================================================
    # REGISTROS
    # ========================================================

    records = (
        payload
        .get("sections", {})
        .get("registros_api", [])
        or []
    )

    # ========================================================
    # TABELA
    # ========================================================

    table_data = [
        [
            "Bloco",
            "Itinerário",
            "Início",
            "Percurso",
            "Parada",
            "Fim",
            "Headway",

            "Bloco",
            "Itinerário",
            "Início",
            "Percurso",
            "Parada",
            "Fim",
            "Headway",

            "Início",
            "Fim",
            "SOC Ant.",

            "Início",
            "Fim",
            "SOC Após",
            "T. Recarga",
        ]
    ]

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

    for record in records:

        row = []

        for field in fields:

            value = record.get(
                field,
                "",
            )

            if value is None:
                value = ""

            row.append(
                str(value)
            )

        table_data.append(row)

    # ========================================================
    # LARGURAS
    # ========================================================

    available_width = (
        landscape(A4)[0]
        - doc.leftMargin
        - doc.rightMargin
    )

    column_width = (
        available_width
        / len(fields)
    )

    table = Table(
        table_data,
        colWidths=[
            column_width
        ] * len(fields),
        repeatRows=1,
    )

    # ========================================================
    # ESTILO
    # ========================================================

    table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#D9EAF7"),
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold",
                ),
                (
                    "FONTSIZE",
                    (0, 0),
                    (-1, -1),
                    5.5,
                ),
                (
                    "ALIGN",
                    (0, 0),
                    (-1, -1),
                    "CENTER",
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.25,
                    colors.HexColor("#B7B7B7"),
                ),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [
                        colors.white,
                        colors.HexColor("#F2F2F2"),
                    ],
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
            ]
        )
    )

    elements.append(table)

    # ========================================================
    # GERA PDF
    # ========================================================

    doc.build(elements)

    output.seek(0)

    return output
