from io import BytesIO

from fastapi import FastAPI, Depends
from fastapi.responses import StreamingResponse

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from app.core.security import verify_api_key
from app.core.registry import load_report_config, list_catalog
from app.core.generator import generate
from app.core.pdf_generator import generate_pdf


app = FastAPI(title="Excel Report API", version="2.0.0")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/test-pdf", dependencies=[Depends(verify_api_key)])
def test_pdf():

    output = BytesIO()

    pdf = canvas.Canvas(output, pagesize=A4)
    pdf.setTitle("Teste PDF - GoalBus")

    pdf.drawString(72, 800, "PDF OK - GoalBus Report API")
    pdf.drawString(72, 780, "ReportLab funcionando corretamente.")

    pdf.save()

    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/pdf",
        headers={
            "Content-Disposition": 'attachment; filename="teste_goalbus.pdf"'
        },
    )


@app.get("/catalog", dependencies=[Depends(verify_api_key)])
def catalog():

    return list_catalog()


@app.post("/reports/{tenant}/{report}", dependencies=[Depends(verify_api_key)])
def generate_report(tenant: str, report: str, payload: dict):

    template_path, config = load_report_config(tenant, report)

    output = generate(template_path, config, payload)

    filename = config.get("output_filename", f"{tenant}_{report}.xlsx")

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post(
    "/reports/{tenant}/{report}/pdf",
    dependencies=[Depends(verify_api_key)],
)
def generate_pdf_report(
    tenant: str,
    report: str,
    payload: dict,
):

    # Valida que tenant/relatório existem na configuração atual.
    load_report_config(tenant, report)

    output = generate_pdf(
        tenant,
        report,
        payload,
    )

    filename = f"{report}.pdf"

    return StreamingResponse(
        output,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        },
    )
