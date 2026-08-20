from fastapi import FastAPI, Depends
from fastapi.responses import StreamingResponse
from app.core.security import verify_api_key
from app.core.registry import load_report_config, list_catalog
from app.core.generator import generate

app = FastAPI(title="Excel Report API", version="2.0.0")

@app.get("/health")
def health():
    return {"status": "ok"}

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
