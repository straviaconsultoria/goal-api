from pathlib import Path
import yaml
from fastapi import HTTPException

ROOT = Path(__file__).resolve().parents[2]
TENANTS_DIR = ROOT / "tenants"


def safe_name(value: str) -> str:
    if not value or not all(c.isalnum() or c in "-_" for c in value):
        raise HTTPException(status_code=400, detail="Nome de tenant/relatório inválido")
    return value


def report_dir(tenant: str, report: str) -> Path:
    tenant = safe_name(tenant)
    report = safe_name(report)
    path = TENANTS_DIR / tenant / "reports" / report
    if not path.is_dir():
        raise HTTPException(status_code=404, detail="Tenant ou relatório não encontrado")
    return path


def load_report_config(tenant: str, report: str) -> tuple[Path, dict]:
    path = report_dir(tenant, report)
    config_path = path / "config.yaml"
    if not config_path.exists():
        raise HTTPException(status_code=500, detail="config.yaml não encontrado para o relatório")
    with config_path.open("r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}
    template_name = config.get("template")
    if not template_name:
        raise HTTPException(status_code=500, detail="Campo template ausente no config.yaml")
    template_path = path / "templates" / template_name
    if not template_path.exists():
        raise HTTPException(status_code=500, detail=f"Template não encontrado: {template_name}")
    return template_path, config


def list_catalog() -> dict:
    result = {}
    if not TENANTS_DIR.exists():
        return result
    for tenant in sorted(p for p in TENANTS_DIR.iterdir() if p.is_dir()):
        reports_dir = tenant / "reports"
        if not reports_dir.is_dir():
            continue
        result[tenant.name] = sorted(p.name for p in reports_dir.iterdir() if p.is_dir())
    return result
