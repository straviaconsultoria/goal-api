import os
from fastapi import Header, HTTPException


def verify_api_key(x_api_key: str | None = Header(default=None)):
    expected = os.getenv("API_KEY")
    if not expected:
        raise HTTPException(status_code=500, detail="API_KEY não configurada no servidor")
    if x_api_key != expected:
        raise HTTPException(status_code=401, detail="API key inválida")
