import logging
from typing import Optional
from fastapi import Header, HTTPException
from config.settings import REQUIRED_API_KEY

log = logging.getLogger("ei_stream_server.auth")

def verify_api_key(x_api_key: Optional[str] = Header(None)) -> str:
    if REQUIRED_API_KEY:
        if not x_api_key or x_api_key != REQUIRED_API_KEY:
            log.warning(f"Unauthorized API request. Key provided: {x_api_key}")
            raise HTTPException(status_code=403, detail="Invalid or missing X-API-KEY header")
    return x_api_key or ""
