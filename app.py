import sys
import logging
from pathlib import Path
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

SERVER_DIR = Path(__file__).resolve().parent
if str(SERVER_DIR) not in sys.path:
    sys.path.insert(0, str(SERVER_DIR))

from core.jobs import recover_jobs_from_disk
from api.routes import router as api_router
from api.ui import router as ui_router

log = logging.getLogger("ei_stream_server.app")

@asynccontextmanager
async def lifespan(app: FastAPI):
    log.info("Booting EI Stream Server: recovering job state from disk...")
    recover_jobs_from_disk()
    yield
    log.info("Shutting down EI Stream Server.")

def create_app() -> FastAPI:
    app = FastAPI(
        title="EI Stream Server",
        description="High-Speed In-Flight Streaming & Async Processing for 350MB+ Large Reports",
        version="3.0.0",
        lifespan=lifespan
    )

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    app.include_router(api_router)
    app.include_router(ui_router)

    return app
