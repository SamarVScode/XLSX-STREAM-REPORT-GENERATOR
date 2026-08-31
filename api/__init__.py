"""
API package for ei_stream_server.
"""
from .routes import router as routes_router
from .ui import router as ui_router

__all__ = ["routes_router", "ui_router"]
