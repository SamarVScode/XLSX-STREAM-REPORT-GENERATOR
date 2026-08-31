"""
Configuration package for ei_stream_server.
"""
from .settings import (
    REQUIRED_API_KEY,
    CACHE_DIR,
    CACHE_TTL,
    CACHE_MAX_AGE,
    MAX_CONCURRENT_JOBS,
)
from .dc_config import (
    ALLOWED_SOURCE_DCS,
    ALLOWED_DCS_SET,
    ALLOWED_DCS_SET_LOWER,
)

__all__ = [
    "REQUIRED_API_KEY",
    "CACHE_DIR",
    "CACHE_TTL",
    "CACHE_MAX_AGE",
    "MAX_CONCURRENT_JOBS",
    "ALLOWED_SOURCE_DCS",
    "ALLOWED_DCS_SET",
    "ALLOWED_DCS_SET_LOWER",
]
