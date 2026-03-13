"""Backend detection and environment configuration."""

from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import structlog

log = structlog.get_logger()

# Paths
LISP_DIR = Path(__file__).resolve().parent.parent.parent / "lisp-code"
IPC_DIR = Path(os.environ.get("AUTOCAD_MCP_IPC_DIR", "C:/temp"))

# Backend selection: auto | file_ipc | com | ezdxf
BACKEND_DEFAULT = "auto"

# IPC timeout (seconds), clamped to [1, 300]
IPC_TIMEOUT = max(1.0, min(300.0, float(os.environ.get("AUTOCAD_MCP_IPC_TIMEOUT", "30.0"))))

# Screenshot
ONLY_TEXT_FEEDBACK = os.environ.get("AUTOCAD_MCP_ONLY_TEXT", "").lower() in ("1", "true", "yes")

# Win32 availability
WIN32_AVAILABLE = sys.platform == "win32"

# COM backend: which CAD to target
COM_CAD_TYPE = os.environ.get("AUTOCAD_MCP_CAD_TYPE", "autocad").strip().lower()

# ---------------------------------------------------------------------------
# Multi-CAD configuration
# ---------------------------------------------------------------------------

CAD_PROG_IDS: dict[str, str] = {
    "autocad": "AutoCAD.Application",
    "zwcad": "ZWCAD.Application",
    "gcad": "GCAD.Application",
    "bricscad": "BricscadApp.AcadApplication",
}

CAD_STARTUP_WAIT: dict[str, float] = {
    "autocad": 20.0,
    "zwcad": 15.0,
    "gcad": 15.0,
    "bricscad": 15.0,
}


@dataclass
class CADConfig:
    """Configuration for a specific CAD application."""
    cad_type: str
    prog_id: str
    startup_wait_time: float


def get_cad_config(cad_type: str | None = None) -> CADConfig:
    """Get configuration for a specific CAD application."""
    cad = (cad_type or COM_CAD_TYPE).lower()
    return CADConfig(
        cad_type=cad,
        prog_id=CAD_PROG_IDS.get(cad, "AutoCAD.Application"),
        startup_wait_time=CAD_STARTUP_WAIT.get(cad, 20.0),
    )


def get_supported_cads() -> list[str]:
    """Return list of supported CAD application types."""
    return list(CAD_PROG_IDS.keys())


def _load_config_json() -> dict[str, Any]:
    """Load config.json if present, returning empty dict otherwise."""
    search_paths = [
        Path("config.json"),
        Path(__file__).resolve().parent.parent.parent / "config" / "config.json",
        Path(__file__).resolve().parent.parent.parent.parent / "config.json",
    ]
    for p in search_paths:
        if p.exists():
            try:
                with p.open("r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
    return {}


def _current_backend_env() -> str:
    """Read backend selection from env with normalization."""
    return os.environ.get("AUTOCAD_MCP_BACKEND", BACKEND_DEFAULT).strip().lower()


def _is_wsl() -> bool:
    """Detect WSL Linux runtime."""
    if os.environ.get("WSL_INTEROP"):
        return True
    try:
        return "microsoft" in os.uname().release.lower()
    except AttributeError:
        return False


def _write_debug_snapshot(backend_env: str):
    """Optionally write backend detection debug information."""
    debug_file = os.environ.get("AUTOCAD_MCP_DEBUG_DETECT_FILE", "").strip()
    if not debug_file:
        return
    try:
        debug_path = Path(debug_file)
        debug_path.parent.mkdir(parents=True, exist_ok=True)
        with debug_path.open("w", encoding="utf-8") as f:
            f.write(f"sys.platform={sys.platform}\n")
            f.write(f"WIN32_AVAILABLE={WIN32_AVAILABLE}\n")
            f.write(f"BACKEND_ENV={backend_env}\n")
            f.write(f"python={sys.executable}\n")
    except Exception:
        pass


def _try_com_connection() -> bool:
    """Check if a COM-based CAD application is running."""
    if not WIN32_AVAILABLE:
        return False
    try:
        import win32com.client
        cad_cfg = get_cad_config()
        app = win32com.client.GetActiveObject(cad_cfg.prog_id)
        return app is not None
    except Exception:
        return False


def detect_backend() -> str:
    """Return the backend name to use: 'file_ipc', 'com', or 'ezdxf'.

    Priority (in auto mode):
      1. file_ipc — if AutoCAD LT window detected
      2. com — if COM-based CAD application is running
      3. ezdxf — headless fallback
    """
    backend_env = _current_backend_env()
    _write_debug_snapshot(backend_env)

    # Explicit selection
    if backend_env == "ezdxf":
        return "ezdxf"

    if backend_env == "com":
        if not WIN32_AVAILABLE:
            raise RuntimeError(
                "AUTOCAD_MCP_BACKEND=com requires Windows with pywin32."
            )
        return "com"

    if backend_env in ("auto", "file_ipc"):
        if WIN32_AVAILABLE:
            try:
                from autocad_mcp.backends.file_ipc import find_autocad_window
                hwnd = find_autocad_window()
                if hwnd:
                    log.info("autocad_window_found", hwnd=hwnd)
                    return "file_ipc"
                elif backend_env == "file_ipc":
                    raise RuntimeError(
                        "AUTOCAD_MCP_BACKEND=file_ipc but no AutoCAD window found. "
                        "Start AutoCAD LT and open a .dwg file."
                    )
            except ImportError:
                if backend_env == "file_ipc":
                    raise RuntimeError(
                        "AUTOCAD_MCP_BACKEND=file_ipc requires pywin32. "
                        "Install with: pip install pywin32"
                    )
                log.info("win32_deps_missing_fallback")
        elif backend_env == "file_ipc":
            raise RuntimeError(
                "AUTOCAD_MCP_BACKEND=file_ipc requires Windows. "
                "Use AUTOCAD_MCP_BACKEND=ezdxf for headless mode."
            )
        elif _is_wsl():
            log.info(
                "wsl_linux_python_fallback_ezdxf",
                platform=sys.platform,
                python=sys.executable,
                hint="Launch MCP with Windows python.exe for File IPC backend.",
            )

    # Auto mode: try COM before falling back to ezdxf
    if backend_env == "auto" and WIN32_AVAILABLE:
        if _try_com_connection():
            log.info("com_cad_detected", cad_type=COM_CAD_TYPE)
            return "com"

    log.info("using_ezdxf_backend")
    return "ezdxf"
