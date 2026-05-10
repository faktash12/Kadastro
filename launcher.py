from __future__ import annotations

import os
import sys
import webbrowser
from pathlib import Path


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def main() -> None:
    app_path = resource_path("app.py")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "true")

    try:
        from streamlit.web import bootstrap
        from streamlit import config
    except Exception as exc:  # pragma: no cover
        raise SystemExit(f"Streamlit başlatılamadı: {exc}") from exc

    port = int(os.environ.get("KADASTRO_PORT", "8501"))
    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    config.set_option("server.port", port, "KadastroHarc launcher")
    config.set_option("server.headless", True, "KadastroHarc launcher")
    config.set_option("browser.gatherUsageStats", False, "KadastroHarc launcher")
    config.set_option("global.developmentMode", False, "KadastroHarc launcher")
    url = f"http://localhost:{port}"
    webbrowser.open(url, new=2)
    bootstrap.run(
        str(app_path),
        False,
        [],
        {
            "server.port": port,
            "server.headless": True,
            "browser.gatherUsageStats": False,
        },
    )


if __name__ == "__main__":
    main()
