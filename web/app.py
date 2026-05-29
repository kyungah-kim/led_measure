"""Flask application factory."""
from __future__ import annotations

import os
import sys

# Allow `from core.xxx import ...` when running from project root
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from flask import Flask

from core.engine import MeasurementEngine


def create_app(brand: str = "", model_name: str = "") -> Flask:
    """Create and configure the Flask application.

    Parameters
    ----------
    brand:      Pre-populated brand string (can be updated via /api/connect).
    model_name: Pre-populated model name.
    """
    app = Flask(__name__)
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-in-prod")

    # Single engine instance attached to the app for the process lifetime
    engine = MeasurementEngine(brand=brand, model_name=model_name)
    app.extensions["engine"] = engine

    from web.routes import bp, ui_bp
    app.register_blueprint(bp)
    app.register_blueprint(ui_bp)

    return app


if __name__ == "__main__":
    flask_app = create_app()
    flask_app.run(debug=True, host="0.0.0.0", port=5000, threaded=True)
