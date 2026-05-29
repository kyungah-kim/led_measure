"""Flask routes: REST API + Server-Sent Events for live measurement progress."""
from __future__ import annotations

import dataclasses
import json
import queue
import threading
from typing import Any, Generator

from flask import Blueprint, Response, current_app, jsonify, request, render_template, send_file

from core.engine import MeasurementEngine
from core.export import ExcelExporter


def _to_json_safe(obj: Any) -> Any:
    """Recursively convert dataclasses / lists / dicts to JSON-serialisable types."""
    if dataclasses.is_dataclass(obj) and not isinstance(obj, type):
        return dataclasses.asdict(obj)  # type: ignore[arg-type]
    if isinstance(obj, dict):
        return {k: _to_json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_to_json_safe(v) for v in obj]
    if isinstance(obj, (int, float, str, bool)) or obj is None:
        return obj
    return str(obj)

bp = Blueprint("api", __name__, url_prefix="/api")

ui_bp = Blueprint("ui", __name__)

@ui_bp.route("/")
def index():
    return render_template("index.html")

# Thread-safe queue for SSE progress events
_progress_queue: queue.Queue[str] = queue.Queue(maxsize=500)

# Most-recent full results, keyed by sequence name (for export)
_last_results: dict[str, Any] = {}


def _get_engine() -> MeasurementEngine:
    return current_app.extensions["engine"]


def _sse_event(data: dict[str, Any]) -> str:
    return f"data: {json.dumps(data)}\n\n"


# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------

@bp.route("/connect", methods=["POST"])
def connect() -> Response:
    """Connect meter and generator.

    JSON body:
        {
            "brand": "...", "model_name": "...",
            "meter_port": "COM3", "meter_model": "CA-410",
            "gen_port": "COM4", "gen_model": "VG-876"
        }
    """
    body = request.get_json(force=True)
    engine = _get_engine()

    engine.brand = body.get("brand", engine.brand)
    engine.model_name = body.get("model_name", engine.model_name)

    errors: list[str] = []

    meter_port = body.get("meter_port")
    if meter_port:
        try:
            engine.connect_meter(meter_port, body.get("meter_model", "CA-410"))
        except Exception as exc:
            errors.append(f"Meter: {exc}")

    gen_port = body.get("gen_port")
    if gen_port:
        try:
            engine.connect_generator(gen_port, body.get("gen_model", "VG-876"))
        except Exception as exc:
            errors.append(f"Generator: {exc}")

    if errors:
        return jsonify({"ok": False, "errors": errors}), 500

    return jsonify({"ok": True, "ready": engine.is_ready})


@bp.route("/connect/mock", methods=["POST"])
def connect_mock() -> Response:
    """Connect Mock meter and generator for testing without physical hardware."""
    from core.equipment.mock import MockMeter, MockGenerator
    body = request.get_json(force=True) or {}
    engine = _get_engine()
    engine.brand = body.get("brand") or engine.brand or "Samsung"
    engine.model_name = body.get("model_name") or engine.model_name or "QN65S95D"
    m = MockMeter(); m.connect("MOCK")
    g = MockGenerator(); g.connect("MOCK")
    engine.meter = m
    engine.generator = g
    return jsonify({"ok": True, "ready": engine.is_ready})


# ---------------------------------------------------------------------------
# Status
# ---------------------------------------------------------------------------

@bp.route("/status", methods=["GET"])
def status() -> Response:
    engine = _get_engine()
    return jsonify({
        "ready": engine.is_ready,
        "brand": engine.brand,
        "model_name": engine.model_name,
        "meter_connected": engine.meter is not None and engine.meter.is_connected,
        "generator_connected": engine.generator is not None and engine.generator.is_connected,
    })


@bp.route("/hdr", methods=["POST"])
def set_hdr() -> Response:
    """Switch the connected generator between SDR and HDR immediately."""
    engine = _get_engine()
    gen = engine.generator
    if gen is None or not gen.is_connected:
        return jsonify({"ok": False, "error": "Generator is not connected"}), 400

    body = request.get_json(force=True) or {}
    enabled = bool(body.get("enabled", False))
    try:
        gen.set_hdr(enabled)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500

    return jsonify({"ok": True, "hdr": enabled})


# ---------------------------------------------------------------------------
# Run sequence
# ---------------------------------------------------------------------------

@bp.route("/run/<seq_name>", methods=["POST"])
def run_sequence(seq_name: str) -> Response:
    """Start a named measurement sequence in a background thread.

    Sequence parameters are passed as JSON body.  Progress is streamed via
    GET /api/progress (SSE).
    """
    engine = _get_engine()
    if not engine.is_ready:
        return jsonify({"ok": False, "error": "Engine not ready"}), 400

    params: dict[str, Any] = request.get_json(force=True) or {}

    def _progress(step: str, pct: float, data: Any) -> None:
        try:
            serialisable_data: Any = _to_json_safe(data)
        except Exception:
            serialisable_data = None

        event = _sse_event({"step": step, "progress": pct, "data": serialisable_data})
        try:
            _progress_queue.put_nowait(event)
        except queue.Full:
            pass  # drop if consumer is too slow

    engine.on_progress = _progress

    def _run() -> None:
        try:
            result = engine.run_sequence(seq_name, **params)
            _last_results[seq_name] = result
            _progress_queue.put_nowait(_sse_event({"step": seq_name, "progress": 1.0, "done": True}))
        except Exception as exc:
            _progress_queue.put_nowait(_sse_event({"error": str(exc)}))

    thread = threading.Thread(target=_run, daemon=True)
    thread.start()

    return jsonify({"ok": True, "seq_name": seq_name})


# ---------------------------------------------------------------------------
# Stop
# ---------------------------------------------------------------------------

@bp.route("/stop", methods=["POST"])
def stop_sequence() -> Response:
    """Signal the currently running sequence to stop."""
    _get_engine().stop_sequence()
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# SSE progress stream
# ---------------------------------------------------------------------------

@bp.route("/progress", methods=["GET"])
def progress_stream() -> Response:
    """Server-Sent Events endpoint.  Connect once; receive events until done."""

    def _generate() -> Generator[str, None, None]:
        yield _sse_event({"connected": True})
        while True:
            try:
                event = _progress_queue.get(timeout=30)
                yield event
                if '"done": true' in event or '"error":' in event:
                    break
            except queue.Empty:
                # Keep-alive comment to prevent proxy timeouts
                yield ": keep-alive\n\n"

    return Response(
        _generate(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

@bp.route("/export/<export_type>", methods=["GET"])
def export_results(export_type: str) -> Response:
    """Download an Excel file for the requested result type.

    export_type: "lum_swing" | "lum_loading" | "gamut" | "contrast"
    """
    engine = _get_engine()
    exporter = ExcelExporter()
    use_avg = request.args.get("use_avg", "true").lower() == "true"

    try:
        if export_type == "lum_swing":
            data = _last_results.get("lum_swing")
            if not data:
                return jsonify({"error": "No lum_swing results available"}), 404
            path = exporter.export_lum_swing(data, engine.brand, engine.model_name)

        elif export_type == "lum_loading":
            data = _last_results.get("lum_loading")
            if not data:
                return jsonify({"error": "No lum_loading results available"}), 404
            path = exporter.export_lum_loading(data, engine.brand, engine.model_name, use_avg=use_avg)

        elif export_type == "gamut":
            data = _last_results.get("gamut")
            if not data:
                return jsonify({"error": "No gamut results available"}), 404
            path = exporter.export_gamut(data, engine.brand, engine.model_name)

        elif export_type == "contrast":
            data = _last_results.get("contrast")
            if not data:
                return jsonify({"error": "No contrast results available"}), 404
            path = exporter.export_contrast(data, engine.brand, engine.model_name)

        else:
            return jsonify({"error": f"Unknown export type: {export_type}"}), 400

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    return send_file(path, as_attachment=True, download_name=path.split("/")[-1])
