"""Flask routes: REST API + Server-Sent Events for live measurement progress."""
from __future__ import annotations

import dataclasses
import json
import os
import queue
import threading
from typing import Any, Generator

from flask import Blueprint, Response, current_app, jsonify, request, render_template, send_file

from core.engine import MeasurementEngine
from core.export import ExcelExporter


def _save_all_session(engine: MeasurementEngine) -> str:
    """세션 데이터를 {brand}_{model}_all.xlsx 로 자동 저장. 경로 반환."""
    if not engine.auto_save_dir:
        return ""
    brand = engine.brand or "brand"
    model = engine.model_name or "model"
    path = os.path.join(engine.auto_save_dir, f"{brand}_{model}_all.xlsx")
    try:
        ExcelExporter().export_all_session(
            brand=brand, model=model,
            session_swing=engine.session_swing,
            session_loading=engine.session_loading,
            session_gamut=engine.session_gamut,
            session_contrast=engine.session_contrast,
            file_path=path,
        )
    except Exception as e:
        print(f"[auto_save] 오류: {e}")
        return ""
    return path


def _update_session(engine: MeasurementEngine, seq_name: str,
                    params: dict[str, Any], result: Any) -> None:
    """시퀀스 완료 후 engine 세션 데이터를 갱신한다."""
    mode = "HDR" if params.get("is_hdr") else "SDR"
    case = params.get("case", "")
    if seq_name == "lum_swing" and isinstance(result, dict):
        for k, v in result.items():
            engine.session_swing[f"{mode}_{k}"] = list(v)
    elif seq_name == "lum_loading" and isinstance(result, dict):
        engine.session_loading[f"{mode}_{case}"] = dict(result)
    elif seq_name == "gamut" and isinstance(result, dict):
        engine.session_gamut[mode] = dict(result)
    elif seq_name == "contrast" and isinstance(result, dict):
        engine.session_contrast[mode] = dict(result)


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


# ---------------------------------------------------------------------------
# Mock connect / Disconnect
# ---------------------------------------------------------------------------

@bp.route("/mock_connect", methods=["POST"])
def mock_connect() -> Response:
    """Connect mock meter and generator for offline testing."""
    engine = _get_engine()
    from core.equipment.mock import MockMeter, MockGenerator
    engine.meter = MockMeter()
    engine.generator = MockGenerator()
    return jsonify({"ok": True, "ready": engine.is_ready})


@bp.route("/disconnect", methods=["POST"])
def disconnect_all() -> Response:
    """Disconnect all equipment."""
    engine = _get_engine()
    engine.disconnect_all()
    engine.meter = None
    engine.generator = None
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# Port scan
# ---------------------------------------------------------------------------

@bp.route("/ports", methods=["GET"])
def list_ports() -> Response:
    """Return a sorted list of available serial port device names."""
    try:
        import serial.tools.list_ports
        ports = sorted(p.device for p in serial.tools.list_ports.comports())
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc), "ports": []}), 500
    return jsonify({"ok": True, "ports": ports})


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

    # 이전 측정의 잔류 이벤트 제거 — 새 SSE 연결이 구 데이터를 받지 않도록
    while not _progress_queue.empty():
        try:
            _progress_queue.get_nowait()
        except queue.Empty:
            break

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
            _update_session(engine, seq_name, params, result)
            save_path = _save_all_session(engine)
            _progress_queue.put_nowait(_sse_event({
                "step": seq_name, "progress": 1.0, "done": True,
                "auto_save_path": save_path,
            }))
        except Exception as exc:
            _progress_queue.put_nowait(_sse_event({"error": str(exc)}))

    thread = threading.Thread(target=_run, daemon=True)
    thread.start()

    return jsonify({"ok": True, "seq_name": seq_name})


# ---------------------------------------------------------------------------
# Auto-save directory
# ---------------------------------------------------------------------------

@bp.route("/auto_save_dir", methods=["GET"])
def get_auto_save_dir() -> Response:
    engine = _get_engine()
    return jsonify({"ok": True, "path": engine.auto_save_dir})


@bp.route("/auto_save_dir", methods=["POST"])
def set_auto_save_dir():
    body = request.get_json(force=True) or {}
    path = body.get("path", "").strip()
    if path:
        try:
            os.makedirs(path, exist_ok=True)
        except Exception as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
    _get_engine().auto_save_dir = path
    return jsonify({"ok": True, "path": path})


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
