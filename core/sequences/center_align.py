from __future__ import annotations

from typing import TYPE_CHECKING, Callable


if TYPE_CHECKING:
    from ..engine import MeasurementEngine


class CenterAlignSequence:
    """Outputs the 네모-abc crosshair pattern for physical centre alignment.

    The sequence just outputs the pattern and then yields control back to the
    UI.  The UI is responsible for displaying [OK] and calling confirm() once
    the user has physically aligned the meter.
    """

    def __init__(self, engine: "MeasurementEngine") -> None:
        self._engine = engine
        self._confirmed = False

    def start(self, on_ready: Callable[[], None] | None = None) -> None:
        """Output the alignment pattern and notify the UI that it is displayed."""
        gen = self._engine.generator
        if gen is None or not gen.is_connected:
            raise RuntimeError("Generator is not connected")

        gen.show_center_align()

        if on_ready:
            on_ready()

    def confirm(self) -> None:
        """Called by the UI when the user presses [OK]."""
        self._confirmed = True

    @property
    def is_confirmed(self) -> bool:
        return self._confirmed
