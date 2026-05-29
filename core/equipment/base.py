from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class PatternInfo:
    type: str                  # e.g. "window", "full_field", "crosshair"
    apl_pct: float             # Average Picture Level 0-100
    width_pct: float           # window width as % of screen
    height_pct: float          # window height as % of screen
    color: str                 # "white", "black", "red", "green", "blue"
    is_hdr: bool = False


@dataclass
class PatternConfig:
    type: str                  # "full_field" | "window" | "crosshair"
    color: str                 # primary color name or "rgb(r,g,b)"
    r: int = 255
    g: int = 255
    b: int = 255
    width_pct: float = 100.0   # for window patterns
    height_pct: float = 100.0
    bg_r: int = 0
    bg_g: int = 0
    bg_b: int = 0
    bit_mode: int = 8
    is_hdr: bool = False


@dataclass
class MeasureResult:
    timestamp_ms: int
    Lv: float
    x: float
    y: float
    u_prime: float
    v_prime: float
    X: float
    Y: float
    Z: float
    cct: float = 0.0   # Correlated Color Temperature (K) — from device
    duv: float = 0.0   # Distance from Planckian locus — from device
    pattern_info: PatternInfo = field(default_factory=lambda: PatternInfo(
        type="unknown", apl_pct=0.0, width_pct=0.0, height_pct=0.0, color="unknown"
    ))


class MeterBase(ABC):
    """Abstract base for all colorimeter/luminance meter drivers."""

    @abstractmethod
    def connect(self, port: str) -> None:
        """Open serial connection to meter on given port."""
        ...

    @abstractmethod
    def disconnect(self) -> None:
        """Close serial connection."""
        ...

    @abstractmethod
    def measure(self) -> MeasureResult:
        """Trigger one measurement and return a fully populated MeasureResult."""
        ...

    @property
    @abstractmethod
    def is_connected(self) -> bool:
        ...


class GeneratorBase(ABC):
    """Abstract base for all pattern generator drivers."""

    @abstractmethod
    def connect(self, port: str) -> None:
        """Open serial/USB connection to generator."""
        ...

    @abstractmethod
    def disconnect(self) -> None:
        """Close connection."""
        ...

    @abstractmethod
    def set_pattern(self, cfg: PatternConfig) -> None:
        """Output the pattern described by cfg."""
        ...

    @abstractmethod
    def set_hdr(self, enabled: bool) -> None:
        """Switch to HDR output mode (4K 60p YCbCr 4:2:2, SMPTE ST2084) or back to SDR."""
        ...

    @abstractmethod
    def set_sdr(self) -> None:
        """Revert generator output to SDR mode."""
        ...

    def reset(self) -> None:
        """Recover from freeze: re-enter terminal mode and reload base program.

        Default is no-op; override in hardware drivers.
        """

    @property
    @abstractmethod
    def is_connected(self) -> bool:
        ...
