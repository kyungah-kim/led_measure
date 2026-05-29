"""CCT and Duv calculation via Robertson's method (CIE 1960 UCS)."""
from __future__ import annotations

import math

# Robertson's isotemperature table
# Each entry: (reciprocal_megakelvin, u_planck, v_planck, slope_t)
# u, v are CIE 1960 UCS coordinates of the Planckian locus
_ROBERTSON: list[tuple[float, float, float, float]] = [
    (0,    0.18006, 0.26352, -0.24341),
    (10,   0.18066, 0.26589, -0.25479),
    (20,   0.18133, 0.26846, -0.26876),
    (30,   0.18208, 0.27119, -0.28539),
    (40,   0.18293, 0.27407, -0.30470),
    (50,   0.18388, 0.27709, -0.32675),
    (60,   0.18494, 0.28021, -0.35156),
    (70,   0.18611, 0.28342, -0.37915),
    (80,   0.18740, 0.28668, -0.40955),
    (90,   0.18880, 0.28997, -0.44278),
    (100,  0.19032, 0.29326, -0.47888),
    (125,  0.19462, 0.30141, -0.58204),
    (150,  0.19962, 0.30921, -0.70471),
    (175,  0.20525, 0.31647, -0.84901),
    (200,  0.21142, 0.32312, -1.01820),
    (225,  0.21807, 0.32909, -1.21479),
    (250,  0.22511, 0.33439, -1.44359),
    (275,  0.23247, 0.33904, -1.70826),
    (300,  0.24010, 0.34308, -2.01199),
    (325,  0.24792, 0.34655, -2.35511),
    (350,  0.25591, 0.34951, -2.73974),
    (375,  0.26400, 0.35200, -3.16618),
    (400,  0.27218, 0.35407, -3.64188),
    (425,  0.28039, 0.35577, -4.15870),
    (450,  0.28863, 0.35714, -4.71582),
    (475,  0.29685, 0.35823, -5.30879),
    (500,  0.30505, 0.35907, -5.92628),
    (525,  0.31320, 0.35968, -6.56118),
    (550,  0.32129, 0.36011, -7.20277),
    (575,  0.32931, 0.36038, -7.84539),
    (600,  0.33724, 0.36051, -8.48514),
]


def xy_to_cct_duv(x: float, y: float) -> tuple[float, float]:
    """Return (CCT in K, Duv) from CIE xy chromaticity.

    CCT: correlated color temperature via Robertson's method.
    Duv: signed distance from the Planckian locus in CIE 1960 UCS.
         Positive = above locus (green shift), negative = below (magenta shift).
    Returns (0.0, 0.0) if the point is outside the table range.
    """
    denom = -2.0 * x + 12.0 * y + 3.0
    if abs(denom) < 1e-12:
        return 0.0, 0.0

    # CIE 1960 UCS
    u = 4.0 * x / denom
    v = 6.0 * y / denom

    prev_di = 0.0
    for i, (rcp_T, u_i, v_i, t_i) in enumerate(_ROBERTSON):
        di = (v - v_i - t_i * (u - u_i)) / math.sqrt(1.0 + t_i * t_i)
        if i > 0 and di * prev_di < 0:
            # Interpolate between entry i-1 and i
            f = prev_di / (prev_di - di)
            rcp_prev = _ROBERTSON[i - 1][0]
            cct = 1e6 / (rcp_prev + f * (rcp_T - rcp_prev))

            u_p = _ROBERTSON[i - 1][1] + f * (u_i - _ROBERTSON[i - 1][1])
            v_p = _ROBERTSON[i - 1][2] + f * (v_i - _ROBERTSON[i - 1][2])
            duv = math.sqrt((u - u_p) ** 2 + (v - v_p) ** 2)
            if v < v_p:
                duv = -duv

            return round(cct, 1), round(duv, 5)
        prev_di = di

    return 0.0, 0.0
