from __future__ import annotations

from typing import List, Tuple

# xy → u'v' 변환
def xy_to_uv(x: float, y: float) -> Tuple[float, float]:
    denom = -2 * x + 12 * y + 3
    if denom == 0:
        return (0.0, 0.0)
    u = 4 * x / denom
    v = 9 * y / denom
    return (u, v)


def _xy_list_to_uv(primaries_xy: List[Tuple[float, float]]) -> List[Tuple[float, float]]:
    return [xy_to_uv(x, y) for x, y in primaries_xy]


# DCI-P3 primaries: R, G, B (xy)
_DCI_P3_XY = [(0.680, 0.320), (0.265, 0.690), (0.150, 0.060)]
# BT.2020 primaries: R, G, B (xy)
_BT2020_XY = [(0.708, 0.292), (0.170, 0.797), (0.131, 0.046)]

DCI_P3_UV: List[Tuple[float, float]] = _xy_list_to_uv(_DCI_P3_XY)
BT2020_UV: List[Tuple[float, float]] = _xy_list_to_uv(_BT2020_XY)


def _polygon_area(polygon: List[Tuple[float, float]]) -> float:
    """Shoelace 공식으로 polygon 면적 계산."""
    n = len(polygon)
    if n < 3:
        return 0.0
    area = 0.0
    for i in range(n):
        j = (i + 1) % n
        area += polygon[i][0] * polygon[j][1]
        area -= polygon[j][0] * polygon[i][1]
    return abs(area) / 2.0


def _clip_polygon_by_half_plane(
    polygon: List[Tuple[float, float]],
    edge_a: Tuple[float, float],
    edge_b: Tuple[float, float],
) -> List[Tuple[float, float]]:
    """Sutherland-Hodgman: edge_a → edge_b 왼쪽(내부) 클리핑."""
    if not polygon:
        return []

    def inside(p: Tuple[float, float]) -> bool:
        return (
            (edge_b[0] - edge_a[0]) * (p[1] - edge_a[1])
            - (edge_b[1] - edge_a[1]) * (p[0] - edge_a[0])
        ) >= 0

    def intersect(p1: Tuple[float, float], p2: Tuple[float, float]) -> Tuple[float, float]:
        dx_c = edge_b[0] - edge_a[0]
        dy_c = edge_b[1] - edge_a[1]
        dx_p = p2[0] - p1[0]
        dy_p = p2[1] - p1[1]
        denom = dx_c * dy_p - dy_c * dx_p
        if abs(denom) < 1e-12:
            return p1
        t = ((p1[0] - edge_a[0]) * dy_c - (p1[1] - edge_a[1]) * dx_c) / denom
        return (p1[0] + t * dx_p, p1[1] + t * dy_p)

    output: List[Tuple[float, float]] = []
    for i in range(len(polygon)):
        current = polygon[i]
        previous = polygon[i - 1]
        if inside(current):
            if not inside(previous):
                output.append(intersect(previous, current))
            output.append(current)
        elif inside(previous):
            output.append(intersect(previous, current))

    return output


def _polygon_intersection(
    subject: List[Tuple[float, float]],
    clip: List[Tuple[float, float]],
) -> List[Tuple[float, float]]:
    """Sutherland-Hodgman 알고리즘으로 두 polygon의 교집합 반환."""
    output = list(subject)
    n = len(clip)
    for i in range(n):
        if not output:
            break
        output = _clip_polygon_by_half_plane(output, clip[i], clip[(i + 1) % n])
    return output


def calc_gamut_stats(
    r_uv: Tuple[float, float],
    g_uv: Tuple[float, float],
    b_uv: Tuple[float, float],
) -> dict:
    """
    반환: dci_overlap(%), bt2020_overlap(%), meas_area, dci_area, bt2020_area
    """
    meas_poly = [r_uv, g_uv, b_uv]
    meas_area = _polygon_area(meas_poly)

    dci_area = _polygon_area(DCI_P3_UV)
    bt2020_area = _polygon_area(BT2020_UV)

    dci_inter = _polygon_intersection(meas_poly, DCI_P3_UV)
    dci_inter_area = _polygon_area(dci_inter) if dci_inter else 0.0

    bt2020_inter = _polygon_intersection(meas_poly, BT2020_UV)
    bt2020_inter_area = _polygon_area(bt2020_inter) if bt2020_inter else 0.0

    dci_overlap = (dci_inter_area / dci_area * 100.0) if dci_area > 0 else 0.0
    bt2020_overlap = (bt2020_inter_area / bt2020_area * 100.0) if bt2020_area > 0 else 0.0

    return {
        "dci_overlap": round(dci_overlap, 2),
        "bt2020_overlap": round(bt2020_overlap, 2),
        "meas_area": meas_area,
        "dci_area": dci_area,
        "bt2020_area": bt2020_area,
    }
