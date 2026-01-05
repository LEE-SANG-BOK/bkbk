from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Color:
    r: int
    g: int
    b: int
    a: int = 255

    def rgba(self) -> tuple[int, int, int, int]:
        return (self.r, self.g, self.b, self.a)


# Sample-style palette (roughly matching the user's drone overlay screenshots)
BOUNDARY_OUTLINE = Color(255, 255, 255, 255)
BOUNDARY_FILL = Color(255, 255, 255, 32)

FACILITY_POND = Color(53, 120, 255, 140)  # blue-ish
FACILITY_WALKWAY = Color(240, 240, 240, 200)  # light
FACILITY_ROAD = Color(80, 80, 80, 120)  # dark gray
FACILITY_BUILDING = Color(220, 180, 120, 180)  # sand
FACILITY_OTHER = Color(180, 180, 180, 120)

DRAINAGE_LINE = Color(45, 140, 255, 220)

LABEL_BG = Color(60, 60, 60, 140)
LABEL_FG = (255, 255, 255, 255)

LEGEND_BG = Color(255, 255, 255, 180)
LEGEND_BORDER = Color(220, 220, 220, 255)

