from __future__ import annotations

from eia_gen.models.case import Case
from eia_gen.models.fields import QuantityField, TextField


ITEM_ID_ALIASES: dict[str, str] = {
    "NE-TOPO": "NAT_TG",
    "NE-ECO": "NAT_ECO",
    "NE-WATER": "NAT_WATER",
    "LE-AIR": "LIFE_AIR",
    "LE-NOISE": "LIFE_NOISE",
    "LE-ODOR": "LIFE_ODOR",
    "SE-LANDUSE": "SOC_LANDUSE",
    "SE-LANDSCAPE": "SOC_LANDSCAPE",
    "SE-POP": "SOC_POP",
}


def canonicalize_item_id(item_id: str) -> str:
    return ITEM_ID_ALIASES.get(item_id, item_id)


def canonicalize_case(case: Case) -> Case:
    # scoping_matrix
    for item in case.scoping_matrix:
        item.item_id = canonicalize_item_id(item.item_id)

    # mitigation.related_impacts
    for m in case.mitigation.measures:
        m.related_impacts = [canonicalize_item_id(i) for i in m.related_impacts]

    # project_overview.area: compute derived fields (best-effort)
    area = case.project_overview.area

    # (1) total_area_m2: if missing, sum parcels
    if area.total_area_m2.v is None:
        total = 0.0
        srcs: list[str] = []
        has_any = False
        for p in area.parcels:
            if p.area_m2.v is None:
                continue
            has_any = True
            total += float(p.area_m2.v)
            for sid in (p.area_m2.src or []) + (p.jibun.src or []) + (p.zoning.src or []):
                s = (sid or "").strip()
                if s and s not in srcs:
                    srcs.append(s)
        if has_any:
            area.total_area_m2 = QuantityField(v=total, u=area.total_area_m2.u or "m2", src=srcs)

    # (2) zoning_area_m2: if empty, aggregate by parcel zoning
    if not area.zoning_area_m2:
        grouped: dict[str, tuple[float, list[str]]] = {}
        for p in area.parcels:
            zoning = (p.zoning.t or "").strip()
            if not zoning:
                continue
            if p.area_m2.v is None:
                continue
            cur_v, cur_srcs = grouped.get(zoning, (0.0, []))
            cur_v += float(p.area_m2.v)
            for sid in (p.area_m2.src or []) + (p.zoning.src or []) + (p.jibun.src or []):
                s = (sid or "").strip()
                if s and s not in cur_srcs:
                    cur_srcs.append(s)
            grouped[zoning] = (cur_v, cur_srcs)
        if grouped:
            area.zoning_area_m2 = {
                k: QuantityField(v=v, u="m2", src=srcs) for k, (v, srcs) in grouped.items()
            }

    # baseline.landuse_landscape: best-effort summaries from parcels
    ll = case.baseline.landuse_landscape
    if ll.current_landcover_summary.is_empty():
        # Prefer parcel land category distribution when available.
        from collections import Counter

        landcats = [(p.land_category.t or "").strip() for p in area.parcels]
        landcats = [x for x in landcats if x]
        counts = Counter(landcats)
        total_m2 = area.total_area_m2.v
        parcel_count = len(area.parcels)

        summary = ""
        if counts and parcel_count and total_m2 is not None:
            items = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
            joined = ", ".join([f"{k} {v}필지" for k, v in items])
            summary = (
                f"대상지 지번({parcel_count}필지)은 지목 기준 {joined}로 구성되며, "
                f"총 면적은 {int(total_m2):,}m²이다."
            )
        elif counts and parcel_count:
            items = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
            joined = ", ".join([f"{k} {v}필지" for k, v in items])
            summary = f"대상지 지번({parcel_count}필지)은 지목 기준 {joined}로 구성된다."
        elif parcel_count and total_m2 is not None:
            summary = f"대상지 총 면적은 {int(total_m2):,}m²이다."

        if summary:
            srcs: list[str] = []
            for p in area.parcels:
                for sid in (p.land_category.src or []) + (p.area_m2.src or []) + (p.jibun.src or []):
                    s = (sid or "").strip()
                    if s and s not in srcs:
                        srcs.append(s)
            for sid in area.total_area_m2.src or []:
                s = (sid or "").strip()
                if s and s not in srcs:
                    srcs.append(s)
            ll.current_landcover_summary = TextField(t=summary, src=srcs)

    # baseline.population_traffic: best-effort nearest_village from address/admin
    pt = case.baseline.population_traffic
    if pt.nearest_village.is_empty():
        addr = (case.project_overview.location.address.t or "").strip()
        addr_src = list(case.project_overview.location.address.src or [])
        if addr.endswith("일원"):
            addr = addr[: -len("일원")].strip()
        if addr:
            pt.nearest_village = TextField(t=addr, src=addr_src)

    return case
