# -*- coding: utf-8 -*-
"""כפתור כתום – מוצר 1 עם הדפס משתנה בצד אחד או יותר."""
from __future__ import annotations

import os
import time
import pythoncom
from typing import Any, Dict, List, Optional

from illustrator_ops import (
    apply_extra_colors,
    apply_text_overrides_in_layer,
    clean_layout,
    delete_information_layer,
    delete_side_assets,
    open_and_color_template,
    place_and_simulate_print,
    run_jsx,
    set_order_number_in_simulation,
    setup_variable_print_slots,
    update_size_label,
    variable_layer_and_artboard,
)

SIDE_KEYS = ["front", "back", "right_sleeve", "left_sleeve"]
SIDE_TO_PREFIX = {"front": "F", "back": "B", "right_sleeve": "RS", "left_sleeve": "LS"}
SIDE_DEFAULTS = {
    "front": {"prefix": "F", "category": "A4", "label": "size_Front", "heb": "קדמי"},
    "back": {"prefix": "B", "category": "A4", "label": "size_Back", "heb": "אחורי"},
    "right_sleeve": {"prefix": "RS", "category": "Sleeve", "label": "size_Right_Sleeve", "heb": "שרוול ימין"},
    "left_sleeve": {"prefix": "LS", "category": "Sleeve", "label": "size_Left_Sleeve", "heb": "שרוול שמאל"},
}


def is_variable_side(loc_data: Optional[Dict[str, Any]]) -> bool:
    if not loc_data:
        return False
    return bool(loc_data.get("variable_print")) and bool(loc_data.get("variants"))


def normalize_variants(loc_data: Dict[str, Any]) -> List[Dict[str, Any]]:
    variants = loc_data.get("variants") or []
    if not isinstance(variants, list) or not variants:
        raise ValueError("variable_print requires a non-empty variants array")
    out = []
    for i, v in enumerate(variants):
        if not isinstance(v, dict):
            continue
        idx = v.get("index")
        if idx is None:
            idx = i + 1
        out.append({**v, "index": int(idx)})
    out.sort(key=lambda x: x["index"])
    if not out:
        raise ValueError("variable_print variants are empty after normalization")
    return out


def download_variable_product_files(product: Dict[str, Any], order_prefix: str) -> None:
    from main import download_image

    for side in SIDE_KEYS:
        loc = product.get(side, {}) or {}
        if is_variable_side(loc):
            normalized = normalize_variants(loc)
            for v in normalized:
                if v.get("file_url") and not v.get("file"):
                    path = download_image(
                        v["file_url"],
                        f"{order_prefix}_var_{side}_{v['index']}",
                    )
                    if path:
                        v["file"] = path
            loc["variants"] = normalized
            product[side] = loc
        elif _side_is_active(loc) and loc.get("file_url") and not loc.get("file"):
            path = download_image(loc["file_url"], f"{order_prefix}_0_{side}")
            if path:
                loc["file"] = path
                product[side] = loc


def _side_is_active(loc_data: Dict[str, Any]) -> bool:
    if not loc_data:
        return False
    return bool(loc_data.get("exists"))


def _enrich_side_data(side: str, data: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    from main import enrich_side_data

    return enrich_side_data(side, data)


def vectorize_variant(
    variant: Dict[str, Any],
    side: str,
    folder: str,
    side_defaults: Dict[str, Any],
) -> Optional[str]:
    from main import API_ID, API_SECRET, vec_single

    if not variant.get("file") or not os.path.exists(variant["file"]):
        return None
    d = {
        "exists": True,
        "file": variant["file"],
        "no_vectorization": variant.get(
            "no_vectorization", side_defaults.get("no_vectorization", False)
        ),
        "prefix": side_defaults.get("prefix") or SIDE_TO_PREFIX.get(side, "F"),
    }
    return vec_single(d, folder, API_ID, API_SECRET)


def process_variable_product_to_temp(
    order: Dict[str, Any],
    folder: str,
    is_wholesale: bool = False,
    order_id: Optional[str] = None,
) -> Optional[str]:
    """יוצר קובץ AI יחיד למוצר 1 עם הדפסים משתנים."""
    from main import (
        TEMP_AI_DIR,
        TEMPLATES,
        get_hex_smart,
        normalize_split_product_color,
        resolve_print_color,
        side_is_active,
        vec_single,
    )

    pythoncom.CoInitialize()
    doc = None
    app = None
    try:
        prod = order.get("product_type", "Shirt")
        print(f"\n>> [Variable] Processing Product 1: {prod}")
        t_path = TEMPLATES.get(prod)
        if not t_path or not os.path.exists(t_path):
            return None

        variable_sides = [s for s in SIDE_KEYS if is_variable_side(order.get(s, {}))]
        if not variable_sides:
            raise ValueError("No variable_print side found in product")

        col_raw = normalize_split_product_color(order.get("product_color_hebrew", ""))
        parts = [p.strip() for p in col_raw.split("-")] if "-" in col_raw else [col_raw]
        h1 = get_hex_smart(parts[0])
        h2 = get_hex_smart(parts[1]) if len(parts) >= 2 else h1
        is_split = len(parts) >= 2
        doc, app = open_and_color_template(t_path, h1, h2, is_split, prod)

        extra = order.get("extra_colors_hebrew", [])
        extra_data = []
        for name in extra:
            p = [x.strip() for x in name.split("-")]
            pair = [get_hex_smart(c) for c in p[:2] if get_hex_smart(c) != "ORIGINAL"]
            if pair:
                extra_data.append(pair)
        apply_extra_colors(app, extra_data)

        for side in variable_sides:
            loc = _enrich_side_data(side, order.get(side, {}) or {})
            variants = normalize_variants(loc)
            setup_variable_print_slots(app, side, len(variants))
            order[side] = loc

        for side in SIDE_KEYS:
            if side in variable_sides:
                continue
            loc = _enrich_side_data(side, order.get(side, {}) or {})
            if not side_is_active(loc):
                continue
            if not loc.get("file") or not os.path.exists(loc["file"]):
                print(f"   > Skip {side}: missing file")
                continue
            svg = vec_single(loc, folder, __import__("main", fromlist=["API_ID"]).API_ID, __import__("main", fromlist=["API_SECRET"]).API_SECRET)
            if not svg:
                continue
            prefix = loc.get("prefix") or SIDE_TO_PREFIX.get(side, "F")
            is_r = loc.get("no_vectorization", False)
            fc = resolve_print_color(loc.get("req_color_hebrew"), h1)
            cp = fc if fc != "#FFFFFF" else "#000000"
            w = place_and_simulate_print(
                doc, app, svg, prefix, loc.get("category", "A4"), cp, fc, is_r
            )
            label = loc.get("label") or SIDE_DEFAULTS.get(side, {}).get("label", "")
            heb = loc.get("heb") or SIDE_DEFAULTS.get(side, {}).get("heb", "")
            if w > 0 and label:
                update_size_label(doc, app, label, w, heb)

        for side in variable_sides:
            loc = order[side]
            variants = normalize_variants(loc)
            prefix = loc.get("prefix") or SIDE_TO_PREFIX.get(side, "F")
            side_req = loc.get("req_color_hebrew")
            cat = loc.get("category", "A4")
            label = loc.get("label") or SIDE_DEFAULTS.get(side, {}).get("label", "")
            heb = loc.get("heb") or SIDE_DEFAULTS.get(side, {}).get("heb", "")

            for vi, variant in enumerate(variants):
                v_idx = variant["index"]
                svg = vectorize_variant(variant, side, folder, loc)
                if not svg:
                    print(f"   > Skip {side} variant {v_idx}: vectorization failed")
                    continue
                layer_name, artboard_name = variable_layer_and_artboard(side, v_idx)
                req = variant.get("req_color_hebrew") or side_req
                is_r = variant.get("no_vectorization", loc.get("no_vectorization", False))
                fc = resolve_print_color(req, h1)
                cp = fc if fc != "#FFFFFF" else "#000000"
                with_sim = vi == 0
                w = place_and_simulate_print(
                    doc,
                    app,
                    svg,
                    prefix,
                    cat,
                    cp,
                    fc,
                    is_r,
                    layer_name=layer_name,
                    artboard_name=artboard_name,
                    skip_simulation=not with_sim,
                    should_update_size_label=with_sim,
                )
                if w > 0 and with_sim and label:
                    update_size_label(doc, app, label, w, heb)
                overrides = variant.get("text_overrides") or variant.get("textOverrides") or {}
                if overrides:
                    apply_text_overrides_in_layer(app, layer_name, overrides)
                print(f"   > {side} variant {v_idx}: placed on {layer_name} (sim={with_sim})")

        rs_active = side_is_active(order.get("right_sleeve", {})) or "right_sleeve" in variable_sides
        ls_active = side_is_active(order.get("left_sleeve", {})) or "left_sleeve" in variable_sides
        if not rs_active and not ls_active:
            delete_side_assets(doc, app, "Print_Sleeves", "size_Right_Sleeve")
            run_jsx(app, "try{app.activeDocument.textFrames.getByName('size_Left_Sleeve').remove();}catch(e){}")
        elif not rs_active and "right_sleeve" not in variable_sides:
            run_jsx(app, "try{app.activeDocument.textFrames.getByName('size_Right_Sleeve').remove();}catch(e){}")
        elif not ls_active and "left_sleeve" not in variable_sides:
            run_jsx(app, "try{app.activeDocument.textFrames.getByName('size_Left_Sleeve').remove();}catch(e){}")

        if not side_is_active(order.get("front", {})) and "front" not in variable_sides:
            delete_side_assets(doc, app, "Print_Front", "size_Front")
        if not side_is_active(order.get("back", {})) and "back" not in variable_sides:
            delete_side_assets(doc, app, "Print_Back", "size_Back")

        clean_layout(app)
        if order_id:
            set_order_number_in_simulation(app, order_id)
        if is_wholesale:
            delete_information_layer(app)

        out_path = os.path.join(TEMP_AI_DIR, "temp_0.ai")
        doc.SaveAs(out_path)
        doc.Close(2)
        time.sleep(0.5)
        print(f"   > Saved variable product: {out_path}")
        return out_path
    except Exception as e:
        print(f"Error processing variable product: {e}")
        import traceback

        traceback.print_exc()
        if doc:
            try:
                doc.Close(2)
            except Exception:
                pass
        return None
    finally:
        pythoncom.CoUninitialize()
