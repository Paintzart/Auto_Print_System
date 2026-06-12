# -*- coding: utf-8 -*-
"""כפתור כתום – מוצר 1 עם הדפס משתנה בצד אחד או יותר."""
from __future__ import annotations

import os
import shutil
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
    outline_text_in_layers,
    outline_document_text,
    place_and_simulate_print,
    place_variable_template_variant,
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


def is_template_side(loc_data: Optional[Dict[str, Any]]) -> bool:
    if not loc_data:
        return False
    if not loc_data.get("template_mode"):
        return False
    return bool(loc_data.get("template_file") or loc_data.get("template_url"))


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


def _download_template(loc: Dict[str, Any], order_prefix: str, side: str) -> None:
    from main import download_image

    url = loc.get("template_url")
    if not url:
        print(f"   > [{side}] template_mode without template_url")
        return
    if isinstance(url, str) and url.startswith("blob:"):
        print(f"   > [{side}] blob URL not supported: {url[:80]}")
        return
    if loc.get("template_file") and os.path.exists(loc["template_file"]):
        return
    print(f"   > [{side}] downloading template: {str(url)[:120]}")
    path = download_image(url, f"{order_prefix}_tpl_{side}")
    if path and not path.lower().endswith(".ai"):
        ai_path = os.path.splitext(path)[0] + ".ai"
        if path != ai_path:
            shutil.move(path, ai_path)
        path = ai_path
    if path and os.path.exists(path):
        loc["template_file"] = path
        print(f"   > [{side}] template saved: {path}")
    else:
        print(f"   > [{side}] template download failed")


def _ensure_template_file(loc: Dict[str, Any], order_prefix: str, side: str) -> str:
    _download_template(loc, order_prefix, side)
    path = loc.get("template_file")
    if path and os.path.exists(path):
        return path
    url = loc.get("template_url") or ""
    raise ValueError(
        f"template_mode on {side} requires a valid template file. "
        f"Could not download template_url={url!r}. "
        f"Send locations.{side}.template_url (HTTPS / local .ai path / base64)."
    )


def normalize_text_overrides(overrides: Optional[dict]) -> dict:
    """מפתחות legacy בלבד. הערך = מחרוזת חופשית (עברית/אנגלית/מספרים/כל פורמט)."""
    if not overrides:
        return {}
    legacy = {
        "TEXT_NAME": "TEXT1",
        "TEXTNAME": "TEXT1",
        "TEXT_NUMBER": "TEXT2",
        "TEXTNUMBER": "TEXT2",
    }
    out: dict = {}
    for raw_key, value in overrides.items():
        if value is None:
            continue
        key = str(raw_key).strip()
        mapped = legacy.get(key.upper())
        out[mapped if mapped else key] = str(value)
    return out


def _resolve_variant_template_path(
    loc: Dict[str, Any], variant: Dict[str, Any], order_prefix: str, side: str
) -> tuple[str, bool]:
    """מחזיר (path, is_shared_template). קובץ variant נפרד עדיף על template_url משותף."""
    vf = variant.get("file")
    if vf and os.path.exists(vf):
        if not vf.lower().endswith(".ai"):
            ai_path = os.path.splitext(vf)[0] + ".ai"
            if os.path.exists(ai_path):
                vf = ai_path
        print(f"   > [{side}] variant {variant.get('index')} template: {os.path.basename(vf)}")
        return vf, False
    shared = _ensure_template_file(loc, order_prefix, side)
    print(f"   > [{side}] variant {variant.get('index')} using shared template: {os.path.basename(shared)}")
    return shared, True


def download_variable_product_files(product: Dict[str, Any], order_prefix: str) -> None:
    from main import download_image

    for side in SIDE_KEYS:
        loc = product.get(side, {}) or {}
        if is_variable_side(loc):
            if is_template_side(loc) or loc.get("template_url"):
                _download_template(loc, order_prefix, side)
            normalized = normalize_variants(loc)
            for v in normalized:
                if v.get("file_url") and not v.get("file"):
                    path = download_image(
                        v["file_url"],
                        f"{order_prefix}_var_{side}_{v['index']}",
                    )
                    if path:
                        v["file"] = path
                image_overrides = v.get("image_overrides") or v.get("imageOverrides") or {}
                if image_overrides:
                    image_files = v.get("image_files") or {}
                    for img_name, img_url in image_overrides.items():
                        if not img_url or image_files.get(img_name):
                            continue
                        img_path = download_image(
                            img_url,
                            f"{order_prefix}_var_{side}_{v['index']}_{img_name}",
                        )
                        if img_path:
                            image_files[str(img_name)] = img_path
                    v["image_files"] = image_files
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


def _should_outline_variable_order(order: Dict[str, Any], variable_sides: List[str]) -> bool:
    for side in variable_sides:
        loc = order.get(side) or {}
        if loc.get("outline_text", True):
            return True
    return False


def _save_order_editable(doc, folder: str, order_id: str) -> Optional[str]:
    if not order_id or not folder:
        return None
    safe_id = str(order_id).strip()
    editable_path = os.path.join(folder, f"{safe_id}_editable.ai")
    counter = 1
    while os.path.exists(editable_path):
        editable_path = os.path.join(folder, f"{safe_id}_editable ({counter}).ai")
        counter += 1
    doc.SaveAs(editable_path)
    print(f"   > Saved editable (live text): {editable_path}")
    return editable_path


def _process_template_variants(
    app,
    doc,
    side: str,
    loc: Dict[str, Any],
    h1: str,
    label: str,
    heb: str,
    order_prefix: str = "",
    defer_outline: bool = False,
) -> None:
    from main import resolve_print_color

    variants = normalize_variants(loc)
    prefix = loc.get("prefix") or SIDE_TO_PREFIX.get(side, "F")
    side_req = loc.get("req_color_hebrew")
    outline_text = False if defer_outline else loc.get("outline_text", True)
    category = loc.get("category", "A4")
    doc_name = doc.Name

    def _clean_variable_slots() -> None:
        cfg = {"front": "Print_Front", "back": "Print_Back", "right_sleeve": "Print_Right_Sleeve", "left_sleeve": "Print_Left_Sleeve"}
        base = cfg.get(side, "Print_Front")
        n = len(variants)
        jsx_clean = f"""
        #target illustrator
        (function() {{
            var doc = app.activeDocument;
            var base = "{base}";
            var total = {int(n)};
            app.executeMenuCommand("unlockAll");
            for (var i = 1; i <= total; i++) {{
                var ln = base + "_" + i;
                try {{
                    var layer = doc.layers.getByName(ln);
                    layer.locked = false;
                    for (var pi = layer.pageItems.length - 1; pi >= 0; pi--) {{
                        var it = layer.pageItems[pi];
                        var nm = it.name || "";
                        if (nm.indexOf("_Box_") !== -1) continue;
                        it.remove();
                    }}
                }} catch(e) {{}}
            }}
        }})();
        """
        run_jsx(app, jsx_clean)

    _clean_variable_slots()

    for vi, variant in enumerate(variants):
        v_idx = variant["index"]
        slot_idx = vi + 1
        layer_name, artboard_name = variable_layer_and_artboard(side, slot_idx)
        req = variant.get("req_color_hebrew") or side_req
        print_hex = resolve_print_color(req, h1) if req else None
        if not print_hex:
            print_hex = resolve_print_color(side_req, h1) if side_req else None
        if not print_hex:
            print_hex = resolve_print_color("", h1)
        text_overrides = normalize_text_overrides(
            variant.get("text_overrides") or variant.get("textOverrides") or {}
        )
        image_files = variant.get("image_files") or {}
        template_path, is_shared = _resolve_variant_template_path(loc, variant, order_prefix, side)
        print(f"   > {side} variant {v_idx} print color: {print_hex} (req={req or side_req})")
        if text_overrides:
            print(f"   > {side} variant {v_idx} text: {text_overrides}")
        elif is_shared:
            print(f"   > {side} variant {v_idx}: no text_overrides — template text unchanged")
        if not is_shared and not text_overrides:
            print(f"   > {side} variant {v_idx}: per-variant file as-is")
        print(
            f"   > {side} variant {v_idx} -> slot {slot_idx} "
            f"({layer_name} / {artboard_name}) text keys: {list(text_overrides.keys())}"
        )
        with_sim = vi == 0
        w = place_variable_template_variant(
            app,
            template_path,
            doc_name,
            layer_name,
            artboard_name,
            prefix,
            text_overrides=text_overrides,
            image_files=image_files,
            outline_text=outline_text,
            skip_simulation=not with_sim,
            sim_hex=print_hex,
            print_hex=print_hex,
            product_doc=doc,
            category=category,
            shared_template=is_shared,
            slot_id=slot_idx,
        )
        if w <= 0:
            print(f"   > Skip {side} variant {v_idx}: template placement failed")
            continue
        if with_sim and label:
            update_size_label(doc, app, label, w, heb)
        print(f"   > {side} variant {v_idx}: template on {layer_name} (sim={with_sim})")


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

        order_prefix = os.path.basename(folder.rstrip("\\/")) or "order"

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
            label = loc.get("label") or SIDE_DEFAULTS.get(side, {}).get("label", "")
            heb = loc.get("heb") or SIDE_DEFAULTS.get(side, {}).get("heb", "")

            if is_template_side(loc):
                _process_template_variants(
                    app, doc, side, loc, h1, label, heb, order_prefix, defer_outline=True
                )
                continue

            variants = normalize_variants(loc)
            prefix = loc.get("prefix") or SIDE_TO_PREFIX.get(side, "F")
            side_req = loc.get("req_color_hebrew")
            cat = loc.get("category", "A4")

            for vi, variant in enumerate(variants):
                v_idx = variant["index"]
                slot_idx = vi + 1
                svg = vectorize_variant(variant, side, folder, loc)
                if not svg:
                    print(f"   > Skip {side} variant {v_idx}: vectorization failed")
                    continue
                layer_name, artboard_name = variable_layer_and_artboard(side, slot_idx)
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

        _save_order_editable(doc, folder, order_id or "")
        if _should_outline_variable_order(order, variable_sides):
            outline_document_text(app)

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
