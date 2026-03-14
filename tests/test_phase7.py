"""Integration test for Phase 7 PowerPoint MCP tools (13 tools)."""
import asyncio
import json
import importlib.util
import os
import tempfile

_server_path = os.path.join(os.path.dirname(__file__), "..", "server.py")
spec = importlib.util.spec_from_file_location("server", _server_path)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


async def call(name, args=None):
    r = await mod.mcp.call_tool(name, args or {})
    if isinstance(r, tuple):
        r = r[0]
    if isinstance(r, list):
        item = r[0]
        text = item.text if hasattr(item, "text") else str(item)
    elif hasattr(r, "text"):
        text = r.text
    else:
        text = str(r)
    try:
        return json.loads(text)
    except (json.JSONDecodeError, TypeError):
        return text


async def test():
    results = {}
    temp_files = []  # track files to clean up

    # ── Setup ──────────────────────────────────────────────────────────
    print("=" * 60)
    print("Phase 7 Integration Test — 13 Analysis & Export Tools")
    print("=" * 60)

    print("\n[Setup] Launching PowerPoint ...")
    r = await call("launch_powerpoint")
    print(f"  launch_powerpoint -> {r}")

    print("[Setup] Creating new presentation ...")
    r = await call("new_presentation")
    print(f"  new_presentation -> {r}")

    # Add 3 slides with different layouts
    print("[Setup] Adding slide 1 (blank) ...")
    r = await call("add_slide", {"layout": "blank"})
    print(f"  add_slide (blank) -> {r}")

    print("[Setup] Adding slide 2 (title) ...")
    r = await call("add_slide", {"layout": "title"})
    print(f"  add_slide (title) -> {r}")

    print("[Setup] Adding slide 3 (title_content) ...")
    r = await call("add_slide", {"layout": "title_content"})
    print(f"  add_slide (title_content) -> {r}")

    # Add textboxes with text on each slide
    print("[Setup] Adding textbox on slide 1 ...")
    r = await call("add_textbox", {
        "slide_index": 1, "text": "Hello from slide one",
        "left": 1, "top": 1, "width": 4, "height": 1,
    })
    print(f"  add_textbox (slide 1) -> {r}")

    print("[Setup] Adding textbox on slide 2 ...")
    r = await call("add_textbox", {
        "slide_index": 2, "text": "Content on slide two",
        "left": 1, "top": 1, "width": 4, "height": 1,
    })
    print(f"  add_textbox (slide 2) -> {r}")

    print("[Setup] Adding textbox on slide 3 ...")
    r = await call("add_textbox", {
        "slide_index": 3, "text": "Text on slide three",
        "left": 1, "top": 1, "width": 4, "height": 1,
    })
    print(f"  add_textbox (slide 3) -> {r}")

    # Add a shape with fill color on slide 1
    print("[Setup] Adding shape with fill color on slide 1 ...")
    r = await call("add_shape", {
        "slide_index": 1, "shape_type": "rectangle",
        "left": 5, "top": 2, "width": 2, "height": 1.5,
        "fill_color": "#FF5500",
    })
    print(f"  add_shape (slide 1) -> {r}")

    # ── 1. get_presentation_stats ──────────────────────────────────────
    print("\n[1/13] get_presentation_stats")
    r = await call("get_presentation_stats")
    print(f"  -> {r}")
    if isinstance(r, dict) and "error" not in r:
        assert "slide_count" in r, "Missing slide_count"
        assert "total_shapes" in r, "Missing total_shapes (shapes_count)"
        assert "total_word_count" in r, "Missing total_word_count (word_count)"
        results["get_presentation_stats"] = "PASS"
    else:
        results["get_presentation_stats"] = f"FAIL: {r}"

    # ── 2. export_slide_image ──────────────────────────────────────────
    print("\n[2/13] export_slide_image")
    tmp_png = os.path.join(tempfile.gettempdir(), "phase7_slide1.png")
    temp_files.append(tmp_png)
    r = await call("export_slide_image", {
        "slide_index": 1, "output_path": tmp_png,
    })
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "exported":
        results["export_slide_image"] = "PASS"
    else:
        results["export_slide_image"] = f"FAIL: {r}"

    # ── 3. export_all_slides_images ────────────────────────────────────
    print("\n[3/13] export_all_slides_images")
    tmp_dir = os.path.join(tempfile.gettempdir(), "phase7_all_slides")
    temp_files.append(tmp_dir)
    r = await call("export_all_slides_images", {"output_dir": tmp_dir})
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "exported" and r.get("count", 0) >= 3:
        results["export_all_slides_images"] = "PASS"
    else:
        results["export_all_slides_images"] = f"FAIL: {r}"

    # ── 4. export_pdf ──────────────────────────────────────────────────
    print("\n[4/13] export_pdf")
    tmp_pdf = os.path.join(tempfile.gettempdir(), "phase7_export.pdf")
    temp_files.append(tmp_pdf)
    r = await call("export_pdf", {"output_path": tmp_pdf})
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "exported":
        results["export_pdf"] = "PASS"
    else:
        results["export_pdf"] = f"FAIL: {r}"

    # ── 5. get_fonts_used ──────────────────────────────────────────────
    print("\n[5/13] get_fonts_used")
    r = await call("get_fonts_used")
    print(f"  -> {r}")
    if isinstance(r, dict) and "error" not in r and isinstance(r.get("fonts"), list):
        results["get_fonts_used"] = "PASS"
    else:
        results["get_fonts_used"] = f"FAIL: {r}"

    # ── 6. get_linked_files ────────────────────────────────────────────
    print("\n[6/13] get_linked_files")
    r = await call("get_linked_files")
    print(f"  -> {r}")
    if isinstance(r, dict) and "error" not in r and isinstance(r.get("linked_files"), list):
        results["get_linked_files"] = "PASS"
    else:
        results["get_linked_files"] = f"FAIL: {r}"

    # ── 7. check_accessibility ─────────────────────────────────────────
    print("\n[7/13] check_accessibility")
    r = await call("check_accessibility")
    print(f"  -> {r}")
    if isinstance(r, dict) and "error" not in r and "issues" in r:
        results["check_accessibility"] = "PASS"
    else:
        results["check_accessibility"] = f"FAIL: {r}"

    # ── 8. get_slide_thumbnails_base64 ─────────────────────────────────
    print("\n[8/13] get_slide_thumbnails_base64")
    r = await call("get_slide_thumbnails_base64", {
        "slide_indices_json": "[1]",
    })
    print(f"  -> (keys: {list(r.keys()) if isinstance(r, dict) else 'N/A'})")
    if isinstance(r, dict) and "error" not in r:
        thumbs = r.get("thumbnails", [])
        assert isinstance(thumbs, list) and len(thumbs) > 0, "No thumbnails returned"
        assert "base64_image" in thumbs[0], "Missing base64 data"
        results["get_slide_thumbnails_base64"] = "PASS"
    else:
        results["get_slide_thumbnails_base64"] = f"FAIL: {r}"

    # ── 9. compare_slides ──────────────────────────────────────────────
    print("\n[9/13] compare_slides")
    r = await call("compare_slides", {"slide_a": 1, "slide_b": 2})
    print(f"  -> {r}")
    if isinstance(r, dict) and "error" not in r and "differences" in r:
        results["compare_slides"] = "PASS"
    else:
        results["compare_slides"] = f"FAIL: {r}"

    # ── 10. snapshot_to_json ───────────────────────────────────────────
    print("\n[10/13] snapshot_to_json")
    r = await call("snapshot_to_json")
    print(f"  -> (keys: {list(r.keys()) if isinstance(r, dict) else 'N/A'})")
    if isinstance(r, dict) and "error" not in r and "slides" in r:
        assert isinstance(r["slides"], list), "slides should be a list"
        results["snapshot_to_json"] = "PASS"
    else:
        results["snapshot_to_json"] = f"FAIL: {r}"

    # ── 11. get_color_usage ────────────────────────────────────────────
    print("\n[11/13] get_color_usage")
    r = await call("get_color_usage")
    print(f"  -> {r}")
    if isinstance(r, dict) and "error" not in r:
        assert "colors" in r, "Missing colors key"
        results["get_color_usage"] = "PASS"
    else:
        results["get_color_usage"] = f"FAIL: {r}"

    # ── 12. validate_presentation ──────────────────────────────────────
    print("\n[12/13] validate_presentation")
    r = await call("validate_presentation")
    print(f"  -> {r}")
    if isinstance(r, dict) and "error" not in r:
        assert "issues" in r, "Missing issues key"
        assert "score" in r, "Missing score key"
        results["validate_presentation"] = "PASS"
    else:
        results["validate_presentation"] = f"FAIL: {r}"

    # ── 13. get_text_by_slide ──────────────────────────────────────────
    print("\n[13/13] get_text_by_slide")
    r = await call("get_text_by_slide")
    print(f"  -> {r}")
    if isinstance(r, dict) and "error" not in r:
        slides_list = r.get("slides", [])
        assert isinstance(slides_list, list) and len(slides_list) > 0, "No slide text returned"
        results["get_text_by_slide"] = "PASS"
    else:
        results["get_text_by_slide"] = f"FAIL: {r}"

    # ── Summary & Cleanup ──────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    passed = sum(1 for v in results.values() if v == "PASS")
    total = len(results)
    for name, status in results.items():
        mark = "PASS" if status == "PASS" else "FAIL"
        print(f"  [{mark}] {name}")
    print(f"\n  {passed}/{total} passed")

    # Close without saving
    print("\n[Cleanup] Closing presentation without saving ...")
    r = await call("close_presentation", {"save": False})
    print(f"  close_presentation -> {r}")

    # Remove temp files
    import shutil
    for path in temp_files:
        try:
            if os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
            elif os.path.isfile(path):
                os.remove(path)
        except Exception:
            pass
    print("[Cleanup] Temp files removed.")

    print("\nDone.")


if __name__ == "__main__":
    asyncio.run(test())
