"""Integration test for Phase 4 PowerPoint MCP tools (14 tools)."""
import asyncio
import json
import importlib.util
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(__file__))
from fixtures import create_test_bmp, create_test_bmp_2, create_test_wav, create_test_avi, create_test_txt, cleanup_files

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
    table_shape_name = None
    chart_shape_name = None

    # ── Setup ──────────────────────────────────────────────────────────
    print("=" * 60)
    print("Phase 4 Integration Test — 14 Rich Media Tools")
    print("=" * 60)

    print("\n[SETUP] Launching PowerPoint...")
    r = await call("launch_powerpoint")
    print(f"  launch_powerpoint: {r}")

    print("[SETUP] Creating new presentation...")
    r = await call("new_presentation")
    print(f"  new_presentation: {r}")

    print("[SETUP] Adding slide 1 (blank)...")
    r = await call("add_slide", {"layout": "blank"})
    print(f"  add_slide 1: {r}")

    print("[SETUP] Adding slide 2 (blank)...")
    r = await call("add_slide", {"layout": "blank"})
    print(f"  add_slide 2: {r}")

    # ── Fixture files ───────────────────────────────────────────────────
    tmp = tempfile.gettempdir()
    test_bmp = create_test_bmp(os.path.join(tmp, "phase4_test.bmp"))
    test_bmp2 = create_test_bmp_2(os.path.join(tmp, "phase4_test2.bmp"))
    test_wav = create_test_wav(os.path.join(tmp, "phase4_test.wav"))
    test_avi = create_test_avi(os.path.join(tmp, "phase4_test.avi"))
    test_txt = create_test_txt(os.path.join(tmp, "phase4_test.txt"))
    image_shape_name = None

    # ── 1. insert_image ─────────────────────────────────────────────
    print("\n[1/14] insert_image ...")
    try:
        r = await call("insert_image", {
            "slide_index": 1,
            "image_path": test_bmp,
            "left": 1, "top": 1,
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and "error" not in r, f"insert_image failed: {r}"
        image_shape_name = r.get("name")
        results["insert_image"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["insert_image"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 2. insert_image_from_url (use local file:// URL) ────────────
    print("\n[2/14] insert_image_from_url ...")
    try:
        # Use a file:// URL pointing to the local test BMP
        file_url = "file:///" + test_bmp2.replace("\\", "/")
        r = await call("insert_image_from_url", {
            "slide_index": 1,
            "url": file_url,
            "left": 4, "top": 1,
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and "error" not in r, f"insert_image_from_url failed: {r}"
        results["insert_image_from_url"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["insert_image_from_url"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 3. add_table ──────────────────────────────────────────────────
    print("\n[3/14] add_table ...")
    try:
        r = await call("add_table", {
            "slide_index": 1,
            "rows": 3,
            "cols": 3,
            "left": 1,
            "top": 1,
            "width": 6,
            "height": 3,
            "data_json": json.dumps([["A", "B", "C"], ["1", "2", "3"], ["X", "Y", "Z"]]),
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and ("shape_name" in r or "name" in r), f"Expected shape info, got {r}"
        table_shape_name = r.get("shape_name") or r.get("name")
        results["add_table"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["add_table"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 4. modify_table_cell ──────────────────────────────────────────
    print("\n[4/14] modify_table_cell ...")
    try:
        assert table_shape_name, "No table shape from add_table"
        r = await call("modify_table_cell", {
            "slide_index": 1,
            "shape_name": table_shape_name,
            "row": 1,
            "col": 1,
            "text": "Modified",
            "font_size": 14,
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and r.get("status") == "updated", f"Expected status=updated, got {r}"
        results["modify_table_cell"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["modify_table_cell"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 5. bulk_fill_table ────────────────────────────────────────────
    print("\n[5/14] bulk_fill_table ...")
    try:
        assert table_shape_name, "No table shape from add_table"
        r = await call("bulk_fill_table", {
            "slide_index": 1,
            "shape_name": table_shape_name,
            "data_json": json.dumps([["H1", "H2", "H3"], ["D1", "D2", "D3"], ["E1", "E2", "E3"]]),
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and r.get("status") == "filled", f"Expected status=filled, got {r}"
        results["bulk_fill_table"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["bulk_fill_table"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 6. format_table ───────────────────────────────────────────────
    print("\n[6/14] format_table ...")
    try:
        assert table_shape_name, "No table shape from add_table"
        r = await call("format_table", {
            "slide_index": 1,
            "shape_name": table_shape_name,
            "has_header": True,
            "header_color": "#003366",
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and r.get("status") == "formatted", f"Expected status=formatted, got {r}"
        results["format_table"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["format_table"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 7. add_chart ──────────────────────────────────────────────────
    print("\n[7/14] add_chart ...")
    try:
        chart_data = {
            "categories": ["Q1", "Q2", "Q3", "Q4"],
            "series": [
                {"name": "Revenue", "values": [100, 150, 130, 180]},
                {"name": "Cost", "values": [80, 90, 85, 110]},
            ],
        }
        r = await call("add_chart", {
            "slide_index": 2,
            "chart_type": "column",
            "data_json": json.dumps(chart_data),
            "left": 1,
            "top": 1,
            "width": 6,
            "height": 4,
            "title": "Test Chart",
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and "error" not in r, f"Expected chart info, got {r}"
        chart_shape_name = r.get("shape_name") or r.get("name")
        results["add_chart"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["add_chart"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 8. modify_chart ───────────────────────────────────────────────
    print("\n[8/14] modify_chart ...")
    try:
        assert chart_shape_name, "No chart shape from add_chart"
        r = await call("modify_chart", {
            "slide_index": 2,
            "shape_name": chart_shape_name,
            "title": "Updated Chart",
            "has_legend": True,
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and r.get("status") == "updated", f"Expected status=updated, got {r}"
        results["modify_chart"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["modify_chart"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 9. update_chart_data ──────────────────────────────────────────
    print("\n[9/14] update_chart_data ...")
    try:
        assert chart_shape_name, "No chart shape from add_chart"
        new_chart_data = {
            "categories": ["Jan", "Feb", "Mar"],
            "series": [
                {"name": "Sales", "values": [200, 250, 300]},
            ],
        }
        r = await call("update_chart_data", {
            "slide_index": 2,
            "shape_name": chart_shape_name,
            "data_json": json.dumps(new_chart_data),
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and r.get("status") == "updated", f"Expected status=updated, got {r}"
        results["update_chart_data"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["update_chart_data"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 10. insert_video ────────────────────────────────────────────────
    print("\n[10/14] insert_video ...")
    try:
        r = await call("insert_video", {
            "slide_index": 2,
            "video_path": test_avi,
            "left": 1, "top": 5, "width": 3, "height": 2,
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and "error" not in r, f"insert_video failed: {r}"
        results["insert_video"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["insert_video"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 11. insert_audio ────────────────────────────────────────────────
    print("\n[11/14] insert_audio ...")
    try:
        r = await call("insert_audio", {
            "slide_index": 2,
            "audio_path": test_wav,
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and "error" not in r, f"insert_audio failed: {r}"
        results["insert_audio"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["insert_audio"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 12. insert_ole_object ───────────────────────────────────────────
    print("\n[12/14] insert_ole_object ...")
    try:
        r = await call("insert_ole_object", {
            "slide_index": 2,
            "file_path": test_txt,
            "left": 5, "top": 5, "width": 2, "height": 1,
            "as_icon": True,
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and "error" not in r, f"insert_ole_object failed: {r}"
        results["insert_ole_object"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["insert_ole_object"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 13. crop_image ──────────────────────────────────────────────────
    print("\n[13/14] crop_image ...")
    try:
        assert image_shape_name, "No image shape from insert_image"
        r = await call("crop_image", {
            "slide_index": 1,
            "shape_name": image_shape_name,
            "crop_left": 2, "crop_top": 2,
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and "error" not in r, f"crop_image failed: {r}"
        results["crop_image"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["crop_image"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── 14. replace_image ───────────────────────────────────────────────
    print("\n[14/14] replace_image ...")
    try:
        assert image_shape_name, "No image shape from insert_image"
        r = await call("replace_image", {
            "slide_index": 1,
            "shape_name": image_shape_name,
            "new_image_path": test_bmp2,
        })
        print(f"  Result: {r}")
        assert isinstance(r, dict) and "error" not in r, f"replace_image failed: {r}"
        results["replace_image"] = "PASS"
        print("  -> PASS")
    except Exception as e:
        results["replace_image"] = "FAIL"
        print(f"  -> FAIL: {e}")

    # ── Summary ───────────────────────────────────────────────────────
    passed = sum(1 for v in results.values() if v == "PASS")
    failed = sum(1 for v in results.values() if v == "FAIL")
    skipped = sum(1 for v in results.values() if v == "SKIP")

    print("\n" + "=" * 60)
    print(f"Phase 4 Results: {passed} PASS / {failed} FAIL / {skipped} SKIP  (of {len(results)} tools)")
    print("=" * 60)

    for tool_name, status in results.items():
        marker = {"PASS": "[PASS]", "FAIL": "[FAIL]", "SKIP": "[SKIP]"}[status]
        print(f"  {marker} {tool_name}")

    # ── Cleanup ───────────────────────────────────────────────────────
    print("\n[CLEANUP] Closing presentation without saving...")
    try:
        r = await call("close_presentation", {"save": False})
        print(f"  close_presentation: {r}")
    except Exception:
        print("  close_presentation: skipped (tool may not exist)")

    # Clean up fixture files
    cleanup_files(test_bmp, test_bmp2, test_wav, test_avi, test_txt)
    print("[CLEANUP] Fixture files removed.")

    if failed > 0:
        raise SystemExit(f"\n{failed} test(s) FAILED.")

    print("\nAll executed tests passed.")


if __name__ == "__main__":
    asyncio.run(test())
