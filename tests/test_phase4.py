"""Integration test for Phase 4 PowerPoint MCP tools (14 tools)."""
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

    # ── 1. insert_image — SKIP ────────────────────────────────────────
    print("\n[1/14] insert_image — SKIP (no test image file available)")
    results["insert_image"] = "SKIP"

    # ── 2. insert_image_from_url — SKIP ───────────────────────────────
    print("[2/14] insert_image_from_url — SKIP (requires network)")
    results["insert_image_from_url"] = "SKIP"

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

    # ── 10. insert_video — SKIP ───────────────────────────────────────
    print("\n[10/14] insert_video — SKIP (no test video file available)")
    results["insert_video"] = "SKIP"

    # ── 11. insert_audio — SKIP ───────────────────────────────────────
    print("[11/14] insert_audio — SKIP (no test audio file available)")
    results["insert_audio"] = "SKIP"

    # ── 12. insert_ole_object — SKIP ──────────────────────────────────
    print("[12/14] insert_ole_object — SKIP (no test file available)")
    results["insert_ole_object"] = "SKIP"

    # ── 13. crop_image — SKIP ─────────────────────────────────────────
    print("[13/14] crop_image — SKIP (requires image on slide)")
    results["crop_image"] = "SKIP"

    # ── 14. replace_image — SKIP ──────────────────────────────────────
    print("[14/14] replace_image — SKIP (requires image on slide)")
    results["replace_image"] = "SKIP"

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

    if failed > 0:
        raise SystemExit(f"\n{failed} test(s) FAILED.")

    print("\nAll executed tests passed.")


if __name__ == "__main__":
    asyncio.run(test())
