"""Integration test for Phase 3 PowerPoint MCP tools (18 tools)."""
import asyncio
import json
import importlib.util
import os
import traceback

_server_path = os.path.join(os.path.dirname(__file__), "..", "server.py")
spec = importlib.util.spec_from_file_location("server", _server_path)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


async def call(name, args=None):
    """Call an MCP tool and return parsed JSON or raw text."""
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

    # ===================================================================
    # Setup: launch PowerPoint, create new presentation, add blank slide
    # ===================================================================
    print("=== Phase 3: Text & Shapes (18 tools) ===\n")
    print("--- Setup ---")

    try:
        r = await call("launch_powerpoint")
        assert "status" in r, f"launch_powerpoint failed: {r}"
        print(f"  [OK] launch_powerpoint -> version={r.get('version')}")
    except Exception as e:
        print(f"  [FAIL] launch_powerpoint -> {e}")
        traceback.print_exc()
        return

    try:
        r = await call("new_presentation")
        assert r.get("status") == "created", f"new_presentation failed: {r}"
        print(f"  [OK] new_presentation -> {r.get('name')}")
    except Exception as e:
        print(f"  [FAIL] new_presentation -> {e}")
        traceback.print_exc()
        return

    try:
        r = await call("add_slide", {"layout": "blank"})
        assert "index" in r, f"add_slide failed: {r}"
        print(f"  [OK] add_slide (blank) -> index={r.get('index')}")
    except Exception as e:
        print(f"  [FAIL] add_slide -> {e}")
        traceback.print_exc()
        return

    print("\n--- Tests ---")

    # Track shape names for later reuse
    textbox_name = None
    rect_name = None
    line_name = None

    # -----------------------------------------------------------------------
    # 1. add_textbox
    # -----------------------------------------------------------------------
    try:
        r = await call("add_textbox", {
            "slide_index": 1,
            "text": "Hello World",
            "left": 1, "top": 1, "width": 4, "height": 1,
            "font_size": 24,
        })
        assert "name" in r, f"Missing 'name': {r}"
        textbox_name = r["name"]
        results["add_textbox"] = "PASS"
        print(f"  [PASS] add_textbox        -> name={textbox_name}")
    except Exception as e:
        results["add_textbox"] = "FAIL"
        print(f"  [FAIL] add_textbox        -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 2. get_shapes
    # -----------------------------------------------------------------------
    try:
        r = await call("get_shapes", {"slide_index": 1})
        assert isinstance(r, list), f"Expected list: {r}"
        assert len(r) >= 1, f"Expected at least 1 shape: {r}"
        results["get_shapes"] = "PASS"
        print(f"  [PASS] get_shapes         -> {len(r)} shape(s)")
    except Exception as e:
        results["get_shapes"] = "FAIL"
        print(f"  [FAIL] get_shapes         -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 3. get_shape_details
    # -----------------------------------------------------------------------
    try:
        assert textbox_name, "No textbox_name from add_textbox"
        r = await call("get_shape_details", {
            "slide_index": 1,
            "shape_name": textbox_name,
        })
        assert "name" in r, f"Missing 'name': {r}"
        results["get_shape_details"] = "PASS"
        print(f"  [PASS] get_shape_details  -> name={r['name']}")
    except Exception as e:
        results["get_shape_details"] = "FAIL"
        print(f"  [FAIL] get_shape_details  -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 4. modify_text
    # -----------------------------------------------------------------------
    try:
        assert textbox_name, "No textbox_name from add_textbox"
        r = await call("modify_text", {
            "slide_index": 1,
            "shape_name": textbox_name,
            "text": "Updated Text",
        })
        assert "error" not in r, f"Error: {r}"
        results["modify_text"] = "PASS"
        print(f"  [PASS] modify_text        -> updated")
    except Exception as e:
        results["modify_text"] = "FAIL"
        print(f"  [FAIL] modify_text        -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 5. add_shape (rectangle)
    # -----------------------------------------------------------------------
    try:
        r = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "rectangle",
            "left": 5, "top": 1, "width": 2, "height": 2,
            "fill_color": "#0000FF",
            "text": "Box",
        })
        assert "name" in r, f"Missing 'name': {r}"
        rect_name = r["name"]
        results["add_shape"] = "PASS"
        print(f"  [PASS] add_shape          -> name={rect_name}")
    except Exception as e:
        results["add_shape"] = "FAIL"
        print(f"  [FAIL] add_shape          -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 6. add_line
    # -----------------------------------------------------------------------
    try:
        r = await call("add_line", {
            "slide_index": 1,
            "begin_x": 1, "begin_y": 4, "end_x": 5, "end_y": 4,
            "color": "#FF0000",
            "weight": 2.0,
        })
        assert "name" in r, f"Missing 'name': {r}"
        line_name = r["name"]
        results["add_line"] = "PASS"
        print(f"  [PASS] add_line           -> name={line_name}")
    except Exception as e:
        results["add_line"] = "FAIL"
        print(f"  [FAIL] add_line           -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 7. modify_shape (change rectangle width to 3)
    # -----------------------------------------------------------------------
    try:
        assert rect_name, "No rect_name from add_shape"
        r = await call("modify_shape", {
            "slide_index": 1,
            "shape_name": rect_name,
            "width": 3,
        })
        assert "error" not in r, f"Error: {r}"
        results["modify_shape"] = "PASS"
        print(f"  [PASS] modify_shape       -> width set to 3")
    except Exception as e:
        results["modify_shape"] = "FAIL"
        print(f"  [FAIL] modify_shape       -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 8. delete_shape (delete the line)
    # -----------------------------------------------------------------------
    try:
        assert line_name, "No line_name from add_line"
        r = await call("delete_shape", {
            "slide_index": 1,
            "shape_name": line_name,
        })
        assert r.get("status") == "deleted", f"Expected status='deleted': {r}"
        results["delete_shape"] = "PASS"
        print(f"  [PASS] delete_shape       -> deleted {line_name}")
    except Exception as e:
        results["delete_shape"] = "FAIL"
        print(f"  [FAIL] delete_shape       -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 9. group_shapes (add 2 shapes, then group)
    # -----------------------------------------------------------------------
    group_name = None
    grp_shape1 = None
    grp_shape2 = None
    try:
        r1 = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "rectangle",
            "left": 0.5, "top": 5, "width": 1, "height": 1,
            "fill_color": "#00FF00",
        })
        r2 = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "oval",
            "left": 2, "top": 5, "width": 1, "height": 1,
            "fill_color": "#FFFF00",
        })
        grp_shape1 = r1["name"]
        grp_shape2 = r2["name"]
        names_json = json.dumps([grp_shape1, grp_shape2])
        r = await call("group_shapes", {
            "slide_index": 1,
            "shape_names_json": names_json,
            "group_name": "TestGroup",
        })
        assert "name" in r, f"Missing 'name': {r}"
        group_name = r["name"]
        results["group_shapes"] = "PASS"
        print(f"  [PASS] group_shapes       -> name={group_name}")
    except Exception as e:
        results["group_shapes"] = "FAIL"
        print(f"  [FAIL] group_shapes       -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 10. ungroup_shapes
    # -----------------------------------------------------------------------
    try:
        assert group_name, "No group_name from group_shapes"
        r = await call("ungroup_shapes", {
            "slide_index": 1,
            "group_name": group_name,
        })
        assert r.get("status") == "ungrouped", f"Expected status='ungrouped': {r}"
        results["ungroup_shapes"] = "PASS"
        print(f"  [PASS] ungroup_shapes     -> ungrouped {group_name}")
    except Exception as e:
        results["ungroup_shapes"] = "FAIL"
        print(f"  [FAIL] ungroup_shapes     -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 11. add_hyperlink
    # -----------------------------------------------------------------------
    try:
        assert textbox_name, "No textbox_name from add_textbox"
        r = await call("add_hyperlink", {
            "slide_index": 1,
            "shape_name": textbox_name,
            "url": "https://example.com",
        })
        assert "status" in r, f"Missing 'status': {r}"
        results["add_hyperlink"] = "PASS"
        print(f"  [PASS] add_hyperlink      -> {r.get('status')}")
    except Exception as e:
        results["add_hyperlink"] = "FAIL"
        print(f"  [FAIL] add_hyperlink      -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 12. add_connector (add 2 shapes, connect them)
    # -----------------------------------------------------------------------
    try:
        r1 = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "oval",
            "left": 1, "top": 6, "width": 1, "height": 1,
        })
        r2 = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "oval",
            "left": 5, "top": 6, "width": 1, "height": 1,
        })
        conn_shape1 = r1["name"]
        conn_shape2 = r2["name"]
        r = await call("add_connector", {
            "slide_index": 1,
            "shape1_name": conn_shape1,
            "shape2_name": conn_shape2,
            "connector_type": "straight",
        })
        assert "error" not in r, f"Error: {r}"
        results["add_connector"] = "PASS"
        print(f"  [PASS] add_connector      -> connected {conn_shape1} to {conn_shape2}")
    except Exception as e:
        results["add_connector"] = "FAIL"
        print(f"  [FAIL] add_connector      -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 13. align_shapes (create 2 shapes, align center)
    # -----------------------------------------------------------------------
    try:
        r1 = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "rectangle",
            "left": 1, "top": 0.5, "width": 1, "height": 0.5,
        })
        r2 = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "rectangle",
            "left": 3, "top": 0.5, "width": 1, "height": 0.5,
        })
        align_names = json.dumps([r1["name"], r2["name"]])
        r = await call("align_shapes", {
            "slide_index": 1,
            "shape_names_json": align_names,
            "alignment": "center",
        })
        assert r.get("status") == "aligned", f"Expected status='aligned': {r}"
        results["align_shapes"] = "PASS"
        print(f"  [PASS] align_shapes       -> aligned center")
    except Exception as e:
        results["align_shapes"] = "FAIL"
        print(f"  [FAIL] align_shapes       -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 14. distribute_shapes (3 shapes, distribute horizontally)
    # -----------------------------------------------------------------------
    try:
        d1 = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "rectangle",
            "left": 0.5, "top": 3, "width": 0.8, "height": 0.5,
        })
        d2 = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "rectangle",
            "left": 3, "top": 3, "width": 0.8, "height": 0.5,
        })
        d3 = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "rectangle",
            "left": 6, "top": 3, "width": 0.8, "height": 0.5,
        })
        dist_names = json.dumps([d1["name"], d2["name"], d3["name"]])
        r = await call("distribute_shapes", {
            "slide_index": 1,
            "shape_names_json": dist_names,
            "direction": "horizontal",
        })
        assert r.get("status") == "distributed", f"Expected status='distributed': {r}"
        results["distribute_shapes"] = "PASS"
        print(f"  [PASS] distribute_shapes  -> distributed horizontally")
    except Exception as e:
        results["distribute_shapes"] = "FAIL"
        print(f"  [FAIL] distribute_shapes  -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 15. duplicate_shape (duplicate the rectangle)
    # -----------------------------------------------------------------------
    try:
        assert rect_name, "No rect_name from add_shape"
        r = await call("duplicate_shape", {
            "slide_index": 1,
            "shape_name": rect_name,
        })
        assert "name" in r, f"Missing 'name': {r}"
        results["duplicate_shape"] = "PASS"
        print(f"  [PASS] duplicate_shape    -> name={r['name']}")
    except Exception as e:
        results["duplicate_shape"] = "FAIL"
        print(f"  [FAIL] duplicate_shape    -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 16. set_shape_z_order (send a shape to back)
    # -----------------------------------------------------------------------
    try:
        assert rect_name, "No rect_name from add_shape"
        r = await call("set_shape_z_order", {
            "slide_index": 1,
            "shape_name": rect_name,
            "action": "back",
        })
        assert r.get("status") == "updated", f"Expected status='updated': {r}"
        results["set_shape_z_order"] = "PASS"
        print(f"  [PASS] set_shape_z_order  -> sent to back")
    except Exception as e:
        results["set_shape_z_order"] = "FAIL"
        print(f"  [FAIL] set_shape_z_order  -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 17. bulk_add_shapes (add 2 shapes via JSON array)
    # -----------------------------------------------------------------------
    try:
        shapes_spec = json.dumps([
            {"type": "rectangle", "left": 7, "top": 1, "width": 1.5, "height": 1,
             "fill_color": "#FF00FF", "text": "Bulk1"},
            {"type": "oval", "left": 7, "top": 3, "width": 1.5, "height": 1,
             "fill_color": "#00FFFF", "text": "Bulk2"},
        ])
        r = await call("bulk_add_shapes", {
            "slide_index": 1,
            "shapes_json": shapes_spec,
        })
        assert isinstance(r, list), f"Expected list: {r}"
        assert len(r) == 2, f"Expected 2 results: {r}"
        results["bulk_add_shapes"] = "PASS"
        print(f"  [PASS] bulk_add_shapes    -> {len(r)} shapes added")
    except Exception as e:
        results["bulk_add_shapes"] = "FAIL"
        print(f"  [FAIL] bulk_add_shapes    -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 18. set_placeholder_text (add title slide first, then set placeholder)
    # -----------------------------------------------------------------------
    try:
        slide_r = await call("add_slide", {"layout": "title"})
        title_slide_idx = slide_r["index"]
        r = await call("set_placeholder_text", {
            "slide_index": title_slide_idx,
            "placeholder_index": 1,
            "text": "Phase 3 Test Title",
        })
        assert r.get("status") == "updated", f"Expected status='updated': {r}"
        results["set_placeholder_text"] = "PASS"
        print(f"  [PASS] set_placeholder_text -> slide {title_slide_idx}, placeholder 1")
    except Exception as e:
        results["set_placeholder_text"] = "FAIL"
        print(f"  [FAIL] set_placeholder_text -> {e}")
        traceback.print_exc()

    # ===================================================================
    # Summary
    # ===================================================================
    print("\n=== Summary ===")
    passed = sum(1 for v in results.values() if v == "PASS")
    failed = sum(1 for v in results.values() if v == "FAIL")
    total = len(results)
    print(f"  Total: {total}  |  Passed: {passed}  |  Failed: {failed}")
    for tool_name, status in results.items():
        mark = "PASS" if status == "PASS" else "FAIL"
        print(f"    [{mark}] {tool_name}")

    # ===================================================================
    # Cleanup: close without saving
    # ===================================================================
    print("\n--- Cleanup ---")
    for _ in range(5):
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            break

    print("Done.\n")


if __name__ == "__main__":
    asyncio.run(test())
