"""Negative tests for PowerPoint MCP server — validates error codes and messages."""
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


def assert_error(r, expected_code, label):
    """Assert that the response is an error with the given code."""
    assert "error" in r, f"{label}: expected error response, got {r}"
    actual_code = r.get("code", "")
    assert actual_code == expected_code, (
        f"{label}: expected code='{expected_code}', got code='{actual_code}', "
        f"error='{r.get('error', '')}'"
    )


async def test():
    results = {}

    # -----------------------------------------------------------------------
    # Setup: ensure a presentation is open
    # -----------------------------------------------------------------------
    try:
        await call("launch_powerpoint")
        await call("new_presentation")
    except Exception:
        print("  [SKIP] Could not set up PowerPoint for negative tests")
        return results

    # -----------------------------------------------------------------------
    # 1. Invalid slide_index (out of bounds)
    # -----------------------------------------------------------------------
    for idx, label in [(0, "zero"), (-1, "negative"), (9999, "too_large")]:
        test_name = f"invalid_slide_index_{label}"
        try:
            r = await call("get_slide_info", {"slide_index": idx})
            assert_error(r, "OUT_OF_BOUNDS", test_name)
            results[test_name] = "PASS"
            print(f"  [PASS] {test_name}")
        except Exception as e:
            results[test_name] = "FAIL"
            print(f"  [FAIL] {test_name} -> {e}")

    # -----------------------------------------------------------------------
    # 2. Non-existent shape_name
    # -----------------------------------------------------------------------
    try:
        r = await call("get_shape_details", {
            "slide_index": 1,
            "shape_name": "NonExistentShapeXYZ_12345",
        })
        assert_error(r, "NOT_FOUND", "non_existent_shape")
        results["non_existent_shape"] = "PASS"
        print(f"  [PASS] non_existent_shape")
    except Exception as e:
        results["non_existent_shape"] = "FAIL"
        print(f"  [FAIL] non_existent_shape -> {e}")

    # -----------------------------------------------------------------------
    # 3. Zero/negative dimensions
    # -----------------------------------------------------------------------
    for dim_name, dim_val in [("zero_width", 0), ("neg_width", -5)]:
        test_name = f"invalid_dimension_{dim_name}"
        try:
            r = await call("add_textbox", {
                "slide_index": 1,
                "text": "Test",
                "left": 1, "top": 1,
                "width": dim_val if "width" in dim_name else 2,
                "height": 1,
            })
            assert_error(r, "VALIDATION_ERROR", test_name)
            results[test_name] = "PASS"
            print(f"  [PASS] {test_name}")
        except Exception as e:
            results[test_name] = "FAIL"
            print(f"  [FAIL] {test_name} -> {e}")

    for dim_name, dim_val in [("zero_height", 0), ("neg_height", -3)]:
        test_name = f"invalid_dimension_{dim_name}"
        try:
            r = await call("add_textbox", {
                "slide_index": 1,
                "text": "Test",
                "left": 1, "top": 1,
                "width": 2,
                "height": dim_val,
            })
            assert_error(r, "VALIDATION_ERROR", test_name)
            results[test_name] = "PASS"
            print(f"  [PASS] {test_name}")
        except Exception as e:
            results[test_name] = "FAIL"
            print(f"  [FAIL] {test_name} -> {e}")

    # -----------------------------------------------------------------------
    # 4. Malformed JSON (dict instead of list)
    # -----------------------------------------------------------------------
    try:
        r = await call("bulk_add_slides", {"slides_json": '{"layout": "blank"}'})
        assert_error(r, "VALIDATION_ERROR", "json_dict_as_list")
        results["json_dict_as_list"] = "PASS"
        print(f"  [PASS] json_dict_as_list")
    except Exception as e:
        results["json_dict_as_list"] = "FAIL"
        print(f"  [FAIL] json_dict_as_list -> {e}")

    # -----------------------------------------------------------------------
    # 5. Malformed JSON (syntax error)
    # -----------------------------------------------------------------------
    try:
        r = await call("bulk_add_slides", {"slides_json": '{broken json'})
        assert "error" in r, "Expected error for broken JSON"
        results["json_syntax_error"] = "PASS"
        print(f"  [PASS] json_syntax_error")
    except Exception as e:
        results["json_syntax_error"] = "FAIL"
        print(f"  [FAIL] json_syntax_error -> {e}")

    # -----------------------------------------------------------------------
    # 6. Non-existent file paths
    # -----------------------------------------------------------------------
    try:
        r = await call("insert_image", {
            "slide_index": 1,
            "image_path": "C:/nonexistent/fake_image_12345.png",
            "left": 1, "top": 1,
        })
        assert_error(r, "NOT_FOUND", "nonexistent_file")
        results["nonexistent_file"] = "PASS"
        print(f"  [PASS] nonexistent_file")
    except Exception as e:
        results["nonexistent_file"] = "FAIL"
        print(f"  [FAIL] nonexistent_file -> {e}")

    # -----------------------------------------------------------------------
    # 7. Invalid layout name
    # -----------------------------------------------------------------------
    try:
        r = await call("add_slide", {"layout": "super_mega_layout_xyz"})
        assert "error" in r, "Expected error for invalid layout"
        results["invalid_layout"] = "PASS"
        print(f"  [PASS] invalid_layout")
    except Exception as e:
        results["invalid_layout"] = "FAIL"
        print(f"  [FAIL] invalid_layout -> {e}")

    # -----------------------------------------------------------------------
    # 8. Invalid color format
    # -----------------------------------------------------------------------
    for color_str, label in [
        ("#GGG", "bad_hex"),
        ("#12345", "short_hex"),
        ("300,0,0", "rgb_out_of_range"),
        ("notacolor", "no_format"),
    ]:
        test_name = f"invalid_color_{label}"
        try:
            r = await call("add_textbox", {
                "slide_index": 1,
                "text": "Test",
                "left": 1, "top": 1, "width": 2, "height": 1,
                "font_color": color_str,
            })
            assert_error(r, "VALIDATION_ERROR", test_name)
            results[test_name] = "PASS"
            print(f"  [PASS] {test_name}")
        except Exception as e:
            results[test_name] = "FAIL"
            print(f"  [FAIL] {test_name} -> {e}")

    # -----------------------------------------------------------------------
    # 9. Table cell out of bounds
    # -----------------------------------------------------------------------
    try:
        # First add a table
        r = await call("add_table", {
            "slide_index": 1,
            "rows": 2, "cols": 2,
            "left": 1, "top": 1, "width": 4, "height": 2,
        })
        if "error" not in r:
            table_name = r.get("shape_name", "")
            r2 = await call("modify_table_cell", {
                "slide_index": 1,
                "shape_name": table_name,
                "row": 99, "col": 1, "text": "Bad",
            })
            assert_error(r2, "OUT_OF_BOUNDS", "table_row_oob")
            results["table_cell_oob"] = "PASS"
            print(f"  [PASS] table_cell_oob")
        else:
            results["table_cell_oob"] = "SKIP"
            print(f"  [SKIP] table_cell_oob (table creation failed)")
    except Exception as e:
        results["table_cell_oob"] = "FAIL"
        print(f"  [FAIL] table_cell_oob -> {e}")

    # -----------------------------------------------------------------------
    # 10. Invalid URL
    # -----------------------------------------------------------------------
    try:
        # Need a shape first
        r = await call("add_textbox", {
            "slide_index": 1,
            "text": "Link",
            "left": 1, "top": 5, "width": 2, "height": 1,
        })
        if "error" not in r:
            shape_name = r.get("name", "")
            r2 = await call("add_hyperlink", {
                "slide_index": 1,
                "shape_name": shape_name,
                "url": "ftp://invalid.com",
            })
            assert_error(r2, "VALIDATION_ERROR", "invalid_url")
            results["invalid_url"] = "PASS"
            print(f"  [PASS] invalid_url")
        else:
            results["invalid_url"] = "SKIP"
    except Exception as e:
        results["invalid_url"] = "FAIL"
        print(f"  [FAIL] invalid_url -> {e}")

    # -----------------------------------------------------------------------
    # 11. Invalid shape_type
    # -----------------------------------------------------------------------
    try:
        r = await call("add_shape", {
            "slide_index": 1,
            "shape_type": "mega_blob_xyz",
            "left": 1, "top": 1, "width": 2, "height": 2,
        })
        assert "error" in r, "Expected error for invalid shape_type"
        results["invalid_shape_type"] = "PASS"
        print(f"  [PASS] invalid_shape_type")
    except Exception as e:
        results["invalid_shape_type"] = "FAIL"
        print(f"  [FAIL] invalid_shape_type -> {e}")

    # -----------------------------------------------------------------------
    # 12. Invalid placeholder index
    # -----------------------------------------------------------------------
    try:
        r = await call("set_placeholder_text", {
            "slide_index": 1,
            "placeholder_index": 999,
            "text": "Test",
        })
        assert_error(r, "OUT_OF_BOUNDS", "invalid_placeholder")
        results["invalid_placeholder"] = "PASS"
        print(f"  [PASS] invalid_placeholder")
    except Exception as e:
        results["invalid_placeholder"] = "FAIL"
        print(f"  [FAIL] invalid_placeholder -> {e}")

    # -----------------------------------------------------------------------
    # 13. Reorder with out-of-bounds slide index
    # -----------------------------------------------------------------------
    try:
        r = await call("reorder_slides", {"order_json": "[1, 999]"})
        assert_error(r, "OUT_OF_BOUNDS", "reorder_oob")
        results["reorder_oob"] = "PASS"
        print(f"  [PASS] reorder_oob")
    except Exception as e:
        results["reorder_oob"] = "FAIL"
        print(f"  [FAIL] reorder_oob -> {e}")

    # -----------------------------------------------------------------------
    # 14. Table rows/cols < 1
    # -----------------------------------------------------------------------
    try:
        r = await call("add_table", {
            "slide_index": 1,
            "rows": 0, "cols": 2,
            "left": 1, "top": 1, "width": 4, "height": 2,
        })
        assert_error(r, "VALIDATION_ERROR", "table_zero_rows")
        results["table_zero_rows"] = "PASS"
        print(f"  [PASS] table_zero_rows")
    except Exception as e:
        results["table_zero_rows"] = "FAIL"
        print(f"  [FAIL] table_zero_rows -> {e}")

    # -----------------------------------------------------------------------
    # Summary
    # -----------------------------------------------------------------------
    print(f"\n{'=' * 60}")
    passed = sum(1 for v in results.values() if v == "PASS")
    failed = sum(1 for v in results.values() if v == "FAIL")
    total = len(results)
    print(f"NEGATIVE TESTS: {passed}/{total} passed, {failed} failed")
    if failed:
        print("FAILED:")
        for k, v in results.items():
            if v == "FAIL":
                print(f"  - {k}")
    print(f"{'=' * 60}")

    return results


if __name__ == "__main__":
    asyncio.run(test())
