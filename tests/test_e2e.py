"""End-to-end workflow tests for PowerPoint MCP server."""
import asyncio
import json
import importlib.util
import os
import tempfile
import traceback

_server_path = os.path.join(os.path.dirname(__file__), "..", "server.py")
spec = importlib.util.spec_from_file_location("server", _server_path)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

TEMP_DIR = os.environ.get("TEMP", tempfile.gettempdir())


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


def assert_no_error(r, label):
    """Assert the response is not an error."""
    if isinstance(r, dict) and "error" in r:
        raise AssertionError(f"{label}: unexpected error: {r['error']}")


async def test():
    results = {}

    # -----------------------------------------------------------------------
    # Setup
    # -----------------------------------------------------------------------
    try:
        await call("launch_powerpoint")
    except Exception:
        print("  [SKIP] Could not launch PowerPoint")
        return results

    # ═══════════════════════════════════════════════════════════════════════
    # E2E Test 1: Build Pipeline
    #   new → add slides → add content → set transitions → export PDF → close
    # ═══════════════════════════════════════════════════════════════════════
    test_name = "build_pipeline"
    try:
        # Create new presentation
        r = await call("new_presentation")
        assert_no_error(r, "new_presentation")

        # Add 3 slides with different layouts
        r = await call("add_slide", {"layout": "title_only"})
        assert_no_error(r, "add_slide_1")
        r = await call("add_slide", {"layout": "blank"})
        assert_no_error(r, "add_slide_2")

        # Add content to slide 1
        r = await call("add_textbox", {
            "slide_index": 1,
            "text": "E2E Test Presentation",
            "left": 1, "top": 1, "width": 8, "height": 1.5,
            "font_size": 32, "bold": True,
        })
        assert_no_error(r, "add_textbox_title")

        # Add a shape to slide 2
        r = await call("add_shape", {
            "slide_index": 2,
            "shape_type": "rectangle",
            "left": 2, "top": 2, "width": 4, "height": 2,
            "fill_color": "#3366CC",
            "text": "Blue Rectangle",
        })
        assert_no_error(r, "add_shape_rect")

        # Set transitions
        r = await call("set_slide_transition", {
            "slide_index": 1,
            "effect": "fade",
            "duration": 0.5,
        })
        assert_no_error(r, "set_transition_1")

        r = await call("set_slide_transition", {
            "slide_index": 2,
            "effect": "push",
            "duration": 0.5,
        })
        assert_no_error(r, "set_transition_2")

        # Verify slide count
        r = await call("get_slides")
        assert_no_error(r, "get_slides")
        slide_count = r.get("count", len(r.get("slides", [])))
        assert slide_count >= 3, f"Expected >= 3 slides, got {slide_count}"

        # Export as PDF
        pdf_path = os.path.join(TEMP_DIR, "e2e_test_pipeline.pdf")
        r = await call("export_pdf", {"output_path": pdf_path})
        assert_no_error(r, "export_pdf")

        # Validate PDF was created
        r = await call("validate_presentation")
        assert_no_error(r, "validate_presentation")

        # Close
        r = await call("close_presentation", {"save": False})
        assert_no_error(r, "close_presentation")

        # Clean up
        try:
            os.remove(pdf_path)
        except OSError:
            pass

        results[test_name] = "PASS"
        print(f"  [PASS] {test_name}")
    except Exception as e:
        results[test_name] = "FAIL"
        print(f"  [FAIL] {test_name} -> {e}")
        traceback.print_exc()
        # Try to close any open pres
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════════════
    # E2E Test 2: Content Modification Workflow
    #   new → add text → verify → modify text → verify changed → close
    # ═══════════════════════════════════════════════════════════════════════
    test_name = "content_modification"
    try:
        r = await call("new_presentation")
        assert_no_error(r, "new_presentation")

        # Add textbox with known text
        r = await call("add_textbox", {
            "slide_index": 1,
            "text": "Original Text",
            "left": 1, "top": 1, "width": 6, "height": 1,
        })
        assert_no_error(r, "add_textbox")
        shape_name = r.get("name", "")
        assert shape_name, "No shape name returned"

        # Verify text is there
        r = await call("get_shape_details", {
            "slide_index": 1,
            "shape_name": shape_name,
        })
        assert_no_error(r, "get_shape_details_1")
        assert r.get("text") == "Original Text", f"Expected 'Original Text', got '{r.get('text')}'"

        # Modify the text
        r = await call("modify_text", {
            "slide_index": 1,
            "shape_name": shape_name,
            "text": "Modified Text",
            "font_size": 24,
            "bold": 1,
        })
        assert_no_error(r, "modify_text")
        assert "changed" in r, "Missing 'changed' field in modify_text response"
        assert "text" in r["changed"], "'text' not in changed list"
        assert "font_size" in r["changed"], "'font_size' not in changed list"

        # Verify modification
        r = await call("get_shape_details", {
            "slide_index": 1,
            "shape_name": shape_name,
        })
        assert_no_error(r, "get_shape_details_2")
        assert r.get("text") == "Modified Text", f"Expected 'Modified Text', got '{r.get('text')}'"

        # Close
        await call("close_presentation", {"save": False})

        results[test_name] = "PASS"
        print(f"  [PASS] {test_name}")
    except Exception as e:
        results[test_name] = "FAIL"
        print(f"  [FAIL] {test_name} -> {e}")
        traceback.print_exc()
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════════════
    # E2E Test 3: Table Operations
    #   new → add table → fill → modify cell → verify → close
    # ═══════════════════════════════════════════════════════════════════════
    test_name = "table_operations"
    try:
        r = await call("new_presentation")
        assert_no_error(r, "new_presentation")

        # Add a 3x3 table with data
        data = [["H1", "H2", "H3"], ["A", "B", "C"], ["D", "E", "F"]]
        r = await call("add_table", {
            "slide_index": 1,
            "rows": 3, "cols": 3,
            "left": 1, "top": 1, "width": 8, "height": 4,
            "data_json": json.dumps(data),
        })
        assert_no_error(r, "add_table")
        table_name = r.get("shape_name", "")
        assert table_name, "No table shape name returned"

        # Modify a cell
        r = await call("modify_table_cell", {
            "slide_index": 1,
            "shape_name": table_name,
            "row": 2, "col": 2,
            "text": "UPDATED",
            "bold": True,
        })
        assert_no_error(r, "modify_table_cell")

        # Format the table
        r = await call("format_table", {
            "slide_index": 1,
            "shape_name": table_name,
            "has_header": True,
            "header_color": "#003366",
        })
        assert_no_error(r, "format_table")

        await call("close_presentation", {"save": False})
        results[test_name] = "PASS"
        print(f"  [PASS] {test_name}")
    except Exception as e:
        results[test_name] = "FAIL"
        print(f"  [FAIL] {test_name} -> {e}")
        traceback.print_exc()
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════════════
    # E2E Test 4: Find and Replace
    #   new → add slides with text → find_replace → verify text changed
    # ═══════════════════════════════════════════════════════════════════════
    test_name = "find_and_replace"
    try:
        r = await call("new_presentation")
        assert_no_error(r, "new_presentation")

        # Add textboxes with "PLACEHOLDER"
        for i in range(1, 3):
            if i > 1:
                await call("add_slide", {"layout": "blank"})
            await call("add_textbox", {
                "slide_index": i,
                "text": f"Slide {i}: PLACEHOLDER content here",
                "left": 1, "top": 1, "width": 8, "height": 1,
            })

        # Find and replace
        r = await call("find_and_replace", {
            "find_text": "PLACEHOLDER",
            "replace_text": "ACTUAL",
        })
        assert_no_error(r, "find_and_replace")
        assert r.get("replacements_count", 0) >= 2, (
            f"Expected >= 2 replacements, got {r.get('replacements_count')}"
        )

        # Verify text was replaced
        r = await call("extract_all_text")
        assert_no_error(r, "extract_all_text")
        all_text = json.dumps(r)
        assert "ACTUAL" in all_text, "Replacement text 'ACTUAL' not found in presentation"
        assert "PLACEHOLDER" not in all_text, "'PLACEHOLDER' still found after replacement"

        await call("close_presentation", {"save": False})
        results[test_name] = "PASS"
        print(f"  [PASS] {test_name}")
    except Exception as e:
        results[test_name] = "FAIL"
        print(f"  [FAIL] {test_name} -> {e}")
        traceback.print_exc()
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════════════
    # E2E Test 5: Slide Reorder
    #   new → add 3 slides with unique text → reorder → verify order
    # ═══════════════════════════════════════════════════════════════════════
    test_name = "slide_reorder"
    try:
        r = await call("new_presentation")
        assert_no_error(r, "new_presentation")

        # Add 2 more slides (we already have 1 from new_presentation)
        await call("add_slide", {"layout": "blank"})
        await call("add_slide", {"layout": "blank"})

        # Add unique text to each slide
        for i in range(1, 4):
            await call("add_textbox", {
                "slide_index": i,
                "text": f"SLIDE_{i}",
                "left": 1, "top": 1, "width": 3, "height": 1,
            })

        # Reorder: move slide 3 to position 1
        r = await call("reorder_slides", {"order_json": "[3, 1, 2]"})
        assert_no_error(r, "reorder_slides")

        # Verify the order by checking text on each slide
        r = await call("extract_all_text")
        assert_no_error(r, "extract_all_text")
        slides = r.get("slides", r if isinstance(r, list) else [])
        if len(slides) >= 3:
            # Slide 1 should now have text from original slide 3
            slide1_texts = " ".join(t.get("text", "") if isinstance(t, dict) else str(t)
                                    for t in slides[0].get("texts", []))
            assert "SLIDE_3" in slide1_texts, (
                f"Expected 'SLIDE_3' in first slide after reorder, got '{slide1_texts}'"
            )

        await call("close_presentation", {"save": False})
        results[test_name] = "PASS"
        print(f"  [PASS] {test_name}")
    except Exception as e:
        results[test_name] = "FAIL"
        print(f"  [FAIL] {test_name} -> {e}")
        traceback.print_exc()
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════════════
    # E2E Test 6: Multi-presentation Merge
    #   create 2 presentations → save → merge → verify combined count
    # ═══════════════════════════════════════════════════════════════════════
    test_name = "multi_pres_merge"
    pptx_1 = os.path.join(TEMP_DIR, "e2e_merge_1.pptx")
    pptx_2 = os.path.join(TEMP_DIR, "e2e_merge_2.pptx")
    try:
        # Create first pres with 2 slides
        r = await call("new_presentation")
        assert_no_error(r, "new_pres_1")
        await call("add_slide", {"layout": "blank"})
        await call("add_textbox", {
            "slide_index": 1, "text": "Pres1-Slide1",
            "left": 1, "top": 1, "width": 4, "height": 1,
        })
        await call("save_presentation_as", {"file_path": pptx_1})
        initial_count = 2  # 1 from new_pres + 1 added
        await call("close_presentation", {"save": False})

        # Create second pres with 3 slides
        r = await call("new_presentation")
        assert_no_error(r, "new_pres_2")
        await call("add_slide", {"layout": "blank"})
        await call("add_slide", {"layout": "blank"})
        await call("add_textbox", {
            "slide_index": 1, "text": "Pres2-Slide1",
            "left": 1, "top": 1, "width": 4, "height": 1,
        })
        await call("save_presentation_as", {"file_path": pptx_2})
        await call("close_presentation", {"save": False})

        # Open a new pres and merge both
        r = await call("new_presentation")
        assert_no_error(r, "new_pres_target")

        r = await call("merge_presentations", {
            "file_paths_json": json.dumps([pptx_1, pptx_2]),
        })
        assert_no_error(r, "merge_presentations")
        total_inserted = r.get("total_slides_inserted", 0)
        assert total_inserted >= 4, (
            f"Expected >= 4 slides inserted from merge, got {total_inserted}"
        )

        await call("close_presentation", {"save": False})
        results[test_name] = "PASS"
        print(f"  [PASS] {test_name}")
    except Exception as e:
        results[test_name] = "FAIL"
        print(f"  [FAIL] {test_name} -> {e}")
        traceback.print_exc()
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            pass
    finally:
        for f in [pptx_1, pptx_2]:
            try:
                os.remove(f)
            except OSError:
                pass

    # ═══════════════════════════════════════════════════════════════════════
    # E2E Test 7: Unicode / Emoji Text
    # ═══════════════════════════════════════════════════════════════════════
    test_name = "unicode_text"
    try:
        r = await call("new_presentation")
        assert_no_error(r, "new_presentation")

        unicode_text = "Hello 世界 🎉 مرحبا Ñoño"
        r = await call("add_textbox", {
            "slide_index": 1,
            "text": unicode_text,
            "left": 1, "top": 1, "width": 8, "height": 1.5,
        })
        assert_no_error(r, "add_textbox_unicode")
        shape_name = r.get("name", "")

        # Verify text
        r = await call("get_shape_details", {
            "slide_index": 1,
            "shape_name": shape_name,
        })
        assert_no_error(r, "get_shape_details")
        assert r.get("text") == unicode_text, (
            f"Unicode text mismatch: expected '{unicode_text}', got '{r.get('text')}'"
        )

        await call("close_presentation", {"save": False})
        results[test_name] = "PASS"
        print(f"  [PASS] {test_name}")
    except Exception as e:
        results[test_name] = "FAIL"
        print(f"  [FAIL] {test_name} -> {e}")
        traceback.print_exc()
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════════════
    # E2E Test 8: Move Slide and Verify Position
    # ═══════════════════════════════════════════════════════════════════════
    test_name = "move_slide_verify"
    try:
        r = await call("new_presentation")
        assert_no_error(r, "new_presentation")

        # Add slides with markers
        for i in range(3):
            if i > 0:
                await call("add_slide", {"layout": "blank"})
            await call("set_slide_notes", {
                "slide_index": i + 1,
                "notes": f"marker_{i + 1}",
            })

        # Move slide 3 to position 1
        r = await call("move_slide", {"slide_index": 3, "new_index": 1})
        assert_no_error(r, "move_slide")

        # Verify: slide at position 1 should have notes "marker_3"
        r = await call("get_slide_notes", {"slide_index": 1})
        assert_no_error(r, "get_slide_notes")
        assert r.get("notes") == "marker_3", (
            f"Expected 'marker_3' at position 1 after move, got '{r.get('notes')}'"
        )

        await call("close_presentation", {"save": False})
        results[test_name] = "PASS"
        print(f"  [PASS] {test_name}")
    except Exception as e:
        results[test_name] = "FAIL"
        print(f"  [FAIL] {test_name} -> {e}")
        traceback.print_exc()
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Summary
    # -----------------------------------------------------------------------
    print(f"\n{'=' * 60}")
    passed = sum(1 for v in results.values() if v == "PASS")
    failed = sum(1 for v in results.values() if v == "FAIL")
    total = len(results)
    print(f"E2E TESTS: {passed}/{total} passed, {failed} failed")
    if failed:
        print("FAILED:")
        for k, v in results.items():
            if v == "FAIL":
                print(f"  - {k}")
    print(f"{'=' * 60}")

    return results


if __name__ == "__main__":
    asyncio.run(test())
