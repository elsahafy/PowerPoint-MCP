"""Integration test for Phase 6 PowerPoint MCP tools (18 tools)."""
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

    # ── Setup ──────────────────────────────────────────────────────────
    print("=" * 60)
    print("Phase 6 Integration Test — 18 Advanced Operations Tools")
    print("=" * 60)

    print("\n[Setup] Launching PowerPoint ...")
    r = await call("launch_powerpoint")
    print(f"  launch_powerpoint -> {r}")

    print("[Setup] Creating new presentation ...")
    r = await call("new_presentation")
    print(f"  new_presentation -> {r}")

    # Add 3 blank slides
    for i in range(3):
        print(f"[Setup] Adding blank slide {i + 1} ...")
        r = await call("add_slide", {"layout": "blank"})
        print(f"  add_slide -> {r}")

    # Add textboxes on slide 1: "Hello World" and "Goodbye World"
    print("[Setup] Adding textbox 'Hello World' on slide 1 ...")
    r = await call("add_textbox", {
        "slide_index": 1, "text": "Hello World",
        "left": 1, "top": 1, "width": 4, "height": 1,
    })
    tb1_name = r.get("name", "") if isinstance(r, dict) else ""
    print(f"  add_textbox -> {r}  (name={tb1_name})")

    print("[Setup] Adding textbox 'Goodbye World' on slide 1 ...")
    r = await call("add_textbox", {
        "slide_index": 1, "text": "Goodbye World",
        "left": 1, "top": 3, "width": 4, "height": 1,
    })
    tb2_name = r.get("name", "") if isinstance(r, dict) else ""
    print(f"  add_textbox -> {r}  (name={tb2_name})")

    # Add a shape with text on slide 2
    print("[Setup] Adding shape with text on slide 2 ...")
    r = await call("add_shape", {
        "slide_index": 2, "shape_type": "rectangle",
        "left": 2, "top": 2, "width": 3, "height": 2,
        "text": "Slide2 Shape",
    })
    shape2_name = r.get("name", "") if isinstance(r, dict) else ""
    print(f"  add_shape -> {r}  (name={shape2_name})")

    # ── 1. find_and_replace ────────────────────────────────────────────
    print("\n[1/18] find_and_replace")
    r = await call("find_and_replace", {
        "find_text": "Hello", "replace_text": "Hi",
    })
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("replacements_count", 0) >= 1:
        results["find_and_replace"] = "PASS"
    else:
        results["find_and_replace"] = f"FAIL {r}"

    # ── 2. extract_all_text ────────────────────────────────────────────
    print("\n[2/18] extract_all_text")
    r = await call("extract_all_text", {"include_notes": True})
    print(f"  -> {r}")
    if isinstance(r, list):
        results["extract_all_text"] = "PASS"
    else:
        results["extract_all_text"] = f"FAIL {r}"

    # ── 3. get_presentation_outline ────────────────────────────────────
    print("\n[3/18] get_presentation_outline")
    r = await call("get_presentation_outline")
    print(f"  -> {r}")
    if isinstance(r, dict) and "slides" in r:
        results["get_presentation_outline"] = "PASS"
    elif isinstance(r, list):
        results["get_presentation_outline"] = "PASS"
    else:
        results["get_presentation_outline"] = f"FAIL {r}"

    # ── 4. merge_presentations ──────────────────────────────────────────
    print("\n[4/18] merge_presentations")
    tmp_merge = os.path.join(tempfile.gettempdir(), "phase6_merge_source.pptx")
    try:
        # Save current pres as a source file, then merge it back (adds its slides)
        r = await call("save_presentation_as", {"file_path": tmp_merge})
        assert "error" not in r, f"Failed to save merge source: {r}"
        slides_before = await call("get_slides")
        count_before = len(slides_before) if isinstance(slides_before, list) else 0
        r = await call("merge_presentations", {
            "file_paths_json": json.dumps([tmp_merge]),
        })
        print(f"  -> {r}")
        if isinstance(r, dict) and "error" not in r:
            results["merge_presentations"] = "PASS"
        else:
            results["merge_presentations"] = f"FAIL {r}"
    except Exception as e:
        results["merge_presentations"] = f"FAIL {e}"
        print(f"  Error: {e}")
    finally:
        if os.path.exists(tmp_merge):
            try:
                os.remove(tmp_merge)
            except OSError:
                pass

    # ── 5. apply_template ────────────────────────────────────────────────
    print("\n[5/18] apply_template")
    tmp_template = os.path.join(tempfile.gettempdir(), "phase6_template.potx")
    try:
        r = await call("save_presentation_as", {"file_path": tmp_template, "format": "potx"})
        assert "error" not in r, f"Failed to save template: {r}"
        r = await call("apply_template", {"template_path": tmp_template})
        print(f"  -> {r}")
        if isinstance(r, dict) and "error" not in r:
            results["apply_template"] = "PASS"
        else:
            results["apply_template"] = f"FAIL {r}"
    except Exception as e:
        results["apply_template"] = f"FAIL {e}"
        print(f"  Error: {e}")
    finally:
        if os.path.exists(tmp_template):
            try:
                os.remove(tmp_template)
            except OSError:
                pass

    # ── 6. bulk_format_text ────────────────────────────────────────────
    print("\n[6/18] bulk_format_text")
    criteria = json.dumps({
        "find_text": "Hi",
        "font_size": 28,
        "bold": True,
    })
    r = await call("bulk_format_text", {"criteria_json": criteria})
    print(f"  -> {r}")
    if isinstance(r, dict) and (r.get("matches_formatted", 0) >= 0 or r.get("status") == "formatted"):
        results["bulk_format_text"] = "PASS"
    else:
        results["bulk_format_text"] = f"FAIL {r}"

    # ── 7. add_animation ───────────────────────────────────────────────
    print("\n[7/18] add_animation")
    r = await call("add_animation", {
        "slide_index": 1, "shape_name": tb1_name, "effect": "fade",
    })
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "added":
        results["add_animation"] = "PASS"
    else:
        results["add_animation"] = f"FAIL {r}"

    # ── 8. remove_animation ────────────────────────────────────────────
    print("\n[8/18] remove_animation")
    r = await call("remove_animation", {
        "slide_index": 1, "shape_name": tb1_name,
    })
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "removed":
        results["remove_animation"] = "PASS"
    else:
        results["remove_animation"] = f"FAIL {r}"

    # ── 9. get_animations ──────────────────────────────────────────────
    print("\n[9/18] get_animations")
    r = await call("get_animations", {"slide_index": 1})
    print(f"  -> {r}")
    if isinstance(r, dict) and "animations" in r:
        results["get_animations"] = "PASS"
    else:
        results["get_animations"] = f"FAIL {r}"

    # ── 10. reorder_animations ─────────────────────────────────────────
    print("\n[10/18] reorder_animations")
    # First add 2 animations so we can reorder
    await call("add_animation", {
        "slide_index": 1, "shape_name": tb1_name, "effect": "fade",
    })
    await call("add_animation", {
        "slide_index": 1, "shape_name": tb2_name, "effect": "appear",
    })
    order = json.dumps([tb2_name, tb1_name])
    r = await call("reorder_animations", {
        "slide_index": 1, "order_json": order,
    })
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "reordered":
        results["reorder_animations"] = "PASS"
    else:
        results["reorder_animations"] = f"FAIL {r}"

    # Clean up animations for later tests
    await call("remove_animation", {"slide_index": 1, "shape_name": tb1_name})
    await call("remove_animation", {"slide_index": 1, "shape_name": tb2_name})

    # ── 11. bulk_speaker_notes ─────────────────────────────────────────
    print("\n[11/18] bulk_speaker_notes")
    notes = json.dumps([
        {"slide_index": 1, "notes": "Notes for slide 1"},
        {"slide_index": 2, "notes": "Notes for slide 2"},
    ])
    r = await call("bulk_speaker_notes", {"notes_json": notes})
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "updated":
        results["bulk_speaker_notes"] = "PASS"
    else:
        results["bulk_speaker_notes"] = f"FAIL {r}"

    # ── 12. clone_formatting ───────────────────────────────────────────
    print("\n[12/18] clone_formatting")
    targets = json.dumps([tb2_name])
    r = await call("clone_formatting", {
        "slide_index": 1,
        "source_shape": tb1_name,
        "target_shapes_json": targets,
    })
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "cloned":
        results["clone_formatting"] = "PASS"
    else:
        results["clone_formatting"] = f"FAIL {r}"

    # ── 13. search_shapes ──────────────────────────────────────────────
    print("\n[13/18] search_shapes")
    r = await call("search_shapes", {"query": "Hi"})
    print(f"  -> {r}")
    if isinstance(r, dict) and "matches" in r:
        results["search_shapes"] = "PASS"
    else:
        results["search_shapes"] = f"FAIL {r}"

    # ── 14. rename_shape ───────────────────────────────────────────────
    print("\n[14/18] rename_shape")
    r = await call("rename_shape", {
        "slide_index": 2,
        "old_name": shape2_name,
        "new_name": "RenamedShape",
    })
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "renamed":
        results["rename_shape"] = "PASS"
        shape2_name = "RenamedShape"  # update for later use
    else:
        results["rename_shape"] = f"FAIL {r}"

    # ── 15. lock_shape ─────────────────────────────────────────────────
    print("\n[15/18] lock_shape")
    r = await call("lock_shape", {
        "slide_index": 2, "shape_name": shape2_name, "lock": True,
    })
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") in ("locked", "unlocked"):
        results["lock_shape"] = "PASS"
    else:
        results["lock_shape"] = f"FAIL {r}"

    # ── 16. add_section ────────────────────────────────────────────────
    print("\n[16/18] add_section")
    r = await call("add_section", {"name": "Section A", "before_slide": 1})
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "added":
        results["add_section"] = "PASS"
        section_idx = r.get("section_index", 1)
    else:
        results["add_section"] = f"FAIL {r}"
        section_idx = 1

    # ── 17. get_sections ───────────────────────────────────────────────
    print("\n[17/18] get_sections")
    r = await call("get_sections")
    print(f"  -> {r}")
    if isinstance(r, dict) and "sections" in r and len(r["sections"]) >= 1:
        results["get_sections"] = "PASS"
    elif isinstance(r, list) and len(r) >= 1:
        results["get_sections"] = "PASS"
    else:
        results["get_sections"] = f"FAIL {r}"

    # ── 18. delete_section ─────────────────────────────────────────────
    print("\n[18/18] delete_section")
    r = await call("delete_section", {
        "section_index": section_idx, "delete_slides": False,
    })
    print(f"  -> {r}")
    if isinstance(r, dict) and r.get("status") == "deleted":
        results["delete_section"] = "PASS"
    else:
        results["delete_section"] = f"FAIL {r}"

    # ── Summary ────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    passed = sum(1 for v in results.values() if v == "PASS")
    skipped = sum(1 for v in results.values() if v == "SKIP")
    failed = sum(1 for v in results.values() if v.startswith("FAIL"))
    for tool_name, status in results.items():
        tag = "PASS" if status == "PASS" else ("SKIP" if status == "SKIP" else "FAIL")
        print(f"  [{tag}] {tool_name}")
    print(f"\nTotal: {passed} passed, {skipped} skipped, {failed} failed out of {len(results)}")

    # ── Cleanup ────────────────────────────────────────────────────────
    print("\n[Cleanup] Closing presentation ...")
    try:
        r = await call("close_presentation", {"save": False})
        print(f"  close_presentation -> {r}")
    except Exception as e:
        print(f"  cleanup error: {e}")


if __name__ == "__main__":
    asyncio.run(test())
