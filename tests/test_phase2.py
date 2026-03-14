"""Integration test for Phase 2 PowerPoint MCP tools (16 tools)."""
import asyncio
import json
import importlib.util
import os

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

    # ── Setup ─────────────────────────────────────────────────────────────────
    print("=== Phase 2 Integration Test: 16 Slide Operation Tools ===\n")

    print("[SETUP] Launching PowerPoint...")
    r = await call("launch_powerpoint")
    print(f"  launch_powerpoint -> {r}")

    print("[SETUP] Creating new presentation...")
    r = await call("new_presentation")
    print(f"  new_presentation -> {r}\n")

    # ── 1. get_slides ─────────────────────────────────────────────────────────
    print("[1/16] get_slides")
    r = await call("get_slides")
    assert isinstance(r, list), f"Expected list, got {type(r)}: {r}"
    results["get_slides"] = "PASS"
    print(f"  -> {len(r)} slide(s) returned. PASS\n")

    # ── 2. get_slide_info ─────────────────────────────────────────────────────
    print("[2/16] get_slide_info")
    r = await call("get_slide_info", {"slide_index": 1})
    assert "index" in r, f"Expected 'index' key in response: {r}"
    results["get_slide_info"] = "PASS"
    print(f"  -> slide info keys: {list(r.keys()) if isinstance(r, dict) else r}. PASS\n")

    # ── 3. add_slide ──────────────────────────────────────────────────────────
    print("[3/16] add_slide")
    r = await call("add_slide", {"layout": "blank"})
    assert isinstance(r, dict) and "index" in r, f"Expected dict with 'index': {r}"
    results["add_slide"] = "PASS"
    print(f"  -> added slide at index {r.get('index')}. PASS\n")

    # ── 4. duplicate_slide ────────────────────────────────────────────────────
    print("[4/16] duplicate_slide")
    r = await call("duplicate_slide", {"slide_index": 1})
    assert isinstance(r, dict) and "index" in r, f"Expected dict with 'index': {r}"
    results["duplicate_slide"] = "PASS"
    print(f"  -> {r}. PASS\n")

    # ── 5. delete_slide ───────────────────────────────────────────────────────
    print("[5/16] delete_slide")
    # Get current slide count to delete the last one
    slides = await call("get_slides")
    last_idx = len(slides)
    r = await call("delete_slide", {"slide_index": last_idx})
    assert isinstance(r, dict) and r.get("status") == "deleted", f"Unexpected: {r}"
    results["delete_slide"] = "PASS"
    print(f"  -> deleted slide {last_idx}. PASS\n")

    # ── 6. move_slide ─────────────────────────────────────────────────────────
    print("[6/16] move_slide")
    # Add 2 more slides so we have enough to move
    await call("add_slide", {"layout": "blank"})
    await call("add_slide", {"layout": "blank"})
    slides = await call("get_slides")
    total = len(slides)
    print(f"  -> {total} slides before move")
    r = await call("move_slide", {"slide_index": 1, "new_index": total})
    assert isinstance(r, dict) and r.get("status") == "moved", f"Unexpected: {r}"
    results["move_slide"] = "PASS"
    print(f"  -> moved slide 1 to position {total}. PASS\n")

    # ── 7. copy_slide ─────────────────────────────────────────────────────────
    print("[7/16] copy_slide")
    r = await call("copy_slide", {"source_index": 1})
    assert isinstance(r, dict) and "status" in r, f"Expected dict with 'status': {r}"
    results["copy_slide"] = "PASS"
    print(f"  -> {r}. PASS\n")

    # ── 8. get_slide_notes ────────────────────────────────────────────────────
    print("[8/16] get_slide_notes")
    r = await call("get_slide_notes", {"slide_index": 1})
    assert isinstance(r, dict) and "notes" in r, f"Expected dict with 'notes': {r}"
    results["get_slide_notes"] = "PASS"
    print(f"  -> notes: {repr(r.get('notes', '')[:50])}. PASS\n")

    # ── 9. set_slide_notes ────────────────────────────────────────────────────
    print("[9/16] set_slide_notes")
    r = await call("set_slide_notes", {"slide_index": 1, "notes": "Test notes"})
    assert isinstance(r, dict) and "status" in r, f"Expected dict with 'status': {r}"
    # Verify notes were set
    verify = await call("get_slide_notes", {"slide_index": 1})
    assert verify.get("notes") == "Test notes", f"Notes mismatch: {verify}"
    results["set_slide_notes"] = "PASS"
    print(f"  -> set and verified notes. PASS\n")

    # ── 10. set_slide_layout ──────────────────────────────────────────────────
    print("[10/16] set_slide_layout")
    r = await call("set_slide_layout", {"slide_index": 1, "layout": "title"})
    assert isinstance(r, dict) and r.get("status") == "updated", f"Unexpected: {r}"
    results["set_slide_layout"] = "PASS"
    print(f"  -> {r}. PASS\n")

    # ── 11. set_slide_transition ──────────────────────────────────────────────
    print("[11/16] set_slide_transition")
    r = await call("set_slide_transition", {"slide_index": 1, "effect": "fade"})
    assert isinstance(r, dict) and r.get("status") == "updated", f"Unexpected: {r}"
    results["set_slide_transition"] = "PASS"
    print(f"  -> {r}. PASS\n")

    # ── 12. bulk_add_slides ───────────────────────────────────────────────────
    print("[12/16] bulk_add_slides")
    slides_spec = json.dumps([
        {"layout": "blank"},
        {"layout": "title"},
        {"layout": "blank"},
    ])
    r = await call("bulk_add_slides", {"slides_json": slides_spec})
    assert isinstance(r, list), f"Expected list of results: {r}"
    results["bulk_add_slides"] = "PASS"
    print(f"  -> added {len(r)} slides. PASS\n")

    # ── 13. reorder_slides ────────────────────────────────────────────────────
    print("[13/16] reorder_slides")
    slides = await call("get_slides")
    count = len(slides)
    order = list(range(count, 0, -1))  # reverse order
    r = await call("reorder_slides", {"order_json": json.dumps(order)})
    assert isinstance(r, dict) and "status" in r, f"Expected dict with 'status': {r}"
    results["reorder_slides"] = "PASS"
    print(f"  -> reordered {count} slides. PASS\n")

    # ── 14. get_slide_layout_names ────────────────────────────────────────────
    print("[14/16] get_slide_layout_names")
    r = await call("get_slide_layout_names")
    assert isinstance(r, list) and len(r) >= 1, f"Expected non-empty list: {r}"
    results["get_slide_layout_names"] = "PASS"
    print(f"  -> {len(r)} layouts: {r[:5]}... PASS\n")

    # ── 15. set_slide_background ──────────────────────────────────────────────
    print("[15/16] set_slide_background")
    r = await call("set_slide_background", {"slide_index": 1, "color": "#FF0000"})
    assert isinstance(r, dict) and r.get("status") == "updated", f"Unexpected: {r}"
    results["set_slide_background"] = "PASS"
    print(f"  -> {r}. PASS\n")

    # ── 16. bulk_set_transitions ──────────────────────────────────────────────
    print("[16/16] bulk_set_transitions")
    transitions_spec = json.dumps([
        {"slide_index": 1, "effect": "fade", "duration": 1.0},
        {"slide_index": 2, "effect": "push", "duration": 0.5},
    ])
    r = await call("bulk_set_transitions", {"settings_json": transitions_spec})
    assert isinstance(r, list), f"Expected list of results: {r}"
    results["bulk_set_transitions"] = "PASS"
    print(f"  -> applied transitions to {len(r)} slides. PASS\n")

    # ── Summary ───────────────────────────────────────────────────────────────
    print("=" * 60)
    print("RESULTS SUMMARY")
    print("=" * 60)
    passed = sum(1 for v in results.values() if v == "PASS")
    total = len(results)
    for tool_name, status in results.items():
        print(f"  {status}  {tool_name}")
    print(f"\n  {passed}/{total} tools passed")
    print("=" * 60)

    # ── Cleanup ───────────────────────────────────────────────────────────────
    print("\n[CLEANUP] Closing presentation without saving...")
    await call("close_presentation", {"save": False})
    print("Done.")


if __name__ == "__main__":
    asyncio.run(test())
