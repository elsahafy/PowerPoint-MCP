"""Integration test for Phase 5 PowerPoint MCP tools (12 tools)."""
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
    print("Phase 5 Integration Test — 12 Design & Themes Tools")
    print("=" * 60)

    print("\n[Setup] Launching PowerPoint …")
    r = await call("launch_powerpoint")
    print(f"  launch_powerpoint → {r}")

    print("[Setup] Creating new presentation …")
    r = await call("new_presentation")
    print(f"  new_presentation → {r}")

    print("[Setup] Adding blank slide (slide 1) …")
    r = await call("add_slide", {"layout_index": 7})  # blank layout
    print(f"  add_slide (blank) → {r}")

    print("[Setup] Adding title slide (slide 2) …")
    r = await call("add_slide", {"layout_index": 1})  # title layout
    print(f"  add_slide (title) → {r}")

    # ── 1. get_theme_info ──────────────────────────────────────────────
    print("\n[1/12] get_theme_info")
    r = await call("get_theme_info")
    print(f"  → {r}")
    if isinstance(r, dict) and "error" not in r:
        results["get_theme_info"] = "PASS"
    else:
        results["get_theme_info"] = "FAIL"
    print(f"  Result: {results['get_theme_info']}")

    # ── 2. apply_theme ──────────────────────────────────────────────────
    print("\n[2/12] apply_theme")
    try:
        # Use a built-in Office theme from the installation directory
        import glob
        office_themes = glob.glob(r"C:\Program Files*\Microsoft Office\Root\Document Themes*\**\*.thmx", recursive=True)
        if not office_themes:
            office_themes = glob.glob(os.path.expandvars(r"%APPDATA%\Microsoft\Templates\Document Themes\**\*.thmx"), recursive=True)
        if office_themes:
            theme_path = office_themes[0]
            r = await call("apply_theme", {"theme_path": theme_path})
            print(f"  → {r}")
            if isinstance(r, dict) and "error" not in r:
                results["apply_theme"] = "PASS"
            else:
                results["apply_theme"] = "FAIL"
        else:
            # No theme files found on this system — test the tool still returns clean error
            r = await call("apply_theme", {"theme_path": "nonexistent.thmx"})
            print(f"  → {r}")
            if isinstance(r, dict) and "error" in r:
                results["apply_theme"] = "PASS"  # graceful error is acceptable
            else:
                results["apply_theme"] = "FAIL"
    except Exception as e:
        results["apply_theme"] = "FAIL"
        print(f"  Error: {e}")
    print(f"  Result: {results['apply_theme']}")

    # ── 3. get_theme_colors ────────────────────────────────────────────
    print("\n[3/12] get_theme_colors")
    r = await call("get_theme_colors")
    print(f"  → {r}")
    if isinstance(r, list) and len(r) > 0:
        results["get_theme_colors"] = "PASS"
    else:
        results["get_theme_colors"] = "FAIL"
    print(f"  Result: {results['get_theme_colors']}")

    # ── 4. set_theme_color ─────────────────────────────────────────────
    print("\n[4/12] set_theme_color (Accent1 → #FF5500)")
    r = await call("set_theme_color", {"slot": "Accent1", "color": "#FF5500"})
    print(f"  → {r}")
    if isinstance(r, dict) and r.get("status") == "updated":
        results["set_theme_color"] = "PASS"
    else:
        results["set_theme_color"] = "FAIL"
    print(f"  Result: {results['set_theme_color']}")

    # ── 5. get_theme_fonts ─────────────────────────────────────────────
    print("\n[5/12] get_theme_fonts")
    r = await call("get_theme_fonts")
    print(f"  → {r}")
    if isinstance(r, dict) and ("major_font" in r or "minor_font" in r):
        results["get_theme_fonts"] = "PASS"
    else:
        results["get_theme_fonts"] = "FAIL"
    print(f"  Result: {results['get_theme_fonts']}")

    # ── 6. set_theme_fonts ─────────────────────────────────────────────
    print("\n[6/12] set_theme_fonts (major_font=Arial)")
    r = await call("set_theme_fonts", {"major_font": "Arial"})
    print(f"  → {r}")
    if isinstance(r, dict) and "status" in r and "error" not in r:
        results["set_theme_fonts"] = "PASS"
    else:
        results["set_theme_fonts"] = "FAIL"
    print(f"  Result: {results['set_theme_fonts']}")

    # ── 7. get_master_layouts ──────────────────────────────────────────
    print("\n[7/12] get_master_layouts (master_index=1)")
    r = await call("get_master_layouts", {"master_index": 1})
    print(f"  → {r}")
    if isinstance(r, list) and len(r) > 0:
        results["get_master_layouts"] = "PASS"
    else:
        results["get_master_layouts"] = "FAIL"
    print(f"  Result: {results['get_master_layouts']}")

    # ── 8. modify_master_placeholder ─────────────────────────────────────
    print("\n[8/12] modify_master_placeholder")
    try:
        r = await call("modify_master_placeholder", {
            "master_index": 1,
            "layout_index": 1,
            "placeholder_index": 1,
            "font_size": 36,
            "font_name": "Arial",
        })
        print(f"  → {r}")
        if isinstance(r, dict) and "error" not in r:
            results["modify_master_placeholder"] = "PASS"
        else:
            results["modify_master_placeholder"] = "FAIL"
    except Exception as e:
        results["modify_master_placeholder"] = "FAIL"
        print(f"  Error: {e}")
    print(f"  Result: {results['modify_master_placeholder']}")

    # ── 9. set_background ──────────────────────────────────────────────
    print("\n[9/12] set_background (slide 1, color=#336699)")
    r = await call("set_background", {"slide_index": 1, "color": "#336699"})
    print(f"  → {r}")
    if isinstance(r, dict) and "status" in r and "error" not in r:
        results["set_background"] = "PASS"
    else:
        results["set_background"] = "FAIL"
    print(f"  Result: {results['set_background']}")

    # ── 10. get_placeholders ───────────────────────────────────────────
    print("\n[10/12] get_placeholders")
    try:
        # Use slide 1 (always exists) to avoid index issues
        r = await call("get_placeholders", {"slide_index": 1})
        print(f"  → {r}")
        if isinstance(r, list):
            results["get_placeholders"] = "PASS"
        else:
            results["get_placeholders"] = "FAIL"
    except Exception as e:
        results["get_placeholders"] = "FAIL"
        print(f"  Error: {e}")
    print(f"  Result: {results['get_placeholders']}")

    # ── 11. add_custom_layout ──────────────────────────────────────────
    print("\n[11/12] add_custom_layout (name='Test Layout')")
    r = await call("add_custom_layout", {"master_index": 1, "name": "Test Layout"})
    print(f"  → {r}")
    if isinstance(r, dict) and r.get("status") == "created":
        results["add_custom_layout"] = "PASS"
    else:
        results["add_custom_layout"] = "FAIL"
    print(f"  Result: {results['add_custom_layout']}")

    # ── 12. copy_master_from ─────────────────────────────────────────────
    print("\n[12/12] copy_master_from")
    tmp_source = os.path.join(tempfile.gettempdir(), "phase5_source.pptx")
    try:
        # Save current pres as source file to copy master from
        r = await call("save_presentation_as", {"file_path": tmp_source})
        assert "error" not in r, f"Failed to save source: {r}"
        r = await call("copy_master_from", {"source_path": tmp_source})
        print(f"  → {r}")
        if isinstance(r, dict) and "error" not in r:
            results["copy_master_from"] = "PASS"
        else:
            results["copy_master_from"] = "FAIL"
    except Exception as e:
        results["copy_master_from"] = "FAIL"
        print(f"  Error: {e}")
    finally:
        if os.path.exists(tmp_source):
            try:
                os.remove(tmp_source)
            except OSError:
                pass
    print(f"  Result: {results['copy_master_from']}")

    # ── Summary ────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    passed = sum(1 for v in results.values() if v == "PASS")
    failed = sum(1 for v in results.values() if v == "FAIL")
    skipped = sum(1 for v in results.values() if v == "SKIP")
    for tool, status in results.items():
        tag = {"PASS": "PASS", "FAIL": "FAIL", "SKIP": "SKIP"}[status]
        print(f"  [{tag}] {tool}")
    print(f"\nTotal: {passed} PASS / {failed} FAIL / {skipped} SKIP out of {len(results)}")
    print("=" * 60)

    # ── Cleanup ────────────────────────────────────────────────────────
    print("\n[Cleanup] Closing presentation without saving …")
    try:
        r = await call("close_presentation", {"save": False})
        print(f"  close_presentation → {r}")
    except Exception as e:
        print(f"  close_presentation error (non-fatal): {e}")

    if failed > 0:
        raise SystemExit(f"{failed} tool(s) FAILED")


if __name__ == "__main__":
    asyncio.run(test())
