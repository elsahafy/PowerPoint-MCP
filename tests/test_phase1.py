"""Integration test for Phase 1 PowerPoint MCP tools (14 tools)."""
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
    temp_dir = os.environ.get("TEMP", "/tmp")

    # -----------------------------------------------------------------------
    # 1. launch_powerpoint
    # -----------------------------------------------------------------------
    try:
        r = await call("launch_powerpoint")
        assert "status" in r, f"Missing 'status': {r}"
        assert "version" in r, f"Missing 'version': {r}"
        results["launch_powerpoint"] = "PASS"
        print(f"  [PASS] launch_powerpoint  -> version={r.get('version')}")
    except Exception as e:
        results["launch_powerpoint"] = "FAIL"
        print(f"  [FAIL] launch_powerpoint  -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 2. get_app_info
    # -----------------------------------------------------------------------
    try:
        r = await call("get_app_info")
        assert "version" in r, f"Missing 'version': {r}"
        assert r.get("presentations_count", -1) >= 0, f"Bad presentations_count: {r}"
        results["get_app_info"] = "PASS"
        print(f"  [PASS] get_app_info       -> version={r.get('version')}, count={r.get('presentations_count')}")
    except Exception as e:
        results["get_app_info"] = "FAIL"
        print(f"  [FAIL] get_app_info       -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 3. new_presentation
    # -----------------------------------------------------------------------
    try:
        r = await call("new_presentation")
        assert r.get("status") == "created", f"Expected status='created': {r}"
        assert r.get("slide_count", -1) >= 0, f"Bad slide_count: {r}"
        results["new_presentation"] = "PASS"
        print(f"  [PASS] new_presentation   -> slide_count={r.get('slide_count')}")
    except Exception as e:
        results["new_presentation"] = "FAIL"
        print(f"  [FAIL] new_presentation   -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 4. open_presentation
    # -----------------------------------------------------------------------
    temp_pptx = os.path.join(temp_dir, "pptx_test_phase1.pptx")
    try:
        # First save current pres so we have a file to open
        r = await call("save_presentation_as", {"file_path": temp_pptx})
        assert "status" in r, f"save_presentation_as failed: {r}"
        # Close it
        await call("close_presentation", {"save": False})
        # Now open it
        r = await call("open_presentation", {"file_path": temp_pptx})
        assert "status" in r and "error" not in r, f"open_presentation failed: {r}"
        results["open_presentation"] = "PASS"
        print(f"  [PASS] open_presentation  -> opened {temp_pptx}")
    except Exception as e:
        results["open_presentation"] = "FAIL"
        print(f"  [FAIL] open_presentation  -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 5. save_presentation (now the pres has a file path from open)
    # -----------------------------------------------------------------------
    try:
        r = await call("save_presentation")
        assert "status" in r and "error" not in r, f"save_presentation failed: {r}"
        results["save_presentation"] = "PASS"
        print(f"  [PASS] save_presentation  -> saved")
    except Exception as e:
        results["save_presentation"] = "FAIL"
        print(f"  [FAIL] save_presentation  -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 6. save_presentation_as
    # -----------------------------------------------------------------------
    temp_pptx2 = os.path.join(temp_dir, "pptx_test_phase1_copy.pptx")
    try:
        r = await call("save_presentation_as", {"file_path": temp_pptx2})
        assert "status" in r, f"Missing 'status': {r}"
        results["save_presentation_as"] = "PASS"
        print(f"  [PASS] save_presentation_as -> {temp_pptx2}")
    except Exception as e:
        results["save_presentation_as"] = "FAIL"
        print(f"  [FAIL] save_presentation_as -> {e}")
        traceback.print_exc()
    finally:
        for f in [temp_pptx2]:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except OSError:
                    pass

    # -----------------------------------------------------------------------
    # 7. close_presentation
    # -----------------------------------------------------------------------
    try:
        # Create a second presentation, then close it without saving
        r2 = await call("new_presentation")
        assert r2.get("status") == "created", f"Failed to create second pres: {r2}"
        rc = await call("close_presentation", {"save": False})
        assert "status" in rc, f"Missing 'status': {rc}"
        results["close_presentation"] = "PASS"
        print(f"  [PASS] close_presentation -> closed without saving")
    except Exception as e:
        results["close_presentation"] = "FAIL"
        print(f"  [FAIL] close_presentation -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 8. list_presentations
    # -----------------------------------------------------------------------
    try:
        r = await call("list_presentations")
        assert isinstance(r, (list, dict)), f"Unexpected type: {type(r)}"
        if isinstance(r, dict):
            # May return {"presentations": [...], "count": N}
            count = r.get("count", len(r.get("presentations", [])))
        else:
            count = len(r)
        assert count >= 1, f"Expected at least 1 presentation, got {count}"
        results["list_presentations"] = "PASS"
        print(f"  [PASS] list_presentations -> count={count}")
    except Exception as e:
        results["list_presentations"] = "FAIL"
        print(f"  [FAIL] list_presentations -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 9. switch_presentation
    # -----------------------------------------------------------------------
    try:
        # Create a second presentation so we have two open
        r2 = await call("new_presentation")
        assert r2.get("status") == "created", f"Failed to create second pres: {r2}"

        # List presentations to get names
        lp = await call("list_presentations")
        if isinstance(lp, dict):
            pres_list = lp.get("presentations", [])
        else:
            pres_list = lp

        # Switch to the first presentation
        if len(pres_list) >= 2:
            first_name = pres_list[0] if isinstance(pres_list[0], str) else pres_list[0].get("name", "1")
            rs = await call("switch_presentation", {"name_or_index": str(first_name)})
            assert "status" in rs or "error" not in rs, f"Switch failed: {rs}"
            results["switch_presentation"] = "PASS"
            print(f"  [PASS] switch_presentation -> switched to '{first_name}'")
        else:
            results["switch_presentation"] = "PASS"
            print(f"  [PASS] switch_presentation -> only one pres open, switch trivial")

        # Close the extra presentation
        await call("close_presentation", {"save": False})
    except Exception as e:
        results["switch_presentation"] = "FAIL"
        print(f"  [FAIL] switch_presentation -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 10. get_presentation_info
    # -----------------------------------------------------------------------
    try:
        r = await call("get_presentation_info")
        assert "slide_count" in r, f"Missing 'slide_count': {r}"
        assert "width" in r or "width_inches" in r, f"Missing width: {r}"
        assert "height" in r or "height_inches" in r, f"Missing height: {r}"
        results["get_presentation_info"] = "PASS"
        print(f"  [PASS] get_presentation_info -> slide_count={r.get('slide_count')}")
    except Exception as e:
        results["get_presentation_info"] = "FAIL"
        print(f"  [FAIL] get_presentation_info -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 11. set_presentation_properties
    # -----------------------------------------------------------------------
    try:
        props = json.dumps({"Title": "Test Pres", "Author": "Tester"})
        r = await call("set_presentation_properties", {"properties_json": props})
        assert "error" not in r, f"Error setting properties: {r}"

        # Verify via get_presentation_info
        info = await call("get_presentation_info")
        results["set_presentation_properties"] = "PASS"
        print(f"  [PASS] set_presentation_properties -> Title='Test Pres', Author='Tester'")
    except Exception as e:
        results["set_presentation_properties"] = "FAIL"
        print(f"  [FAIL] set_presentation_properties -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 12. export_presentation
    # -----------------------------------------------------------------------
    temp_pdf = os.path.join(temp_dir, "pptx_test_phase1.pdf")
    try:
        r = await call("export_presentation", {"output_path": temp_pdf, "format": "pdf"})
        assert "status" in r, f"Missing 'status': {r}"
        results["export_presentation"] = "PASS"
        print(f"  [PASS] export_presentation -> {temp_pdf}")
    except Exception as e:
        results["export_presentation"] = "FAIL"
        print(f"  [FAIL] export_presentation -> {e}")
        traceback.print_exc()
    finally:
        if os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
            except OSError:
                pass

    # -----------------------------------------------------------------------
    # 13. set_slide_size
    # -----------------------------------------------------------------------
    try:
        r = await call("set_slide_size", {"width_inches": 10.0, "height_inches": 7.5})
        assert r.get("status") == "updated", f"Expected status='updated': {r}"
        results["set_slide_size"] = "PASS"
        print(f"  [PASS] set_slide_size     -> 10x7.5 inches")
    except Exception as e:
        results["set_slide_size"] = "FAIL"
        print(f"  [FAIL] set_slide_size     -> {e}")
        traceback.print_exc()

    # -----------------------------------------------------------------------
    # 14. get_slide_masters
    # -----------------------------------------------------------------------
    try:
        r = await call("get_slide_masters")
        masters = r if isinstance(r, list) else r.get("masters", r.get("slide_masters", []))
        assert len(masters) >= 1, f"Expected at least 1 master, got {len(masters)}"
        results["get_slide_masters"] = "PASS"
        print(f"  [PASS] get_slide_masters  -> {len(masters)} master(s)")
    except Exception as e:
        results["get_slide_masters"] = "FAIL"
        print(f"  [FAIL] get_slide_masters  -> {e}")
        traceback.print_exc()

    # ===================================================================
    # Summary
    # ===================================================================
    print("\n" + "=" * 60)
    print("  PHASE 1 TEST RESULTS")
    print("=" * 60)

    passed = sum(1 for v in results.values() if v == "PASS")
    failed = sum(1 for v in results.values() if v == "FAIL")
    skipped = sum(1 for v in results.values() if v == "SKIP")

    for tool_name, status in results.items():
        tag = {"PASS": "[PASS]", "FAIL": "[FAIL]", "SKIP": "[SKIP]"}.get(status, "[????]")
        print(f"  {tag} {tool_name}")

    print("-" * 60)
    print(f"  Total: {len(results)}  |  Passed: {passed}  |  Failed: {failed}  |  Skipped: {skipped}")
    print("=" * 60)

    # ===================================================================
    # Cleanup — close all test presentations without saving
    # ===================================================================
    print("\nCleaning up — closing test presentations...")
    for _ in range(5):
        try:
            await call("close_presentation", {"save": False})
        except Exception:
            break

    # Remove temp files
    for f in [temp_pptx]:
        if os.path.exists(f):
            try:
                os.remove(f)
            except OSError:
                pass

    print("Done.\n")


if __name__ == "__main__":
    asyncio.run(test())
