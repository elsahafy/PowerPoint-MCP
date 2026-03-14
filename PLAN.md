# Enterprise PowerPoint MCP Server - Implementation Plan

## Context

Build an enterprise-grade PowerPoint MCP server at `C:\Users\Ibrahim Elsahafy\mcp-servers\powerpoint\server.py` following the same architecture as the existing MS Project MCP server (`msproject/server.py`). The server will control local Microsoft PowerPoint via Windows COM automation, exposing **105 tools** across 7 phases for comprehensive presentation management.

---

## Architecture

| Aspect | Decision |
|--------|----------|
| **Language** | Python (single-file `server.py`) |
| **Framework** | `mcp.server.fastmcp.FastMCP` |
| **COM Target** | `win32com.client` → `PowerPoint.Application` |
| **Return Format** | All tools return `json.dumps({...}, indent=2)` |
| **Entry Point** | `mcp.run()` at bottom of file |
| **Dependencies** | `mcp`, `pywin32` |

---

## Files to Create

| File | Purpose |
|------|---------|
| `powerpoint/server.py` | Main server (~5,000 lines across all phases) |
| `powerpoint/.gitignore` | Exclude `__pycache__`, `*.pyc`, `.env` |
| `powerpoint/README.md` | Tool inventory & setup instructions |
| `powerpoint/LICENSE` | MIT License |
| `powerpoint/tests/test_phase1.py` | Phase 1 integration tests |
| `powerpoint/tests/test_phase2.py` | Phase 2 integration tests |
| `powerpoint/tests/test_phase3.py` | Phase 3 integration tests |
| `powerpoint/tests/test_phase4.py` | Phase 4 integration tests |
| `powerpoint/tests/test_phase5.py` | Phase 5 integration tests |
| `powerpoint/tests/test_phase6.py` | Phase 6 integration tests |
| `powerpoint/tests/test_phase7.py` | Phase 7 integration tests |

## Reference Files

| File | Purpose |
|------|---------|
| `msproject/server.py` | Pattern for helpers, tool decorators, JSON returns, entry point |
| `msproject/tests/test_phase2.py` | Pattern for integration tests |
| `msproject/README.md` | Pattern for documentation structure |

---

## Helper Functions

These sit at the top of `server.py`, before any `@mcp.tool()` definitions.

### Core COM Helpers

```python
def get_app(require_presentation=True):
    """Get running PowerPoint instance. Launches if not running."""
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
    if require_presentation and app.Presentations.Count == 0:
        raise RuntimeError("No presentation is open in PowerPoint.")
    return app

def get_pres(app):
    """Get the active presentation."""
    return app.ActivePresentation

def get_slide(pres, slide_index: int):
    """Get a slide by 1-based index. Raises on invalid index."""
    if slide_index < 1 or slide_index > pres.Slides.Count:
        raise ValueError(f"Slide index {slide_index} out of range (1-{pres.Slides.Count}).")
    return pres.Slides(slide_index)

def get_shape(slide, shape_id):
    """Find shape by name or 1-based index."""
    if isinstance(shape_id, int):
        return slide.Shapes(shape_id)
    for shp in slide.Shapes:
        if shp.Name == shape_id:
            return shp
    raise ValueError(f"Shape '{shape_id}' not found on slide.")
```

### Data Conversion Helpers

```python
def shape_to_dict(shape):
    """Convert COM Shape to dict with common properties."""
    d = {
        "name": shape.Name,
        "shape_id": shape.Id,
        "type": shape.Type,
        "type_name": _shape_type_name(shape.Type),
        "left": round(shape.Left, 2),
        "top": round(shape.Top, 2),
        "width": round(shape.Width, 2),
        "height": round(shape.Height, 2),
        "rotation": round(shape.Rotation, 2),
        "has_text": shape.HasTextFrame == -1,
    }
    if d["has_text"]:
        try:
            d["text"] = shape.TextFrame.TextRange.Text
        except:
            d["text"] = ""
    return d

def slide_to_dict(slide):
    """Convert COM Slide to summary dict."""
    return {
        "index": slide.SlideIndex,
        "slide_id": slide.SlideID,
        "layout": slide.Layout,
        "layout_name": _safe_attr(slide.CustomLayout, "Name", ""),
        "shapes_count": slide.Shapes.Count,
        "name": slide.Name,
    }
```

### Utility Helpers

```python
def _shape_type_name(type_int):
    """Map msoShapeType enum to human name."""
    mapping = {
        1: "AutoShape", 2: "Callout", 3: "Chart", 4: "Comment",
        5: "FreeformShape", 6: "Group", 7: "EmbeddedOLEObject",
        9: "Line", 13: "Picture", 14: "Placeholder",
        15: "TextEffect", 16: "MediaObject", 17: "TextBox",
        19: "Table", 24: "SmartArt", 27: "Graphic",
    }
    return mapping.get(type_int, f"Unknown({type_int})")

def _safe_attr(obj, attr, default=""):
    """Safely read a COM attribute."""
    try:
        return getattr(obj, attr, default) or default
    except:
        return default

def _inches_to_points(inches):
    return round(inches * 72, 2)

def _points_to_inches(pts):
    return round(pts / 72, 4)

def _parse_color(color_str):
    """Parse '#RRGGBB' or 'R,G,B' to COM BGR-packed integer."""
    if color_str.startswith("#"):
        r = int(color_str[1:3], 16)
        g = int(color_str[3:5], 16)
        b = int(color_str[5:7], 16)
    else:
        parts = [int(x.strip()) for x in color_str.split(",")]
        r, g, b = parts[0], parts[1], parts[2]
    return r + (g << 8) + (b << 16)  # COM uses BGR packed int
```

### Constant Maps

```python
LAYOUT_MAP = {
    "blank": 12, "title": 1, "title_content": 2,
    "section_header": 3, "two_content": 4, "comparison": 5,
    "title_only": 11, "content_caption": 7, "picture_caption": 8,
}

SAVE_FORMAT_MAP = {
    "pptx": 24, "ppt": 1, "pdf": 32, "png": 18, "jpg": 17,
    "gif": 16, "bmp": 19, "tif": 21, "odp": 35, "potx": 26,
}

AUTOSHAPE_MAP = {
    "rectangle": 1, "rounded_rectangle": 5, "oval": 9,
    "diamond": 4, "triangle": 7, "right_arrow": 33,
    "left_arrow": 34, "up_arrow": 35, "down_arrow": 36,
    "star_5": 92, "star_4": 91, "heart": 21,
    "lightning": 22, "callout_1": 41, "callout_2": 42,
    "cloud": 179,
}

TRANSITION_MAP = {
    "none": 0, "fade": 3844, "push": 3845, "wipe": 3846,
    "split": 3847, "reveal": 3848, "cut": 257,
    "dissolve": 1537, "checkerboard": 1025,
}

ANIMATION_MAP = {
    "appear": 1, "fade": 10, "fly_in": 2, "wipe": 22,
    "split": 13, "wheel": 21, "grow_shrink": 50,
}
```

---

## Phase 1: Core — Application & Presentation Management (14 tools)

| # | Tool Name | Signature | Description |
|---|-----------|-----------|-------------|
| 1 | `launch_powerpoint` | `() -> str` | Launch PowerPoint or connect to running instance |
| 2 | `get_app_info` | `() -> str` | Version, window state, presentations count |
| 3 | `new_presentation` | `(template_path: str = "") -> str` | Create blank or from template |
| 4 | `open_presentation` | `(file_path: str, read_only: bool = False) -> str` | Open .pptx/.ppt file |
| 5 | `save_presentation` | `() -> str` | Save active presentation |
| 6 | `save_presentation_as` | `(file_path: str, format: str = "pptx") -> str` | Save as pptx/ppt/pdf/potx/odp |
| 7 | `close_presentation` | `(save: bool = True) -> str` | Close active presentation |
| 8 | `list_presentations` | `() -> str` | List all open presentations |
| 9 | `switch_presentation` | `(name_or_index: str) -> str` | Activate a different open presentation |
| 10 | `get_presentation_info` | `() -> str` | Slide count, dimensions, metadata, file path |
| 11 | `set_presentation_properties` | `(properties_json: str) -> str` | Title, author, subject via BuiltInDocumentProperties |
| 12 | `export_presentation` | `(output_path: str, format: str = "pdf") -> str` | Export to PDF/images |
| 13 | `set_slide_size` | `(width_inches: float = 13.333, height_inches: float = 7.5, orientation: str = "landscape") -> str` | Set slide dimensions |
| 14 | `get_slide_masters` | `() -> str` | List slide masters and their layouts |

---

## Phase 2: Slide Operations (16 tools)

| # | Tool Name | Signature | Description |
|---|-----------|-----------|-------------|
| 15 | `get_slides` | `() -> str` | List all slides with summary info |
| 16 | `get_slide_info` | `(slide_index: int) -> str` | Detailed info for one slide |
| 17 | `add_slide` | `(layout: str = "blank", index: int = 0) -> str` | Add slide (0 = append) |
| 18 | `duplicate_slide` | `(slide_index: int) -> str` | Duplicate a slide |
| 19 | `delete_slide` | `(slide_index: int) -> str` | Delete a slide |
| 20 | `move_slide` | `(slide_index: int, new_index: int) -> str` | Move slide to new position |
| 21 | `copy_slide` | `(source_index: int, target_pres_name: str = "", target_index: int = 0) -> str` | Copy within/between presentations |
| 22 | `get_slide_notes` | `(slide_index: int) -> str` | Get speaker notes text |
| 23 | `set_slide_notes` | `(slide_index: int, notes: str) -> str` | Set speaker notes |
| 24 | `set_slide_layout` | `(slide_index: int, layout: str) -> str` | Change slide layout |
| 25 | `set_slide_transition` | `(slide_index: int, effect: str = "fade", duration: float = 1.0, advance_time: float = 0) -> str` | Set transition effect |
| 26 | `bulk_add_slides` | `(slides_json: str) -> str` | Add multiple slides at once |
| 27 | `reorder_slides` | `(order_json: str) -> str` | Reorder all slides by providing new index order |
| 28 | `get_slide_layout_names` | `() -> str` | List available layout names for active master |
| 29 | `set_slide_background` | `(slide_index: int, color: str = "", image_path: str = "") -> str` | Solid color or image background |
| 30 | `bulk_set_transitions` | `(settings_json: str) -> str` | Apply transitions to multiple slides |

---

## Phase 3: Content — Text & Shapes (18 tools)

| # | Tool Name | Signature | Description |
|---|-----------|-----------|-------------|
| 31 | `get_shapes` | `(slide_index: int) -> str` | List all shapes on a slide |
| 32 | `get_shape_details` | `(slide_index: int, shape_name: str) -> str` | Full details of one shape |
| 33 | `add_textbox` | `(slide_index: int, text: str, left: float, top: float, width: float, height: float, ...)` | Add formatted text box |
| 34 | `modify_text` | `(slide_index: int, shape_name: str, text: str, ...)` | Change text and formatting |
| 35 | `add_shape` | `(slide_index: int, shape_type: str, left: float, top: float, width: float, height: float, ...)` | Add auto-shape |
| 36 | `add_line` | `(slide_index: int, begin_x: float, begin_y: float, end_x: float, end_y: float, ...)` | Add line/arrow |
| 37 | `modify_shape` | `(slide_index: int, shape_name: str, ...)` | Update shape properties |
| 38 | `delete_shape` | `(slide_index: int, shape_name: str) -> str` | Remove a shape |
| 39 | `group_shapes` | `(slide_index: int, shape_names_json: str, group_name: str = "") -> str` | Group shapes together |
| 40 | `ungroup_shapes` | `(slide_index: int, group_name: str) -> str` | Ungroup a group shape |
| 41 | `add_hyperlink` | `(slide_index: int, shape_name: str, url: str, tooltip: str = "") -> str` | Add hyperlink to shape |
| 42 | `add_connector` | `(slide_index: int, shape1_name: str, shape2_name: str, connector_type: str = "straight") -> str` | Connect two shapes |
| 43 | `align_shapes` | `(slide_index: int, shape_names_json: str, alignment: str = "center") -> str` | Align multiple shapes |
| 44 | `distribute_shapes` | `(slide_index: int, shape_names_json: str, direction: str = "horizontal") -> str` | Evenly distribute shapes |
| 45 | `duplicate_shape` | `(slide_index: int, shape_name: str, offset_x: float = 0.5, offset_y: float = 0.5) -> str` | Duplicate with offset |
| 46 | `set_shape_z_order` | `(slide_index: int, shape_name: str, action: str = "front") -> str` | Bring to front/send to back |
| 47 | `bulk_add_shapes` | `(slide_index: int, shapes_json: str) -> str` | Add multiple shapes at once |
| 48 | `set_placeholder_text` | `(slide_index: int, placeholder_index: int, text: str, ...)` | Set text in layout placeholder |

---

## Phase 4: Content — Rich Media (14 tools)

| # | Tool Name | Signature | Description |
|---|-----------|-----------|-------------|
| 49 | `insert_image` | `(slide_index: int, image_path: str, left: float, top: float, ...)` | Insert picture from file |
| 50 | `insert_image_from_url` | `(slide_index: int, url: str, left: float, top: float, ...)` | Download and insert image |
| 51 | `add_table` | `(slide_index: int, rows: int, cols: int, left: float, top: float, width: float, height: float, ...)` | Insert table with optional data |
| 52 | `modify_table_cell` | `(slide_index: int, shape_name: str, row: int, col: int, text: str, ...)` | Update a single cell |
| 53 | `bulk_fill_table` | `(slide_index: int, shape_name: str, data_json: str) -> str` | Fill table from 2D array |
| 54 | `format_table` | `(slide_index: int, shape_name: str, ...)` | Style an entire table |
| 55 | `add_chart` | `(slide_index: int, chart_type: str, data_json: str, left: float, top: float, width: float, height: float, ...)` | Insert chart with data |
| 56 | `modify_chart` | `(slide_index: int, shape_name: str, ...)` | Update chart properties |
| 57 | `update_chart_data` | `(slide_index: int, shape_name: str, data_json: str) -> str` | Replace chart data |
| 58 | `insert_video` | `(slide_index: int, video_path: str, left: float, top: float, width: float, height: float) -> str` | Embed video |
| 59 | `insert_audio` | `(slide_index: int, audio_path: str, left: float = 0, top: float = 0) -> str` | Embed audio |
| 60 | `insert_ole_object` | `(slide_index: int, file_path: str, left: float, top: float, width: float, height: float, as_icon: bool = False) -> str` | Embed Excel/Word/etc. |
| 61 | `crop_image` | `(slide_index: int, shape_name: str, crop_left: float = 0, ...)` | Crop an embedded picture |
| 62 | `replace_image` | `(slide_index: int, shape_name: str, new_image_path: str) -> str` | Replace picture keeping position/size |

---

## Phase 5: Design & Themes (12 tools)

| # | Tool Name | Signature | Description |
|---|-----------|-----------|-------------|
| 63 | `get_theme_info` | `() -> str` | Current theme name, colors, fonts |
| 64 | `apply_theme` | `(theme_path: str) -> str` | Apply .thmx theme file |
| 65 | `get_theme_colors` | `() -> str` | Get all theme color slots |
| 66 | `set_theme_color` | `(slot: str, color: str) -> str` | Override a theme color slot |
| 67 | `get_theme_fonts` | `() -> str` | Major/minor theme fonts |
| 68 | `set_theme_fonts` | `(major_font: str = "", minor_font: str = "") -> str` | Change theme fonts |
| 69 | `get_master_layouts` | `(master_index: int = 1) -> str` | List layouts in a slide master |
| 70 | `modify_master_placeholder` | `(master_index: int, layout_index: int, placeholder_index: int, ...)` | Edit master placeholder formatting |
| 71 | `set_background` | `(slide_index: int = 0, color: str = "", ...)` | Background (0 = all slides) |
| 72 | `get_placeholders` | `(slide_index: int) -> str` | List placeholders with types and indices |
| 73 | `add_custom_layout` | `(master_index: int = 1, name: str = "Custom Layout") -> str` | Create new custom layout |
| 74 | `copy_master_from` | `(source_path: str) -> str` | Import slide master from another presentation |

---

## Phase 6: Advanced Operations (18 tools)

| # | Tool Name | Signature | Description |
|---|-----------|-----------|-------------|
| 75 | `find_and_replace` | `(find_text: str, replace_text: str, match_case: bool = False) -> str` | Find/replace across all slides |
| 76 | `extract_all_text` | `(include_notes: bool = True) -> str` | Extract all text from presentation |
| 77 | `get_presentation_outline` | `() -> str` | Hierarchical structure: slides → shapes → text |
| 78 | `merge_presentations` | `(file_paths_json: str, insert_at: int = 0) -> str` | Merge multiple files into active pres |
| 79 | `apply_template` | `(template_path: str) -> str` | Apply design template to active pres |
| 80 | `bulk_format_text` | `(criteria_json: str) -> str` | Apply formatting across matching text spans |
| 81 | `add_animation` | `(slide_index: int, shape_name: str, effect: str = "appear", ...)` | Add entrance animation |
| 82 | `remove_animation` | `(slide_index: int, shape_name: str) -> str` | Remove animations from shape |
| 83 | `get_animations` | `(slide_index: int) -> str` | List animation sequence for a slide |
| 84 | `reorder_animations` | `(slide_index: int, order_json: str) -> str` | Change animation play order |
| 85 | `bulk_speaker_notes` | `(notes_json: str) -> str` | Set notes on multiple slides at once |
| 86 | `clone_formatting` | `(slide_index: int, source_shape: str, target_shapes_json: str) -> str` | Copy formatting from one shape to others |
| 87 | `search_shapes` | `(query: str, search_text: bool = True, search_names: bool = True) -> str` | Find shapes by text or name |
| 88 | `rename_shape` | `(slide_index: int, old_name: str, new_name: str) -> str` | Rename a shape |
| 89 | `lock_shape` | `(slide_index: int, shape_name: str, lock: bool = True) -> str` | Lock/unlock shape |
| 90 | `add_section` | `(name: str, before_slide: int = 0) -> str` | Add a named section |
| 91 | `get_sections` | `() -> str` | List all sections |
| 92 | `delete_section` | `(section_index: int, delete_slides: bool = False) -> str` | Remove a section |

---

## Phase 7: Analysis & Export (13 tools)

| # | Tool Name | Signature | Description |
|---|-----------|-----------|-------------|
| 93 | `get_presentation_stats` | `() -> str` | Word count, shape count, image count, slide count |
| 94 | `export_slide_image` | `(slide_index: int, output_path: str, format: str = "png", width: int = 1920) -> str` | Export one slide as image |
| 95 | `export_all_slides_images` | `(output_dir: str, format: str = "png", width: int = 1920) -> str` | Export all slides as images |
| 96 | `export_pdf` | `(output_path: str, slides_range: str = "", quality: str = "high") -> str` | Export to PDF with options |
| 97 | `get_fonts_used` | `() -> str` | List all fonts used in presentation |
| 98 | `get_linked_files` | `() -> str` | List linked/embedded media files |
| 99 | `check_accessibility` | `() -> str` | Alt-text missing, reading order, contrast issues |
| 100 | `get_slide_thumbnails_base64` | `(slide_indices_json: str = "", width: int = 320) -> str` | Base64 thumbnails for AI review |
| 101 | `compare_slides` | `(slide_a: int, slide_b: int) -> str` | Compare two slides |
| 102 | `snapshot_to_json` | `() -> str` | Full presentation snapshot as JSON |
| 103 | `get_color_usage` | `() -> str` | Audit all colors used across presentation |
| 104 | `validate_presentation` | `() -> str` | Check for common issues |
| 105 | `get_text_by_slide` | `() -> str` | Return text organized by slide |

---

## Tool Count Summary

| Phase | Category | Tools |
|-------|----------|-------|
| 1 | Application & Presentation Management | 14 |
| 2 | Slide Operations | 16 |
| 3 | Text & Shapes | 18 |
| 4 | Rich Media (Images, Tables, Charts) | 14 |
| 5 | Design & Themes | 12 |
| 6 | Advanced Operations | 18 |
| 7 | Analysis & Export | 13 |
| **Total** | | **105** |

---

## Implementation Order

```
Phase 1 (14 tools)
    └─→ Phase 2 (16 tools)
            ├─→ Phase 3 (18 tools) ──┐
            ├─→ Phase 4 (14 tools) ──┤
            └─→ Phase 5 (12 tools) ──┤
                                      └─→ Phase 6 (18 tools)
                                              └─→ Phase 7 (13 tools)
```

Each phase: implement tools → write tests → verify with PowerPoint running → move to next phase.

---

## Key COM Patterns

### MsoTriState
```python
# PowerPoint COM uses -1 (True) and 0 (False), NOT Python bools
shape.TextFrame.TextRange.Font.Bold = -1   # msoTrue
shape.TextFrame.TextRange.Font.Bold = 0    # msoFalse
```

### Colors (BGR, not RGB)
```python
# COM uses BGR-packed integer
# _parse_color("#FF0000") → Red → 0x0000FF = 255
```

### Chart Data (embedded Excel)
```python
chart = shape.Chart
chart_data = chart.ChartData
chart_data.Activate()
wb = chart_data.Workbook
ws = wb.Worksheets(1)
ws.Cells(1, 1).Value = "Category"
# ... populate cells ...
wb.Close(True)  # MUST close or leaks COM references
```

### Slide Creation
```python
layout = pres.SlideMaster.CustomLayouts(layoutIndex)
slide = pres.Slides.AddSlide(insertIndex, layout)
```

### Animations
```python
seq = slide.TimeLine.MainSequence
effect = seq.AddEffect(shape, effectId, trigger=1)  # msoAnimTriggerOnClick
effect.Timing.Duration = 0.5
```

### Export
```python
pres.SaveAs(output_path, 32)  # ppSaveAsPDF
pres.Slides(index).Export(output_path, "PNG", width, height)
```

---

## Error Handling Strategy

1. **COM connection**: `get_app()` tries `GetActiveObject` first, falls back to `Dispatch` + `app.Visible = True`
2. **Tool-level**: Each tool catches COM errors → returns `{"error": "descriptive message"}`
3. **Index validation**: `get_slide()` / `get_shape()` validate bounds before COM access
4. **MsoTriState guards**: Compare against `-1` and `0`, never Python `True`/`False`
5. **HasTextFrame / HasTable / HasChart**: Always check before accessing sub-objects
6. **Bulk operations**: Collect per-item errors, return `{"results": [...], "errors": [...]}`

---

## Testing Approach

Following the MS Project test pattern:

```python
"""Integration test for Phase N PowerPoint MCP tools."""
import asyncio, json, importlib.util, os

_server_path = os.path.join(os.path.dirname(__file__), "..", "server.py")
spec = importlib.util.spec_from_file_location("server", _server_path)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

async def call(name, args=None):
    r = await mod.mcp.call_tool(name, args or {})
    if isinstance(r, list):
        item = r[0]
        text = item.text if hasattr(item, "text") else str(item)
    elif hasattr(r, "text"):
        text = r.text
    else:
        text = str(r)
    try:
        return json.loads(text)
    except:
        return text

async def test():
    # Create test presentation, exercise tools, close without saving
    r = await call("new_presentation")
    assert r["status"] == "created"
    # ... test each tool ...
    await call("close_presentation", {"save": False})

if __name__ == "__main__":
    asyncio.run(test())
```

**Prerequisites**: PowerPoint must be installed and running on Windows.

---

## Verification Plan

1. **Per-phase tests**: `python tests/test_phaseN.py` with PowerPoint open
2. **Manual smoke test**: Register in `claude_desktop_config.json` and test via Claude Desktop
3. **End-to-end**: Create a full presentation programmatically (title slide, content slides, images, charts, theme, transitions, export to PDF)

### Claude Desktop Configuration

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["C:/Users/Ibrahim Elsahafy/mcp-servers/powerpoint/server.py"]
    }
  }
}
```

---

## Potential Challenges

| Challenge | Mitigation |
|-----------|------------|
| Chart data uses embedded Excel workbook | Always `wb.Close(True)` after modification to avoid leaked COM references |
| Custom layouts accessed by index, not name | Helper iterates `CustomLayouts` and matches by `.Name` property |
| PowerPoint COM uses BGR colors (not RGB) | `_parse_color()` handles conversion automatically |
| MsoTriState booleans (-1/0 not True/False) | All boolean sets use explicit `-1` / `0` constants |
| `AddPicture` requires absolute file paths | Resolve to absolute path with `os.path.abspath()` before passing |
| Master slide changes propagate to all slides | Tests use temporary presentations, close without saving |
| Animation sequence is 1-based and ordered | Careful index management in reorder operations |
