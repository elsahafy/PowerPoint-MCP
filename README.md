# PowerPoint MCP Server

Enterprise-grade Microsoft PowerPoint automation via the [Model Context Protocol (MCP)](https://modelcontextprotocol.io). Control PowerPoint presentations programmatically through **105 tools** using Windows COM automation.

## Features

| Phase | Category | Tools |
|-------|----------|------:|
| 1 | Application & Presentation Management | 14 |
| 2 | Slide Operations | 16 |
| 3 | Text & Shapes | 18 |
| 4 | Rich Media (Images, Tables, Charts, Video, Audio) | 14 |
| 5 | Design & Themes | 12 |
| 6 | Advanced Operations (Animations, Sections, Find/Replace) | 18 |
| 7 | Analysis & Export (Stats, PDF, Accessibility, Validation) | 13 |
| **Total** | | **105** |

## Requirements

- **Windows** 10/11
- **Microsoft PowerPoint** (desktop, locally installed)
- **Python** 3.10+
- **Dependencies**: `mcp`, `pywin32`

## Installation

```bash
git clone git@github.com:elsahafy/PowerPoint-MCP.git
cd PowerPoint-MCP
pip install mcp pywin32
```

## Usage

### Standalone

```bash
python server.py
```

### Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["C:/Users/YourName/PowerPoint-MCP/server.py"]
    }
  }
}
```

### Claude Code

Add to your MCP settings:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["C:/Users/YourName/PowerPoint-MCP/server.py"]
    }
  }
}
```

## Tool Reference

### Phase 1 — Application & Presentation Management

| Tool | Description |
|------|-------------|
| `launch_powerpoint` | Launch PowerPoint or connect to running instance |
| `get_app_info` | Version, window state, presentations count |
| `new_presentation` | Create blank or from template |
| `open_presentation` | Open .pptx/.ppt file |
| `save_presentation` | Save active presentation |
| `save_presentation_as` | Save as pptx/ppt/pdf/potx/odp |
| `close_presentation` | Close active presentation |
| `list_presentations` | List all open presentations |
| `switch_presentation` | Activate a different open presentation |
| `get_presentation_info` | Slide count, dimensions, metadata, file path |
| `set_presentation_properties` | Title, author, subject via BuiltInDocumentProperties |
| `export_presentation` | Export to PDF/images |
| `set_slide_size` | Set slide dimensions and orientation |
| `get_slide_masters` | List slide masters and their layouts |

### Phase 2 — Slide Operations

| Tool | Description |
|------|-------------|
| `get_slides` | List all slides with summary info |
| `get_slide_info` | Detailed info for one slide |
| `add_slide` | Add slide with specified layout |
| `duplicate_slide` | Duplicate a slide |
| `delete_slide` | Delete a slide |
| `move_slide` | Move slide to new position |
| `copy_slide` | Copy within/between presentations |
| `get_slide_notes` | Get speaker notes |
| `set_slide_notes` | Set speaker notes |
| `set_slide_layout` | Change slide layout |
| `set_slide_transition` | Set transition effect |
| `bulk_add_slides` | Add multiple slides at once |
| `reorder_slides` | Reorder all slides |
| `get_slide_layout_names` | List available layout names |
| `set_slide_background` | Solid color or image background |
| `bulk_set_transitions` | Apply transitions to multiple slides |

### Phase 3 — Text & Shapes

| Tool | Description |
|------|-------------|
| `get_shapes` | List all shapes on a slide |
| `get_shape_details` | Full details of one shape |
| `add_textbox` | Add formatted text box |
| `modify_text` | Change text and formatting |
| `add_shape` | Add auto-shape (rectangle, oval, arrow, star, etc.) |
| `add_line` | Add line/arrow |
| `modify_shape` | Update shape properties |
| `delete_shape` | Remove a shape |
| `group_shapes` | Group shapes together |
| `ungroup_shapes` | Ungroup a group shape |
| `add_hyperlink` | Add hyperlink to shape |
| `add_connector` | Connect two shapes |
| `align_shapes` | Align multiple shapes |
| `distribute_shapes` | Evenly distribute shapes |
| `duplicate_shape` | Duplicate with offset |
| `set_shape_z_order` | Bring to front/send to back |
| `bulk_add_shapes` | Add multiple shapes at once |
| `set_placeholder_text` | Set text in layout placeholder |

### Phase 4 — Rich Media

| Tool | Description |
|------|-------------|
| `insert_image` | Insert picture from file |
| `insert_image_from_url` | Download and insert image |
| `add_table` | Insert table with optional data |
| `modify_table_cell` | Update a single cell |
| `bulk_fill_table` | Fill table from 2D array |
| `format_table` | Style an entire table |
| `add_chart` | Insert chart (column, bar, line, pie, area, scatter, etc.) |
| `modify_chart` | Update chart properties |
| `update_chart_data` | Replace chart data |
| `insert_video` | Embed video |
| `insert_audio` | Embed audio |
| `insert_ole_object` | Embed Excel/Word/etc. |
| `crop_image` | Crop an embedded picture |
| `replace_image` | Replace picture keeping position/size |

### Phase 5 — Design & Themes

| Tool | Description |
|------|-------------|
| `get_theme_info` | Current theme name, colors, fonts |
| `apply_theme` | Apply .thmx theme file |
| `get_theme_colors` | Get all theme color slots |
| `set_theme_color` | Override a theme color slot |
| `get_theme_fonts` | Major/minor theme fonts |
| `set_theme_fonts` | Change theme fonts |
| `get_master_layouts` | List layouts in a slide master |
| `modify_master_placeholder` | Edit master placeholder formatting |
| `set_background` | Background for slide or all slides |
| `get_placeholders` | List placeholders with types and indices |
| `add_custom_layout` | Create new custom layout |
| `copy_master_from` | Import slide master from another presentation |

### Phase 6 — Advanced Operations

| Tool | Description |
|------|-------------|
| `find_and_replace` | Find/replace across all slides |
| `extract_all_text` | Extract all text from presentation |
| `get_presentation_outline` | Hierarchical structure: slides, shapes, text |
| `merge_presentations` | Merge multiple files into active presentation |
| `apply_template` | Apply design template |
| `bulk_format_text` | Apply formatting across matching text |
| `add_animation` | Add entrance animation |
| `remove_animation` | Remove animations from shape |
| `get_animations` | List animation sequence |
| `reorder_animations` | Change animation play order |
| `bulk_speaker_notes` | Set notes on multiple slides |
| `clone_formatting` | Copy formatting from one shape to others |
| `search_shapes` | Find shapes by text or name |
| `rename_shape` | Rename a shape |
| `lock_shape` | Lock/unlock shape |
| `add_section` | Add a named section |
| `get_sections` | List all sections |
| `delete_section` | Remove a section |

### Phase 7 — Analysis & Export

| Tool | Description |
|------|-------------|
| `get_presentation_stats` | Word count, shape count, image count, slide count |
| `export_slide_image` | Export one slide as image |
| `export_all_slides_images` | Export all slides as images |
| `export_pdf` | Export to PDF with options |
| `get_fonts_used` | List all fonts used |
| `get_linked_files` | List linked/embedded media |
| `check_accessibility` | Alt-text, reading order, contrast issues |
| `get_slide_thumbnails_base64` | Base64 thumbnails for AI review |
| `compare_slides` | Compare two slides |
| `snapshot_to_json` | Full presentation snapshot as JSON |
| `get_color_usage` | Audit all colors used |
| `validate_presentation` | Check for common issues (score 1-100) |
| `get_text_by_slide` | Return text organized by slide |

## Architecture

```
server.py (single file, ~4,800 lines)
├── Constants & Maps (layouts, save formats, 57 shapes, 21 transitions, 15 animations)
├── Error Taxonomy (PPTError, ValidationError, NotFoundError, BoundsError, COMError, ReadOnlyError)
├── Response Envelope (_ok, _ok_list, _err) & Validation Helpers
├── Utility Helpers (_parse_color, _inches_to_points, etc.)
├── COM Helpers (get_app, get_pres, get_slide, get_shape)
├── Data Converters (shape_to_dict, slide_to_dict)
├── Phase 1-7 Tools (105 @mcp.tool() functions)
└── Entry Point (mcp.run())
```

All communication with PowerPoint happens through Windows COM via `win32com.client`. The server launches or connects to a running PowerPoint instance automatically.

## Error Handling

All tools return structured error responses with machine-readable error codes:

| Code | Class | Description |
|------|-------|-------------|
| `VALIDATION_ERROR` | `ValidationError` | Invalid input (bad color, zero dimensions, wrong JSON type) |
| `NOT_FOUND` | `NotFoundError` | Missing file, shape, or presentation |
| `OUT_OF_BOUNDS` | `BoundsError` | Slide/shape/table cell index out of range |
| `COM_ERROR` | `COMError` | PowerPoint COM/HRESULT failure |
| `READ_ONLY` | `ReadOnlyError` | Mutation attempted on read-only presentation |

Error response format:
```json
{"error": "Slide index 99 out of range (1..5).", "code": "OUT_OF_BOUNDS"}
```

All mutating tools include read-only guards. Input validation catches issues before COM calls (dimensions > 0, valid colors, file existence, JSON schema).

## Key Design Decisions

- **Single file** — No modules to manage, easy to deploy and debug.
- **Inches API, points internally** — Tool parameters use inches (human-friendly), converted to points for COM.
- **BGR color handling** — `_parse_color()` accepts `#RRGGBB` or `R,G,B` and converts to COM's BGR format.
- **MsoTriState** — COM booleans use `-1` (True) and `0` (False), never Python bools.
- **JSON in, JSON out** — All tools return `json.dumps(...)`. Bulk operations accept JSON string parameters.
- **Structured errors** — All errors include a `code` field for programmatic handling. COM HRESULT codes are translated to human-readable messages.
- **Resource safety** — Chart workbook handles use `try/finally` to prevent leaks. `reorder_slides` tracks by SlideID to avoid index drift.

## Testing

9 test suites covering all 105 tools:

```bash
# Run all phase tests (requires PowerPoint running on Windows)
python tests/test_phase1.py   # 14 tools — App & presentation management
python tests/test_phase2.py   # 16 tools — Slide operations
python tests/test_phase3.py   # 18 tools — Text & shapes
python tests/test_phase4.py   # 14 tools — Rich media
python tests/test_phase5.py   # 12 tools — Design & themes
python tests/test_phase6.py   # 18 tools — Advanced operations
python tests/test_phase7.py   # 13 tools — Analysis & export

# Negative tests (error codes, validation, bounds checking)
python tests/test_negative.py  # 22 tests

# End-to-end workflows (build pipeline, merge, reorder, unicode, etc.)
python tests/test_e2e.py       # 8 workflows
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## License

[MIT](LICENSE)
