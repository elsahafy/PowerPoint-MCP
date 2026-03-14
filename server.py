"""
PowerPoint MCP Server
Controls local Microsoft PowerPoint via COM automation.
Install: pip install mcp pywin32
Run:     python server.py
"""
import json
import os
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("PowerPoint")


# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANT MAPS
# ═══════════════════════════════════════════════════════════════════════════════

LAYOUT_MAP = {
    "blank": 12,
    "title": 1,
    "title_content": 2,
    "section_header": 3,
    "two_content": 4,
    "comparison": 5,
    "title_only": 11,
    "content_caption": 7,
    "picture_caption": 8,
}

SAVE_FORMAT_MAP = {
    "pptx": 24,
    "ppt": 1,
    "pdf": 32,
    "png": 18,
    "jpg": 17,
    "gif": 16,
    "bmp": 19,
    "tif": 21,
    "odp": 35,
    "potx": 26,
}

AUTOSHAPE_MAP = {
    "rectangle": 1,
    "rounded_rectangle": 5,
    "oval": 9,
    "diamond": 4,
    "triangle": 7,
    "right_arrow": 33,
    "left_arrow": 34,
    "up_arrow": 35,
    "down_arrow": 36,
    "star_5": 92,
    "star_4": 91,
    "heart": 21,
    "lightning": 22,
    "callout_1": 41,
    "callout_2": 42,
    "cloud": 179,
}

TRANSITION_MAP = {
    "none": 0,
    "fade": 3844,
    "push": 3845,
    "wipe": 3846,
    "split": 3847,
    "reveal": 3848,
    "cut": 257,
    "dissolve": 1537,
    "checkerboard": 1025,
}

ANIMATION_MAP = {
    "appear": 1,
    "fade": 10,
    "fly_in": 2,
    "wipe": 22,
    "split": 13,
    "wheel": 21,
    "grow_shrink": 50,
}


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITY HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

_SHAPE_TYPE_NAMES = {
    -2: "Placeholder",
    1: "AutoShape",
    2: "Callout",
    3: "Chart",
    4: "Comment",
    5: "Freeform",
    6: "Group",
    7: "EmbeddedOLEObject",
    8: "FormControl",
    9: "Line",
    10: "LinkedOLEObject",
    11: "LinkedPicture",
    12: "OLEControlObject",
    13: "Picture",
    14: "Placeholder",
    15: "MediaObject",
    16: "TextEffect",
    17: "Table",
    18: "Canvas",
    19: "Diagram",
    20: "Ink",
    21: "InkComment",
    22: "SmartArt",
    23: "Slicer",
    24: "WebVideo",
    25: "ContentApp",
    26: "Graphic",
    27: "LinkedGraphic",
    28: "3DModel",
    29: "Linked3DModel",
}


def _shape_type_name(type_int: int) -> str:
    """Map msoShapeType enum integer to a human-readable name."""
    return _SHAPE_TYPE_NAMES.get(type_int, f"Unknown({type_int})")


def _safe_attr(obj, attr: str, default=""):
    """Safely read a COM attribute, returning *default* on any error."""
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def _inches_to_points(inches: float) -> float:
    """Convert inches to PowerPoint points (1 inch = 72 points)."""
    return inches * 72.0


def _points_to_inches(pts: float) -> float:
    """Convert PowerPoint points to inches."""
    return pts / 72.0


def _parse_color(color_str: str) -> int:
    """Parse '#RRGGBB' or 'R,G,B' to a COM BGR-packed integer."""
    color_str = color_str.strip()
    if color_str.startswith("#"):
        hex_str = color_str.lstrip("#")
        if len(hex_str) != 6:
            raise ValueError(f"Invalid hex color: {color_str}")
        r = int(hex_str[0:2], 16)
        g = int(hex_str[2:4], 16)
        b = int(hex_str[4:6], 16)
    elif "," in color_str:
        parts = [int(p.strip()) for p in color_str.split(",")]
        if len(parts) != 3:
            raise ValueError(f"Invalid RGB color: {color_str}")
        r, g, b = parts
    else:
        raise ValueError(f"Unrecognised color format: {color_str}. Use '#RRGGBB' or 'R,G,B'.")
    return r + (g << 8) + (b << 16)


# ═══════════════════════════════════════════════════════════════════════════════
# CORE COM HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def get_app(require_presentation: bool = True):
    """
    Get the running PowerPoint.Application COM instance.
    Launches PowerPoint if it is not already running.
    If *require_presentation* is True, raises when no presentation is open.
    """
    import win32com.client

    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        app = win32com.client.Dispatch("PowerPoint.Application")

    app.Visible = True  # COM constant msoTrue is -1, but True works for Visible

    if require_presentation and app.Presentations.Count == 0:
        raise RuntimeError("No presentation is open. Create or open one first.")

    return app


def get_pres(app):
    """Return the active presentation, or raise if none."""
    if app.Presentations.Count == 0:
        raise RuntimeError("No presentation is open.")
    return app.ActivePresentation


def get_slide(pres, slide_index: int):
    """Return a slide by 1-based index with bounds checking."""
    count = pres.Slides.Count
    if slide_index < 1 or slide_index > count:
        raise IndexError(f"Slide index {slide_index} out of range (1..{count}).")
    return pres.Slides(slide_index)


def get_shape(slide, shape_id):
    """
    Find a shape on *slide* by name (str) or 1-based index (int/str-of-int).
    """
    # Try numeric index first
    try:
        idx = int(shape_id)
        count = slide.Shapes.Count
        if idx < 1 or idx > count:
            raise IndexError(f"Shape index {idx} out of range (1..{count}).")
        return slide.Shapes(idx)
    except (ValueError, TypeError):
        pass

    # Fall back to name search
    name = str(shape_id)
    for i in range(1, slide.Shapes.Count + 1):
        if slide.Shapes(i).Name == name:
            return slide.Shapes(i)

    raise KeyError(f"Shape '{shape_id}' not found on slide.")


# ═══════════════════════════════════════════════════════════════════════════════
# DATA CONVERSION HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def shape_to_dict(shape) -> dict:
    """Convert a COM Shape object to a plain dictionary."""
    has_text = False
    text = ""
    try:
        if shape.HasTextFrame:
            has_text = True
            text = shape.TextFrame.TextRange.Text
    except Exception:
        pass

    type_int = int(_safe_attr(shape, "Type", 0))

    return {
        "name": _safe_attr(shape, "Name"),
        "shape_id": _safe_attr(shape, "Id"),
        "type": type_int,
        "type_name": _shape_type_name(type_int),
        "left": round(_points_to_inches(float(_safe_attr(shape, "Left", 0))), 4),
        "top": round(_points_to_inches(float(_safe_attr(shape, "Top", 0))), 4),
        "width": round(_points_to_inches(float(_safe_attr(shape, "Width", 0))), 4),
        "height": round(_points_to_inches(float(_safe_attr(shape, "Height", 0))), 4),
        "rotation": float(_safe_attr(shape, "Rotation", 0)),
        "has_text": has_text,
        "text": text,
    }


def slide_to_dict(slide) -> dict:
    """Convert a COM Slide object to a summary dictionary."""
    layout_name = ""
    layout_index = 0
    try:
        layout_name = slide.CustomLayout.Name
    except Exception:
        pass
    try:
        layout_index = slide.Layout
    except Exception:
        pass

    return {
        "index": slide.SlideIndex,
        "slide_id": slide.SlideID,
        "layout": layout_index,
        "layout_name": layout_name,
        "shapes_count": slide.Shapes.Count,
        "name": _safe_attr(slide, "Name", ""),
    }


# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 1 TOOLS
# ═══════════════════════════════════════════════════════════════════════════════

# ---------------------------------------------------------------------------
# Tool 1: launch_powerpoint
# ---------------------------------------------------------------------------

@mcp.tool()
def launch_powerpoint() -> str:
    """Launch PowerPoint or connect to a running instance."""
    try:
        app = get_app(require_presentation=False)
        return json.dumps({
            "status": "connected",
            "version": str(_safe_attr(app, "Version")),
            "visible": bool(app.Visible),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 2: get_app_info
# ---------------------------------------------------------------------------

@mcp.tool()
def get_app_info() -> str:
    """Get PowerPoint application info: version, window state, presentations count, and active presentation name."""
    try:
        app = get_app(require_presentation=False)
        active_name = ""
        try:
            active_name = app.ActivePresentation.Name
        except Exception:
            pass

        return json.dumps({
            "version": str(_safe_attr(app, "Version")),
            "window_state": int(_safe_attr(app, "WindowState", 0)),
            "presentations_count": app.Presentations.Count,
            "active_presentation": active_name,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 3: new_presentation
# ---------------------------------------------------------------------------

@mcp.tool()
def new_presentation(template_path: str = "") -> str:
    """Create a new blank presentation or one based on a template (.potx / .pptx)."""
    try:
        app = get_app(require_presentation=False)
        if template_path:
            abs_path = os.path.abspath(template_path)
            if not os.path.isfile(abs_path):
                return json.dumps({"error": f"Template not found: {abs_path}"}, indent=2)
            pres = app.Presentations.Open(abs_path)
        else:
            pres = app.Presentations.Add()

        # Ensure at least one slide exists (some configs create empty pres)
        if pres.Slides.Count == 0:
            pres.Slides.Add(1, LAYOUT_MAP["blank"])

        return json.dumps({
            "status": "created",
            "name": pres.Name,
            "slide_count": pres.Slides.Count,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 4: open_presentation
# ---------------------------------------------------------------------------

@mcp.tool()
def open_presentation(file_path: str, read_only: bool = False) -> str:
    """Open an existing .pptx / .ppt file."""
    try:
        app = get_app(require_presentation=False)
        abs_path = os.path.abspath(file_path)
        if not os.path.isfile(abs_path):
            return json.dumps({"error": f"File not found: {abs_path}"}, indent=2)

        pres = app.Presentations.Open(abs_path, ReadOnly=read_only)
        return json.dumps({
            "status": "opened",
            "name": pres.Name,
            "path": pres.FullName,
            "slide_count": pres.Slides.Count,
            "read_only": bool(pres.ReadOnly),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 5: save_presentation
# ---------------------------------------------------------------------------

@mcp.tool()
def save_presentation() -> str:
    """Save the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        pres.Save()
        return json.dumps({
            "status": "saved",
            "name": pres.Name,
            "path": pres.FullName,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 6: save_presentation_as
# ---------------------------------------------------------------------------

@mcp.tool()
def save_presentation_as(file_path: str, format: str = "pptx") -> str:
    """Save the active presentation to a new path with an optional format (pptx, pdf, png, jpg, etc.)."""
    try:
        app = get_app()
        pres = get_pres(app)
        abs_path = os.path.abspath(file_path)
        fmt_lower = format.lower()
        format_id = SAVE_FORMAT_MAP.get(fmt_lower)
        if format_id is None:
            return json.dumps({
                "error": f"Unknown format '{format}'. Supported: {', '.join(SAVE_FORMAT_MAP.keys())}"
            }, indent=2)

        pres.SaveAs(abs_path, format_id)
        return json.dumps({
            "status": "saved_as",
            "path": abs_path,
            "format": fmt_lower,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 7: close_presentation
# ---------------------------------------------------------------------------

@mcp.tool()
def close_presentation(save: bool = True) -> str:
    """Close the active presentation, optionally saving first."""
    try:
        app = get_app()
        pres = get_pres(app)
        name = pres.Name
        if save:
            try:
                pres.Save()
            except Exception:
                pass  # May fail if never saved to disk; that's OK
        pres.Close()
        return json.dumps({
            "status": "closed",
            "name": name,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 8: list_presentations
# ---------------------------------------------------------------------------

@mcp.tool()
def list_presentations() -> str:
    """List all open presentations."""
    try:
        app = get_app(require_presentation=False)
        results = []
        for i in range(1, app.Presentations.Count + 1):
            p = app.Presentations(i)
            results.append({
                "name": p.Name,
                "path": _safe_attr(p, "FullName", ""),
                "slide_count": p.Slides.Count,
                "read_only": bool(p.ReadOnly),
            })
        return json.dumps(results, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 9: switch_presentation
# ---------------------------------------------------------------------------

@mcp.tool()
def switch_presentation(name_or_index: str) -> str:
    """Activate a different open presentation by name or 1-based index."""
    try:
        app = get_app(require_presentation=False)
        pres = None

        # Try numeric index first
        try:
            idx = int(name_or_index)
            if 1 <= idx <= app.Presentations.Count:
                pres = app.Presentations(idx)
        except (ValueError, TypeError):
            pass

        # Fall back to name search
        if pres is None:
            for i in range(1, app.Presentations.Count + 1):
                if app.Presentations(i).Name.lower() == name_or_index.lower():
                    pres = app.Presentations(i)
                    break

        if pres is None:
            return json.dumps({"error": f"Presentation '{name_or_index}' not found."}, indent=2)

        # Activate by bringing its first window to the front
        pres.Windows(1).Activate()
        return json.dumps({
            "status": "switched",
            "name": pres.Name,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 10: get_presentation_info
# ---------------------------------------------------------------------------

@mcp.tool()
def get_presentation_info() -> str:
    """Get detailed info about the active presentation: slides, dimensions, metadata, and file path."""
    try:
        app = get_app()
        pres = get_pres(app)

        width_in = round(_points_to_inches(pres.PageSetup.SlideWidth), 4)
        height_in = round(_points_to_inches(pres.PageSetup.SlideHeight), 4)

        metadata = {}
        for key in ("Title", "Author", "Subject", "Keywords", "Comments", "Category"):
            try:
                metadata[key] = pres.BuiltInDocumentProperties(key).Value
            except Exception:
                metadata[key] = ""

        return json.dumps({
            "name": pres.Name,
            "path": pres.FullName,
            "slide_count": pres.Slides.Count,
            "width_inches": width_in,
            "height_inches": height_in,
            "read_only": bool(pres.ReadOnly),
            "metadata": metadata,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 11: set_presentation_properties
# ---------------------------------------------------------------------------

@mcp.tool()
def set_presentation_properties(properties_json: str) -> str:
    """Set built-in document properties (Title, Author, Subject, etc.) from a JSON string."""
    try:
        app = get_app()
        pres = get_pres(app)
        props = json.loads(properties_json)
        updated = []
        for key, value in props.items():
            try:
                pres.BuiltInDocumentProperties(key).Value = str(value)
                updated.append(key)
            except Exception as prop_err:
                return json.dumps({"error": f"Failed to set '{key}': {prop_err}"}, indent=2)

        return json.dumps({
            "status": "updated",
            "properties_set": updated,
        }, indent=2)
    except json.JSONDecodeError as je:
        return json.dumps({"error": f"Invalid JSON: {je}"}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 12: export_presentation
# ---------------------------------------------------------------------------

@mcp.tool()
def export_presentation(output_path: str, format: str = "pdf") -> str:
    """Export the active presentation to PDF or image format."""
    try:
        app = get_app()
        pres = get_pres(app)

        if pres.Slides.Count == 0:
            return json.dumps({"error": "Cannot export — presentation has no slides."}, indent=2)

        abs_path = os.path.abspath(output_path)
        fmt_lower = format.lower()

        format_id = SAVE_FORMAT_MAP.get(fmt_lower)
        if format_id is None:
            return json.dumps({
                "error": f"Unknown format '{format}'. Supported: {', '.join(SAVE_FORMAT_MAP.keys())}"
            }, indent=2)

        pres.SaveAs(abs_path, format_id)
        return json.dumps({
            "status": "exported",
            "path": abs_path,
            "format": fmt_lower,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 13: set_slide_size
# ---------------------------------------------------------------------------

@mcp.tool()
def set_slide_size(
    width_inches: float = 13.333,
    height_inches: float = 7.5,
    orientation: str = "landscape",
) -> str:
    """Set slide dimensions in inches and orientation."""
    try:
        app = get_app()
        pres = get_pres(app)
        pres.PageSetup.SlideWidth = _inches_to_points(width_inches)
        pres.PageSetup.SlideHeight = _inches_to_points(height_inches)

        # ppSlideSizeCustom = 7 avoids the "fit content" prompt
        # Orientation: 1 = landscape (msoOrientationHorizontal), 2 = portrait (msoOrientationVertical)
        if orientation.lower() == "portrait":
            pres.PageSetup.SlideOrientation = 2
        else:
            pres.PageSetup.SlideOrientation = 1

        return json.dumps({
            "status": "updated",
            "width_inches": width_inches,
            "height_inches": height_inches,
            "orientation": orientation.lower(),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 14: get_slide_masters
# ---------------------------------------------------------------------------

@mcp.tool()
def get_slide_masters() -> str:
    """List all slide masters and their layouts in the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        masters = []

        for d in range(1, pres.Designs.Count + 1):
            design = pres.Designs(d)
            layouts = []
            try:
                for l in range(1, design.SlideMaster.CustomLayouts.Count + 1):
                    layout = design.SlideMaster.CustomLayouts(l)
                    layouts.append({
                        "index": l,
                        "name": _safe_attr(layout, "Name", ""),
                    })
            except Exception:
                pass

            masters.append({
                "design_index": d,
                "name": _safe_attr(design, "Name", ""),
                "layouts": layouts,
            })

        return json.dumps(masters, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 2 — Slide Operations (16 tools)
# ═══════════════════════════════════════════════════════════════════════════════


# ---------------------------------------------------------------------------
# Tool 15: get_slides
# ---------------------------------------------------------------------------

@mcp.tool()
def get_slides() -> str:
    """List all slides in the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        slides = []
        for i in range(1, pres.Slides.Count + 1):
            slides.append(slide_to_dict(pres.Slides(i)))
        return json.dumps(slides, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 16: get_slide_info
# ---------------------------------------------------------------------------

@mcp.tool()
def get_slide_info(slide_index: int) -> str:
    """Get detailed info for one slide: layout, shapes, notes, background, and transition."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        info = slide_to_dict(slide)

        # Shapes
        shapes = []
        for i in range(1, slide.Shapes.Count + 1):
            shapes.append(shape_to_dict(slide.Shapes(i)))
        info["shapes"] = shapes

        # Notes
        notes = ""
        try:
            notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        except Exception:
            pass
        info["notes"] = notes

        # Background
        bg_info = {}
        try:
            bg = slide.Background
            bg_info["follow_master"] = bool(slide.FollowMasterBackground)
            bg_info["fill_type"] = int(_safe_attr(bg.Fill, "Type", 0))
        except Exception:
            pass
        info["background"] = bg_info

        # Transition
        trans_info = {}
        try:
            t = slide.SlideShowTransition
            trans_info["entry_effect"] = int(_safe_attr(t, "EntryEffect", 0))
            trans_info["duration"] = float(_safe_attr(t, "Duration", 0))
            trans_info["advance_on_time"] = bool(_safe_attr(t, "AdvanceOnTime", False))
            trans_info["advance_time"] = float(_safe_attr(t, "AdvanceTime", 0))
        except Exception:
            pass
        info["transition"] = trans_info

        return json.dumps(info, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 17: add_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def add_slide(layout: str = "blank", index: int = 0) -> str:
    """Add a new slide with a specified layout. index=0 means append at the end."""
    try:
        app = get_app()
        pres = get_pres(app)

        layout_enum = LAYOUT_MAP.get(layout.lower())
        if layout_enum is None:
            return json.dumps({
                "error": f"Unknown layout '{layout}'. Supported: {', '.join(LAYOUT_MAP.keys())}"
            }, indent=2)

        insert_idx = index if index > 0 else pres.Slides.Count + 1
        new_slide = pres.Slides.Add(insert_idx, layout_enum)

        return json.dumps(slide_to_dict(new_slide), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 18: duplicate_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def duplicate_slide(slide_index: int) -> str:
    """Duplicate a slide. The copy is inserted immediately after the original."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        dup_range = slide.Duplicate()
        new_slide = dup_range(1)
        return json.dumps(slide_to_dict(new_slide), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 19: delete_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def delete_slide(slide_index: int) -> str:
    """Delete a slide by 1-based index."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        slide.Delete()
        return json.dumps({
            "status": "deleted",
            "slide_index": slide_index,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 20: move_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def move_slide(slide_index: int, new_index: int) -> str:
    """Move a slide from one position to another (1-based indices)."""
    try:
        app = get_app()
        pres = get_pres(app)
        get_slide(pres, slide_index)  # validate index
        pres.Slides(slide_index).MoveTo(new_index)
        return json.dumps({
            "status": "moved",
            "from": slide_index,
            "to": new_index,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 21: copy_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def copy_slide(source_index: int, target_pres_name: str = "", target_index: int = 0) -> str:
    """Copy a slide within the same presentation or to another open presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        get_slide(pres, source_index)  # validate

        if not target_pres_name:
            # Copy within same presentation
            dup_range = pres.Slides(source_index).Duplicate()
            new_slide = dup_range(1)
            if target_index > 0:
                new_slide.MoveTo(target_index)
            return json.dumps({
                "status": "copied",
                "target_presentation": pres.Name,
                "new_slide_index": new_slide.SlideIndex,
            }, indent=2)
        else:
            # Copy to another presentation
            target_pres = None
            for i in range(1, app.Presentations.Count + 1):
                if app.Presentations(i).Name.lower() == target_pres_name.lower():
                    target_pres = app.Presentations(i)
                    break
            if target_pres is None:
                return json.dumps({"error": f"Target presentation '{target_pres_name}' not found."}, indent=2)

            insert_at = target_index if target_index > 0 else target_pres.Slides.Count + 1
            target_pres.Slides.InsertFromFile(pres.FullName, insert_at - 1, source_index, source_index)
            return json.dumps({
                "status": "copied",
                "target_presentation": target_pres.Name,
                "new_slide_index": insert_at,
            }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 22: get_slide_notes
# ---------------------------------------------------------------------------

@mcp.tool()
def get_slide_notes(slide_index: int) -> str:
    """Get the speaker notes for a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        notes = ""
        try:
            notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        except Exception:
            pass
        return json.dumps({
            "slide_index": slide_index,
            "notes": notes,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 23: set_slide_notes
# ---------------------------------------------------------------------------

@mcp.tool()
def set_slide_notes(slide_index: int, notes: str) -> str:
    """Set the speaker notes for a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes
        return json.dumps({
            "status": "updated",
            "slide_index": slide_index,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 24: set_slide_layout
# ---------------------------------------------------------------------------

@mcp.tool()
def set_slide_layout(slide_index: int, layout: str) -> str:
    """Change the layout of an existing slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        layout_enum = LAYOUT_MAP.get(layout.lower())
        if layout_enum is None:
            return json.dumps({
                "error": f"Unknown layout '{layout}'. Supported: {', '.join(LAYOUT_MAP.keys())}"
            }, indent=2)

        slide.Layout = layout_enum
        return json.dumps({
            "status": "updated",
            "slide_index": slide_index,
            "layout": layout,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 25: set_slide_transition
# ---------------------------------------------------------------------------

@mcp.tool()
def set_slide_transition(
    slide_index: int,
    effect: str = "fade",
    duration: float = 1.0,
    advance_time: float = 0,
) -> str:
    """Set the transition effect for a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        trans = slide.SlideShowTransition
        trans.EntryEffect = TRANSITION_MAP.get(effect.lower(), 0)
        trans.Duration = duration

        if advance_time > 0:
            trans.AdvanceOnTime = -1  # msoTrue
            trans.AdvanceTime = advance_time

        return json.dumps({
            "status": "updated",
            "slide_index": slide_index,
            "effect": effect,
            "duration": duration,
            "advance_time": advance_time,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 26: bulk_add_slides
# ---------------------------------------------------------------------------

@mcp.tool()
def bulk_add_slides(slides_json: str) -> str:
    """Add multiple slides from a JSON array of {layout, index} objects."""
    try:
        app = get_app()
        pres = get_pres(app)
        slides_spec = json.loads(slides_json)
        results = []

        for spec in slides_spec:
            layout = spec.get("layout", "blank")
            index = spec.get("index", 0)

            layout_enum = LAYOUT_MAP.get(layout.lower())
            if layout_enum is None:
                results.append({"error": f"Unknown layout '{layout}'"})
                continue

            insert_idx = index if index > 0 else pres.Slides.Count + 1
            new_slide = pres.Slides.Add(insert_idx, layout_enum)
            results.append(slide_to_dict(new_slide))

        return json.dumps(results, indent=2)
    except json.JSONDecodeError as je:
        return json.dumps({"error": f"Invalid JSON: {je}"}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 27: reorder_slides
# ---------------------------------------------------------------------------

@mcp.tool()
def reorder_slides(order_json: str) -> str:
    """Reorder slides by providing a JSON array of slide indices in the desired order.
    Example: [3, 1, 2] means slide 3 goes to position 1, slide 1 to position 2, slide 2 to position 3."""
    try:
        app = get_app()
        pres = get_pres(app)
        new_order = json.loads(order_json)

        # Validate all indices
        count = pres.Slides.Count
        for idx in new_order:
            if idx < 1 or idx > count:
                return json.dumps({"error": f"Slide index {idx} out of range (1..{count})."}, indent=2)

        # Apply moves: place each slide at its target position
        for target_pos, slide_idx in enumerate(new_order, start=1):
            # Find current position of the slide with original SlideID
            pres.Slides(slide_idx).MoveTo(target_pos)

        return json.dumps({
            "status": "reordered",
            "new_order": new_order,
        }, indent=2)
    except json.JSONDecodeError as je:
        return json.dumps({"error": f"Invalid JSON: {je}"}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 28: get_slide_layout_names
# ---------------------------------------------------------------------------

@mcp.tool()
def get_slide_layout_names() -> str:
    """List all available custom layout names from the slide master."""
    try:
        app = get_app()
        pres = get_pres(app)
        layouts = []
        try:
            for i in range(1, pres.SlideMaster.CustomLayouts.Count + 1):
                layout = pres.SlideMaster.CustomLayouts(i)
                layouts.append({
                    "index": i,
                    "name": _safe_attr(layout, "Name", ""),
                })
        except Exception:
            pass
        return json.dumps(layouts, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 29: set_slide_background
# ---------------------------------------------------------------------------

@mcp.tool()
def set_slide_background(slide_index: int, color: str = "", image_path: str = "") -> str:
    """Set the background of a slide to a solid color or an image."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        slide.FollowMasterBackground = 0  # msoFalse

        if color:
            slide.Background.Fill.Solid()
            slide.Background.Fill.ForeColor.RGB = _parse_color(color)
        elif image_path:
            abs_path = os.path.abspath(image_path)
            if not os.path.isfile(abs_path):
                return json.dumps({"error": f"Image not found: {abs_path}"}, indent=2)
            slide.Background.Fill.UserPicture(abs_path)
        else:
            return json.dumps({"error": "Provide either 'color' or 'image_path'."}, indent=2)

        return json.dumps({
            "status": "updated",
            "slide_index": slide_index,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 30: bulk_set_transitions
# ---------------------------------------------------------------------------

@mcp.tool()
def bulk_set_transitions(settings_json: str) -> str:
    """Apply transitions to multiple slides from a JSON array of {slide_index, effect, duration, advance_time} objects."""
    try:
        app = get_app()
        pres = get_pres(app)
        settings = json.loads(settings_json)
        results = []

        for spec in settings:
            si = spec.get("slide_index")
            effect = spec.get("effect", "fade")
            duration = spec.get("duration", 1.0)
            advance_time = spec.get("advance_time", 0)

            try:
                slide = get_slide(pres, si)
                trans = slide.SlideShowTransition
                trans.EntryEffect = TRANSITION_MAP.get(effect.lower(), 0)
                trans.Duration = duration
                if advance_time > 0:
                    trans.AdvanceOnTime = -1  # msoTrue
                    trans.AdvanceTime = advance_time
                results.append({
                    "slide_index": si,
                    "status": "updated",
                    "effect": effect,
                })
            except Exception as slide_err:
                results.append({
                    "slide_index": si,
                    "error": str(slide_err),
                })

        return json.dumps(results, indent=2)
    except json.JSONDecodeError as je:
        return json.dumps({"error": f"Invalid JSON: {je}"}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 3 — Text & Shapes (18 tools)
# ═══════════════════════════════════════════════════════════════════════════════


# ---------------------------------------------------------------------------
# Tool 31: get_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def get_shapes(slide_index: int) -> str:
    """List all shapes on a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shapes = []
        for i in range(1, slide.Shapes.Count + 1):
            shapes.append(shape_to_dict(slide.Shapes(i)))
        return json.dumps(shapes, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 32: get_shape_details
# ---------------------------------------------------------------------------

@mcp.tool()
def get_shape_details(slide_index: int, shape_name: str) -> str:
    """Get full details of a shape: dimensions, fill, line, font info, and alt text."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        info = shape_to_dict(shape)

        # Fill color
        try:
            fill = shape.Fill
            if fill.Type == 1:  # msoFillSolid
                rgb = fill.ForeColor.RGB
                r = rgb & 0xFF
                g = (rgb >> 8) & 0xFF
                b = (rgb >> 16) & 0xFF
                info["fill_color"] = f"#{r:02X}{g:02X}{b:02X}"
            else:
                info["fill_color"] = ""
        except Exception:
            info["fill_color"] = ""

        # Line color and weight
        try:
            line = shape.Line
            if line.Visible:
                rgb = line.ForeColor.RGB
                r = rgb & 0xFF
                g = (rgb >> 8) & 0xFF
                b = (rgb >> 16) & 0xFF
                info["line_color"] = f"#{r:02X}{g:02X}{b:02X}"
                info["line_weight"] = float(line.Weight)
            else:
                info["line_color"] = ""
                info["line_weight"] = 0
        except Exception:
            info["line_color"] = ""
            info["line_weight"] = 0

        # Font info (if has text)
        if info.get("has_text"):
            try:
                font = shape.TextFrame.TextRange.Font
                info["font_name"] = str(_safe_attr(font, "Name", ""))
                info["font_size"] = float(_safe_attr(font, "Size", 0))
                info["font_bold"] = bool(font.Bold)
                info["font_italic"] = bool(font.Italic)
                try:
                    rgb = font.Color.RGB
                    r = rgb & 0xFF
                    g = (rgb >> 8) & 0xFF
                    b = (rgb >> 16) & 0xFF
                    info["font_color"] = f"#{r:02X}{g:02X}{b:02X}"
                except Exception:
                    info["font_color"] = ""
            except Exception:
                pass

        # Alt text
        try:
            info["alt_text"] = shape.AlternativeText
        except Exception:
            info["alt_text"] = ""

        return json.dumps(info, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 33: add_textbox
# ---------------------------------------------------------------------------

@mcp.tool()
def add_textbox(
    slide_index: int,
    text: str,
    left: float,
    top: float,
    width: float,
    height: float,
    font_size: float = 18,
    font_name: str = "",
    font_color: str = "",
    bold: bool = False,
    italic: bool = False,
    alignment: str = "left",
) -> str:
    """Add a text box to a slide with optional font formatting and alignment."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        shape = slide.Shapes.AddTextbox(
            1,  # msoTextOrientationHorizontal
            _inches_to_points(left),
            _inches_to_points(top),
            _inches_to_points(width),
            _inches_to_points(height),
        )
        tr = shape.TextFrame.TextRange
        tr.Text = text

        # Font formatting
        tr.Font.Size = font_size
        if font_name:
            tr.Font.Name = font_name
        if font_color:
            tr.Font.Color.RGB = _parse_color(font_color)
        tr.Font.Bold = -1 if bold else 0
        tr.Font.Italic = -1 if italic else 0

        # Alignment
        align_map = {"left": 1, "center": 2, "right": 3, "justify": 4}
        align_val = align_map.get(alignment.lower(), 1)
        tr.ParagraphFormat.Alignment = align_val

        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 34: modify_text
# ---------------------------------------------------------------------------

@mcp.tool()
def modify_text(
    slide_index: int,
    shape_name: str,
    text: str = "",
    font_size: float = 0,
    font_name: str = "",
    font_color: str = "",
    bold: int = -1,
    italic: int = -1,
    alignment: str = "",
) -> str:
    """Change text and/or formatting on an existing shape. Only non-default values are applied.
    bold/italic: -1=unchanged, 0=off, 1=on."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        tr = shape.TextFrame.TextRange
        if text:
            tr.Text = text
        if font_size > 0:
            tr.Font.Size = font_size
        if font_name:
            tr.Font.Name = font_name
        if font_color:
            tr.Font.Color.RGB = _parse_color(font_color)
        if bold != -1:
            tr.Font.Bold = -1 if bold == 1 else 0
        if italic != -1:
            tr.Font.Italic = -1 if italic == 1 else 0
        if alignment:
            align_map = {"left": 1, "center": 2, "right": 3, "justify": 4}
            align_val = align_map.get(alignment.lower())
            if align_val is not None:
                tr.ParagraphFormat.Alignment = align_val

        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 35: add_shape
# ---------------------------------------------------------------------------

@mcp.tool()
def add_shape(
    slide_index: int,
    shape_type: str,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: str = "",
    line_color: str = "",
    text: str = "",
) -> str:
    """Add an auto-shape to a slide. shape_type must be a key from AUTOSHAPE_MAP."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        type_id = AUTOSHAPE_MAP.get(shape_type.lower())
        if type_id is None:
            return json.dumps({
                "error": f"Unknown shape_type '{shape_type}'. Supported: {', '.join(AUTOSHAPE_MAP.keys())}"
            }, indent=2)

        shape = slide.Shapes.AddShape(
            type_id,
            _inches_to_points(left),
            _inches_to_points(top),
            _inches_to_points(width),
            _inches_to_points(height),
        )

        if fill_color:
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = _parse_color(fill_color)
        if line_color:
            shape.Line.ForeColor.RGB = _parse_color(line_color)
        if text:
            shape.TextFrame.TextRange.Text = text

        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 36: add_line
# ---------------------------------------------------------------------------

@mcp.tool()
def add_line(
    slide_index: int,
    begin_x: float,
    begin_y: float,
    end_x: float,
    end_y: float,
    color: str = "",
    weight: float = 1.0,
    dash_style: int = 1,
    arrow_head: bool = False,
) -> str:
    """Add a line to a slide. Coordinates in inches."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        shape = slide.Shapes.AddLine(
            _inches_to_points(begin_x),
            _inches_to_points(begin_y),
            _inches_to_points(end_x),
            _inches_to_points(end_y),
        )

        shape.Line.Weight = weight
        shape.Line.DashStyle = dash_style
        if color:
            shape.Line.ForeColor.RGB = _parse_color(color)
        if arrow_head:
            shape.Line.EndArrowheadStyle = 2  # msoArrowheadTriangle

        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 37: modify_shape
# ---------------------------------------------------------------------------

@mcp.tool()
def modify_shape(
    slide_index: int,
    shape_name: str,
    left: float = -1,
    top: float = -1,
    width: float = -1,
    height: float = -1,
    rotation: float = -1,
    fill_color: str = "",
    line_color: str = "",
    name: str = "",
) -> str:
    """Update shape properties. Only non-default values are applied. Positions in inches."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        if left >= 0:
            shape.Left = _inches_to_points(left)
        if top >= 0:
            shape.Top = _inches_to_points(top)
        if width >= 0:
            shape.Width = _inches_to_points(width)
        if height >= 0:
            shape.Height = _inches_to_points(height)
        if rotation >= 0:
            shape.Rotation = rotation
        if fill_color:
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = _parse_color(fill_color)
        if line_color:
            shape.Line.ForeColor.RGB = _parse_color(line_color)
        if name:
            shape.Name = name

        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 38: delete_shape
# ---------------------------------------------------------------------------

@mcp.tool()
def delete_shape(slide_index: int, shape_name: str) -> str:
    """Delete a shape from a slide by name or index."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)
        shape.Delete()
        return json.dumps({
            "status": "deleted",
            "shape_name": shape_name,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 39: group_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def group_shapes(slide_index: int, shape_names_json: str, group_name: str = "") -> str:
    """Group multiple shapes together. shape_names_json is a JSON array of shape names."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        names = json.loads(shape_names_json)

        import win32com.client
        names_array = win32com.client.VARIANT(0x2008, names)  # VT_ARRAY | VT_BSTR
        group = slide.Shapes.Range(names_array).Group()

        if group_name:
            group.Name = group_name

        return json.dumps(shape_to_dict(group), indent=2)
    except json.JSONDecodeError as je:
        return json.dumps({"error": f"Invalid JSON: {je}"}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 40: ungroup_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def ungroup_shapes(slide_index: int, group_name: str) -> str:
    """Ungroup a grouped shape into its individual shapes."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        group = get_shape(slide, group_name)
        ungrouped = group.Ungroup()

        shapes = []
        for i in range(1, ungrouped.Count + 1):
            shapes.append(shape_to_dict(ungrouped(i)))

        return json.dumps({
            "status": "ungrouped",
            "shapes": shapes,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 41: add_hyperlink
# ---------------------------------------------------------------------------

@mcp.tool()
def add_hyperlink(slide_index: int, shape_name: str, url: str, tooltip: str = "") -> str:
    """Add a hyperlink to a shape."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        shape.ActionSettings(1).Hyperlink.Address = url
        if tooltip:
            shape.ActionSettings(1).Hyperlink.ScreenTip = tooltip

        return json.dumps({
            "status": "added",
            "shape_name": shape_name,
            "url": url,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 42: add_connector
# ---------------------------------------------------------------------------

@mcp.tool()
def add_connector(
    slide_index: int,
    shape1_name: str,
    shape2_name: str,
    connector_type: str = "straight",
) -> str:
    """Add a connector between two shapes. Types: straight, elbow, curve."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape1 = get_shape(slide, shape1_name)
        shape2 = get_shape(slide, shape2_name)

        type_map = {"straight": 1, "elbow": 2, "curve": 3}
        conn_type = type_map.get(connector_type.lower(), 1)

        connector = slide.Shapes.AddConnector(conn_type, 0, 0, 100, 100)
        connector.ConnectorFormat.BeginConnect(shape1, 1)
        connector.ConnectorFormat.EndConnect(shape2, 1)
        connector.RerouteConnections()

        return json.dumps(shape_to_dict(connector), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 43: align_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def align_shapes(slide_index: int, shape_names_json: str, alignment: str = "center") -> str:
    """Align multiple shapes. Alignment: left, center, right, top, middle, bottom."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        names = json.loads(shape_names_json)

        align_map = {
            "left": 0,      # msoAlignLefts
            "center": 1,    # msoAlignCenters
            "right": 2,     # msoAlignRights
            "top": 3,       # msoAlignTops
            "middle": 4,    # msoAlignMiddles
            "bottom": 5,    # msoAlignBottoms
        }
        align_const = align_map.get(alignment.lower(), 1)

        import win32com.client
        names_array = win32com.client.VARIANT(0x2008, names)  # VT_ARRAY | VT_BSTR
        slide.Shapes.Range(names_array).Align(align_const, 0)

        return json.dumps({
            "status": "aligned",
            "alignment": alignment,
            "shapes": names,
        }, indent=2)
    except json.JSONDecodeError as je:
        return json.dumps({"error": f"Invalid JSON: {je}"}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 44: distribute_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def distribute_shapes(slide_index: int, shape_names_json: str, direction: str = "horizontal") -> str:
    """Distribute shapes evenly. Direction: horizontal or vertical."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        names = json.loads(shape_names_json)

        dist_map = {
            "horizontal": 0,  # msoDistributeHorizontally
            "vertical": 1,    # msoDistributeVertically
        }
        dist_const = dist_map.get(direction.lower(), 0)

        import win32com.client
        names_array = win32com.client.VARIANT(0x2008, names)  # VT_ARRAY | VT_BSTR
        slide.Shapes.Range(names_array).Distribute(dist_const, 0)

        return json.dumps({
            "status": "distributed",
            "direction": direction,
            "shapes": names,
        }, indent=2)
    except json.JSONDecodeError as je:
        return json.dumps({"error": f"Invalid JSON: {je}"}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 45: duplicate_shape
# ---------------------------------------------------------------------------

@mcp.tool()
def duplicate_shape(slide_index: int, shape_name: str, offset_x: float = 0.5, offset_y: float = 0.5) -> str:
    """Duplicate a shape and offset the copy by the given inches."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        new_shape = shape.Duplicate()
        new_shape.Left = shape.Left + _inches_to_points(offset_x)
        new_shape.Top = shape.Top + _inches_to_points(offset_y)

        return json.dumps(shape_to_dict(new_shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 46: set_shape_z_order
# ---------------------------------------------------------------------------

@mcp.tool()
def set_shape_z_order(slide_index: int, shape_name: str, action: str = "front") -> str:
    """Change the z-order of a shape. Actions: front, back, forward, backward."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        action_map = {
            "front": 0,     # msoBringToFront
            "back": 1,      # msoSendToBack
            "forward": 2,   # msoBringForward
            "backward": 3,  # msoSendBackward
        }
        action_const = action_map.get(action.lower(), 0)
        shape.ZOrder(action_const)

        return json.dumps({
            "status": "updated",
            "shape_name": shape_name,
            "z_order_action": action,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 47: bulk_add_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def bulk_add_shapes(slide_index: int, shapes_json: str) -> str:
    """Add multiple shapes from a JSON array. Each object: {type, left, top, width, height, fill_color, text, ...}."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shapes_spec = json.loads(shapes_json)
        results = []

        for spec in shapes_spec:
            try:
                shape_type = spec.get("type", "rectangle")
                type_id = AUTOSHAPE_MAP.get(shape_type.lower())
                if type_id is None:
                    results.append({"error": f"Unknown shape type '{shape_type}'"})
                    continue

                shape = slide.Shapes.AddShape(
                    type_id,
                    _inches_to_points(spec.get("left", 0)),
                    _inches_to_points(spec.get("top", 0)),
                    _inches_to_points(spec.get("width", 1)),
                    _inches_to_points(spec.get("height", 1)),
                )

                fill_color = spec.get("fill_color", "")
                if fill_color:
                    shape.Fill.Solid()
                    shape.Fill.ForeColor.RGB = _parse_color(fill_color)

                line_color = spec.get("line_color", "")
                if line_color:
                    shape.Line.ForeColor.RGB = _parse_color(line_color)

                text = spec.get("text", "")
                if text:
                    shape.TextFrame.TextRange.Text = text

                results.append(shape_to_dict(shape))
            except Exception as shape_err:
                results.append({"error": str(shape_err)})

        return json.dumps(results, indent=2)
    except json.JSONDecodeError as je:
        return json.dumps({"error": f"Invalid JSON: {je}"}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 48: set_placeholder_text
# ---------------------------------------------------------------------------

@mcp.tool()
def set_placeholder_text(
    slide_index: int,
    placeholder_index: int,
    text: str,
    font_size: float = 0,
    font_name: str = "",
    font_color: str = "",
    bold: bool = False,
    italic: bool = False,
) -> str:
    """Set text and formatting on a slide placeholder by index."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        placeholder = slide.Shapes.Placeholders(placeholder_index)
        tr = placeholder.TextFrame.TextRange
        tr.Text = text

        if font_size > 0:
            tr.Font.Size = font_size
        if font_name:
            tr.Font.Name = font_name
        if font_color:
            tr.Font.Color.RGB = _parse_color(font_color)
        if bold:
            tr.Font.Bold = -1  # msoTrue
        if italic:
            tr.Font.Italic = -1  # msoTrue

        return json.dumps({
            "status": "updated",
            "placeholder_index": placeholder_index,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 4 — Rich Media (14 tools)
# ═══════════════════════════════════════════════════════════════════════════════


@mcp.tool()
def insert_image(
    slide_index: int,
    image_path: str,
    left: float,
    top: float,
    width: float = 0,
    height: float = 0,
) -> str:
    """Insert an image from a local file path onto a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        abs_path = os.path.abspath(image_path)
        l_pts = _inches_to_points(left)
        t_pts = _inches_to_points(top)
        w_pts = _inches_to_points(width) if width > 0 else -1
        h_pts = _inches_to_points(height) if height > 0 else -1

        shape = slide.Shapes.AddPicture(
            abs_path,
            LinkToFile=0,
            SaveWithDocument=-1,
            Left=l_pts,
            Top=t_pts,
            Width=w_pts,
            Height=h_pts,
        )
        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def insert_image_from_url(
    slide_index: int,
    url: str,
    left: float,
    top: float,
    width: float = 0,
    height: float = 0,
) -> str:
    """Download an image from a URL and insert it onto a slide."""
    try:
        import urllib.request
        import tempfile

        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        # Download to a temp file
        tmp_path, _ = urllib.request.urlretrieve(url)

        l_pts = _inches_to_points(left)
        t_pts = _inches_to_points(top)
        w_pts = _inches_to_points(width) if width > 0 else -1
        h_pts = _inches_to_points(height) if height > 0 else -1

        shape = slide.Shapes.AddPicture(
            tmp_path,
            LinkToFile=0,
            SaveWithDocument=-1,
            Left=l_pts,
            Top=t_pts,
            Width=w_pts,
            Height=h_pts,
        )
        result = shape_to_dict(shape)

        # Clean up temp file
        try:
            os.remove(tmp_path)
        except OSError:
            pass

        return json.dumps(result, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def add_table(
    slide_index: int,
    rows: int,
    cols: int,
    left: float,
    top: float,
    width: float,
    height: float,
    data_json: str = "",
) -> str:
    """Add a table to a slide. Optionally fill with data from a 2D JSON array."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        l_pts = _inches_to_points(left)
        t_pts = _inches_to_points(top)
        w_pts = _inches_to_points(width)
        h_pts = _inches_to_points(height)

        shape = slide.Shapes.AddTable(rows, cols, l_pts, t_pts, w_pts, h_pts)
        table = shape.Table

        if data_json:
            data = json.loads(data_json)
            for r_idx, row_data in enumerate(data):
                for c_idx, cell_val in enumerate(row_data):
                    if r_idx < rows and c_idx < cols:
                        table.Cell(r_idx + 1, c_idx + 1).Shape.TextFrame.TextRange.Text = str(cell_val)

        return json.dumps({
            "status": "created",
            "shape_name": shape.Name,
            "rows": rows,
            "cols": cols,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def modify_table_cell(
    slide_index: int,
    shape_name: str,
    row: int,
    col: int,
    text: str,
    font_size: float = 0,
    font_name: str = "",
    font_color: str = "",
    bold: bool = False,
    fill_color: str = "",
) -> str:
    """Modify the text and formatting of a single table cell."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        cell = shape.Table.Cell(row, col)
        tr = cell.Shape.TextFrame.TextRange
        tr.Text = text

        if font_size > 0:
            tr.Font.Size = font_size
        if font_name:
            tr.Font.Name = font_name
        if font_color:
            tr.Font.Color.RGB = _parse_color(font_color)
        if bold:
            tr.Font.Bold = -1  # msoTrue
        if fill_color:
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = _parse_color(fill_color)

        return json.dumps({
            "status": "updated",
            "row": row,
            "col": col,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def bulk_fill_table(
    slide_index: int,
    shape_name: str,
    data_json: str,
) -> str:
    """Fill an entire table from a 2D JSON array."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        data = json.loads(data_json)
        table = shape.Table
        rows_filled = 0
        cols_filled = 0

        for r_idx, row_data in enumerate(data):
            rows_filled = max(rows_filled, r_idx + 1)
            for c_idx, cell_val in enumerate(row_data):
                cols_filled = max(cols_filled, c_idx + 1)
                table.Cell(r_idx + 1, c_idx + 1).Shape.TextFrame.TextRange.Text = str(cell_val)

        return json.dumps({
            "status": "filled",
            "rows": rows_filled,
            "cols": cols_filled,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def format_table(
    slide_index: int,
    shape_name: str,
    has_header: bool = True,
    band_rows: bool = True,
    band_cols: bool = False,
    first_col: bool = False,
    last_col: bool = False,
    header_color: str = "",
    row_color1: str = "",
    row_color2: str = "",
) -> str:
    """Format a table with banding, header styling, and alternating row colors."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        table = shape.Table

        # Table style properties — set via the parent shape, not the Table object
        try:
            shape.HasTable  # confirm it's a table shape
        except Exception:
            pass
        # These properties may not be available in all PowerPoint versions
        for prop, val in [
            ("FirstRow", has_header), ("BandRows", band_rows),
            ("BandColumns", band_cols), ("FirstCol", first_col),
            ("LastCol", last_col),
        ]:
            try:
                setattr(table, prop, -1 if val else 0)
            except Exception:
                pass  # Property not supported in this version

        if header_color:
            rgb = _parse_color(header_color)
            for c in range(1, table.Columns.Count + 1):
                cell = table.Cell(1, c)
                cell.Shape.Fill.Solid()
                cell.Shape.Fill.ForeColor.RGB = rgb

        if row_color1 or row_color2:
            start_row = 2 if has_header else 1
            for r in range(start_row, table.Rows.Count + 1):
                is_even = (r - start_row) % 2 == 1
                color_str = row_color2 if is_even else row_color1
                if color_str:
                    rgb = _parse_color(color_str)
                    for c in range(1, table.Columns.Count + 1):
                        cell = table.Cell(r, c)
                        cell.Shape.Fill.Solid()
                        cell.Shape.Fill.ForeColor.RGB = rgb

        return json.dumps({
            "status": "formatted",
            "shape_name": shape.Name,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def add_chart(
    slide_index: int,
    chart_type: str,
    data_json: str,
    left: float,
    top: float,
    width: float,
    height: float,
    title: str = "",
) -> str:
    """Add a chart to a slide. chart_type: column, bar, line, pie, area, scatter, doughnut, radar. data_json: {categories: [...], series: [{name, values: [...]}]}."""
    try:
        chart_type_map = {
            "column": 51,
            "bar": 57,
            "line": 4,
            "pie": 5,
            "area": 1,
            "scatter": 65,
            "doughnut": 18,
            "radar": 82,
        }

        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        type_id = chart_type_map.get(chart_type.lower(), 51)
        l_pts = _inches_to_points(left)
        t_pts = _inches_to_points(top)
        w_pts = _inches_to_points(width)
        h_pts = _inches_to_points(height)

        shape = slide.Shapes.AddChart2(-1, type_id, l_pts, t_pts, w_pts, h_pts)
        chart = shape.Chart

        data = json.loads(data_json)
        categories = data.get("categories", [])
        series_list = data.get("series", [])

        # Populate chart data via the embedded workbook
        chart.ChartData.Activate()
        wb = chart.ChartData.Workbook
        ws = wb.Worksheets(1)

        # Clear existing data
        ws.Cells.Clear()

        # Write categories in column A starting at row 2
        for i, cat in enumerate(categories):
            ws.Cells(i + 2, 1).Value = str(cat)

        # Write series
        for s_idx, series in enumerate(series_list):
            col = s_idx + 2
            ws.Cells(1, col).Value = series.get("name", f"Series {s_idx + 1}")
            for v_idx, val in enumerate(series.get("values", [])):
                ws.Cells(v_idx + 2, col).Value = val

        # Set the data range on the workbook so chart picks it up
        total_rows = len(categories) + 1
        total_cols = len(series_list) + 1
        if total_cols <= 26:
            last_col_letter = chr(ord('A') + total_cols - 1)
        else:
            last_col_letter = 'Z'
        range_str = f"A1:{last_col_letter}{total_rows}"
        try:
            # Try setting the range via the worksheet's ListObjects or named range
            ws.Range(f"A1:{last_col_letter}{total_rows}").Select()
        except Exception:
            pass  # Range selection not critical — chart auto-detects data

        wb.Close(True)

        if title:
            chart.HasTitle = -1  # msoTrue
            chart.ChartTitle.Text = title

        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def modify_chart(
    slide_index: int,
    shape_name: str,
    title: str = "",
    has_legend: bool = True,
    legend_position: str = "",
) -> str:
    """Modify chart title and legend properties."""
    try:
        legend_pos_map = {
            "bottom": 4,
            "top": 1,
            "left": 3,
            "right": 2,
        }

        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        chart = shape.Chart

        if title:
            chart.HasTitle = -1  # msoTrue
            chart.ChartTitle.Text = title

        chart.HasLegend = -1 if has_legend else 0

        if has_legend and legend_position:
            pos_id = legend_pos_map.get(legend_position.lower())
            if pos_id is not None:
                chart.Legend.Position = pos_id

        return json.dumps({
            "status": "updated",
            "shape_name": shape.Name,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def update_chart_data(
    slide_index: int,
    shape_name: str,
    data_json: str,
) -> str:
    """Update the data of an existing chart. data_json: {categories: [...], series: [{name, values: [...]}]}."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        chart = shape.Chart
        data = json.loads(data_json)
        categories = data.get("categories", [])
        series_list = data.get("series", [])

        chart.ChartData.Activate()
        wb = chart.ChartData.Workbook
        ws = wb.Worksheets(1)

        # Clear existing data
        ws.Cells.Clear()

        # Write categories in column A starting at row 2
        for i, cat in enumerate(categories):
            ws.Cells(i + 2, 1).Value = str(cat)

        # Write series
        for s_idx, series in enumerate(series_list):
            col = s_idx + 2
            ws.Cells(1, col).Value = series.get("name", f"Series {s_idx + 1}")
            for v_idx, val in enumerate(series.get("values", [])):
                ws.Cells(v_idx + 2, col).Value = val

        # Select the data range so chart picks it up
        total_rows = len(categories) + 1
        total_cols = len(series_list) + 1
        if total_cols <= 26:
            last_col_letter = chr(ord('A') + total_cols - 1)
        else:
            last_col_letter = 'Z'
        try:
            ws.Range(f"A1:{last_col_letter}{total_rows}").Select()
        except Exception:
            pass  # Chart auto-detects data range

        wb.Close(True)

        return json.dumps({
            "status": "updated",
            "shape_name": shape.Name,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def insert_video(
    slide_index: int,
    video_path: str,
    left: float,
    top: float,
    width: float,
    height: float,
) -> str:
    """Insert a video file onto a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        abs_path = os.path.abspath(video_path)
        l_pts = _inches_to_points(left)
        t_pts = _inches_to_points(top)
        w_pts = _inches_to_points(width)
        h_pts = _inches_to_points(height)

        shape = slide.Shapes.AddMediaObject2(
            abs_path,
            LinkToFile=0,
            SaveWithDocument=-1,
            Left=l_pts,
            Top=t_pts,
            Width=w_pts,
            Height=h_pts,
        )
        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def insert_audio(
    slide_index: int,
    audio_path: str,
    left: float = 0,
    top: float = 0,
) -> str:
    """Insert an audio file onto a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        abs_path = os.path.abspath(audio_path)
        l_pts = _inches_to_points(left)
        t_pts = _inches_to_points(top)

        shape = slide.Shapes.AddMediaObject2(
            abs_path,
            LinkToFile=0,
            SaveWithDocument=-1,
            Left=l_pts,
            Top=t_pts,
        )
        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def insert_ole_object(
    slide_index: int,
    file_path: str,
    left: float,
    top: float,
    width: float,
    height: float,
    as_icon: bool = False,
) -> str:
    """Insert an OLE embedded object (e.g., Excel, Word, PDF) onto a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        abs_path = os.path.abspath(file_path)
        l_pts = _inches_to_points(left)
        t_pts = _inches_to_points(top)
        w_pts = _inches_to_points(width)
        h_pts = _inches_to_points(height)

        shape = slide.Shapes.AddOLEObject(
            Left=l_pts,
            Top=t_pts,
            Width=w_pts,
            Height=h_pts,
            FileName=abs_path,
            DisplayAsIcon=-1 if as_icon else 0,
        )
        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def crop_image(
    slide_index: int,
    shape_name: str,
    crop_left: float = 0,
    crop_right: float = 0,
    crop_top: float = 0,
    crop_bottom: float = 0,
) -> str:
    """Crop an image shape by specifying crop values in points for each side."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        pf = shape.PictureFormat
        if crop_left:
            pf.CropLeft = crop_left
        if crop_right:
            pf.CropRight = crop_right
        if crop_top:
            pf.CropTop = crop_top
        if crop_bottom:
            pf.CropBottom = crop_bottom

        return json.dumps({
            "status": "cropped",
            "shape_name": shape.Name,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def replace_image(
    slide_index: int,
    shape_name: str,
    new_image_path: str,
) -> str:
    """Replace an existing image shape with a new image, preserving position and size."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        old_shape = get_shape(slide, shape_name)

        # Save position and size
        old_left = old_shape.Left
        old_top = old_shape.Top
        old_width = old_shape.Width
        old_height = old_shape.Height

        # Delete old shape
        old_shape.Delete()

        # Insert new image at same position and size
        abs_path = os.path.abspath(new_image_path)
        new_shape = slide.Shapes.AddPicture(
            abs_path,
            LinkToFile=0,
            SaveWithDocument=-1,
            Left=old_left,
            Top=old_top,
            Width=old_width,
            Height=old_height,
        )
        return json.dumps(shape_to_dict(new_shape), indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 5 — Design & Themes (12 tools)
# ═══════════════════════════════════════════════════════════════════════════════


@mcp.tool()
def get_theme_info() -> str:
    """Get theme name, colors overview, and fonts from the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)

        # Theme name
        theme_name = ""
        try:
            theme_name = pres.SlideMaster.Theme.Name
        except Exception:
            try:
                theme_name = pres.Designs(1).Name
            except Exception:
                theme_name = "Unknown"

        # Theme colors overview
        colors = []
        color_names = {
            1: "Background1", 2: "Text1", 3: "Background2", 4: "Text2",
            5: "Accent1", 6: "Accent2", 7: "Accent3", 8: "Accent4",
            9: "Accent5", 10: "Accent6", 11: "Hyperlink", 12: "FollowedHyperlink",
        }
        try:
            scheme = pres.SlideMaster.Theme.ThemeColorScheme
            for i in range(1, 13):
                try:
                    c = scheme(i)
                    bgr = c.RGB
                    r = bgr & 0xFF
                    g = (bgr >> 8) & 0xFF
                    b = (bgr >> 16) & 0xFF
                    colors.append({
                        "slot": i,
                        "name": color_names.get(i, f"Color{i}"),
                        "rgb_hex": f"#{r:02X}{g:02X}{b:02X}",
                    })
                except Exception:
                    colors.append({"slot": i, "name": color_names.get(i, f"Color{i}"), "rgb_hex": "N/A"})
        except Exception:
            pass

        # Theme fonts
        major_font = ""
        minor_font = ""
        try:
            font_scheme = pres.SlideMaster.Theme.ThemeFontScheme
            major_font = font_scheme.MajorFont.Item(1).Name
            minor_font = font_scheme.MinorFont.Item(1).Name
        except Exception:
            pass

        return json.dumps({
            "name": theme_name,
            "colors": colors,
            "fonts": {"major_font": major_font, "minor_font": minor_font},
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def apply_theme(theme_path: str) -> str:
    """Apply a .thmx theme file to the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        abs_path = os.path.abspath(theme_path)
        pres.ApplyTheme(abs_path)
        return json.dumps({
            "status": "applied",
            "theme_path": abs_path,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def get_theme_colors() -> str:
    """Get all 12 theme color slots with their RGB hex values."""
    try:
        app = get_app()
        pres = get_pres(app)

        color_names = {
            1: "Background1", 2: "Text1", 3: "Background2", 4: "Text2",
            5: "Accent1", 6: "Accent2", 7: "Accent3", 8: "Accent4",
            9: "Accent5", 10: "Accent6", 11: "Hyperlink", 12: "FollowedHyperlink",
        }

        scheme = pres.SlideMaster.Theme.ThemeColorScheme
        colors = []
        for i in range(1, 13):
            try:
                c = scheme(i)
                bgr = c.RGB
                r = bgr & 0xFF
                g = (bgr >> 8) & 0xFF
                b = (bgr >> 16) & 0xFF
                colors.append({
                    "slot": i,
                    "name": color_names.get(i, f"Color{i}"),
                    "rgb_hex": f"#{r:02X}{g:02X}{b:02X}",
                })
            except Exception:
                colors.append({
                    "slot": i,
                    "name": color_names.get(i, f"Color{i}"),
                    "rgb_hex": "N/A",
                })
        return json.dumps(colors, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def set_theme_color(slot: str, color: str) -> str:
    """Set a theme color slot by name (e.g. 'Accent1', 'Text1') to a new color ('#RRGGBB' or 'R,G,B')."""
    try:
        app = get_app()
        pres = get_pres(app)

        slot_map = {
            "background1": 1, "text1": 2, "background2": 3, "text2": 4,
            "accent1": 5, "accent2": 6, "accent3": 7, "accent4": 8,
            "accent5": 9, "accent6": 10, "hyperlink": 11, "followedhyperlink": 12,
        }
        idx = slot_map.get(slot.lower().replace(" ", ""))
        if idx is None:
            raise ValueError(f"Unknown slot '{slot}'. Valid: {', '.join(slot_map.keys())}")

        scheme = pres.SlideMaster.Theme.ThemeColorScheme
        scheme(idx).RGB = _parse_color(color)

        return json.dumps({
            "status": "updated",
            "slot": slot,
            "color": color,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def get_theme_fonts() -> str:
    """Get the major (headings) and minor (body) theme fonts."""
    try:
        app = get_app()
        pres = get_pres(app)
        font_scheme = pres.SlideMaster.Theme.ThemeFontScheme
        major_font = font_scheme.MajorFont.Item(1).Name
        minor_font = font_scheme.MinorFont.Item(1).Name
        return json.dumps({
            "major_font": major_font,
            "minor_font": minor_font,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def set_theme_fonts(major_font: str = "", minor_font: str = "") -> str:
    """Set the major (headings) and/or minor (body) theme fonts."""
    try:
        app = get_app()
        pres = get_pres(app)
        font_scheme = pres.SlideMaster.Theme.ThemeFontScheme
        updated = []

        if major_font:
            try:
                font_scheme.MajorFont.Item(1).Name = major_font
                updated.append("major_font")
            except Exception:
                # Alternative approach: iterate font items
                for i in range(1, font_scheme.MajorFont.Count + 1):
                    try:
                        font_scheme.MajorFont.Item(i).Name = major_font
                    except Exception:
                        pass
                updated.append("major_font")

        if minor_font:
            try:
                font_scheme.MinorFont.Item(1).Name = minor_font
                updated.append("minor_font")
            except Exception:
                for i in range(1, font_scheme.MinorFont.Count + 1):
                    try:
                        font_scheme.MinorFont.Item(i).Name = minor_font
                    except Exception:
                        pass
                updated.append("minor_font")

        return json.dumps({
            "status": "updated",
            "major_font": major_font or "(unchanged)",
            "minor_font": minor_font or "(unchanged)",
            "updated_fields": updated,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def get_master_layouts(master_index: int = 1) -> str:
    """List all custom layouts for a given slide master (1-based index)."""
    try:
        app = get_app()
        pres = get_pres(app)
        layouts_col = pres.Designs(master_index).SlideMaster.CustomLayouts
        layouts = []
        for i in range(1, layouts_col.Count + 1):
            layout = layouts_col(i)
            layouts.append({
                "index": i,
                "name": _safe_attr(layout, "Name", f"Layout{i}"),
            })
        return json.dumps(layouts, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def modify_master_placeholder(
    master_index: int,
    layout_index: int,
    placeholder_index: int,
    font_size: float = 0,
    font_name: str = "",
    font_color: str = "",
    bold: bool = False,
    italic: bool = False,
) -> str:
    """Modify font formatting of a placeholder in a master layout."""
    try:
        app = get_app()
        pres = get_pres(app)
        layout = pres.Designs(master_index).SlideMaster.CustomLayouts(layout_index)
        ph = layout.Shapes.Placeholders(placeholder_index)
        font = ph.TextFrame.TextRange.Font

        if font_size > 0:
            font.Size = font_size
        if font_name:
            font.Name = font_name
        if font_color:
            font.Color.RGB = _parse_color(font_color)
        if bold:
            font.Bold = -1  # msoTrue
        if italic:
            font.Italic = -1  # msoTrue

        return json.dumps({
            "status": "updated",
            "master_index": master_index,
            "layout_index": layout_index,
            "placeholder_index": placeholder_index,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def set_background(
    slide_index: int = 0,
    color: str = "",
    image_path: str = "",
    gradient_color1: str = "",
    gradient_color2: str = "",
) -> str:
    """Set background of a slide (by 1-based index) or the slide master (index=0). Supports solid color, image, or two-color gradient."""
    try:
        app = get_app()
        pres = get_pres(app)

        if slide_index == 0:
            bg = pres.SlideMaster.Background
        else:
            slide = get_slide(pres, slide_index)
            slide.FollowMasterBackground = 0  # msoFalse
            bg = slide.Background

        fill = bg.Fill

        if color:
            fill.Solid()
            fill.ForeColor.RGB = _parse_color(color)
        elif image_path:
            abs_path = os.path.abspath(image_path)
            fill.UserPicture(abs_path)
        elif gradient_color1 and gradient_color2:
            # msoGradientHorizontal=1, variant=1
            fill.TwoColorGradient(1, 1)
            fill.ForeColor.RGB = _parse_color(gradient_color1)
            fill.BackColor.RGB = _parse_color(gradient_color2)
        else:
            raise ValueError("Provide 'color', 'image_path', or both 'gradient_color1' and 'gradient_color2'.")

        return json.dumps({
            "status": "updated",
            "target": "slide_master" if slide_index == 0 else f"slide_{slide_index}",
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def get_placeholders(slide_index: int) -> str:
    """List all placeholders on a slide with their properties."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        placeholders = []
        for i in range(1, slide.Shapes.Placeholders.Count + 1):
            ph = slide.Shapes.Placeholders(i)
            has_text = False
            text = ""
            try:
                if ph.HasTextFrame:
                    has_text = True
                    text = ph.TextFrame.TextRange.Text
            except Exception:
                pass
            placeholders.append({
                "index": i,
                "name": _safe_attr(ph, "Name", ""),
                "type": _safe_attr(ph.PlaceholderFormat, "Type", ""),
                "left": _points_to_inches(ph.Left),
                "top": _points_to_inches(ph.Top),
                "width": _points_to_inches(ph.Width),
                "height": _points_to_inches(ph.Height),
                "has_text": has_text,
                "text": text,
            })
        return json.dumps(placeholders, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def add_custom_layout(master_index: int = 1, name: str = "Custom Layout") -> str:
    """Add a new custom layout to a slide master."""
    try:
        app = get_app()
        pres = get_pres(app)
        master = pres.Designs(master_index).SlideMaster
        new_index = master.CustomLayouts.Count + 1
        layout = master.CustomLayouts.Add(new_index)
        layout.Name = name
        return json.dumps({
            "status": "created",
            "name": name,
            "index": new_index,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def copy_master_from(source_path: str) -> str:
    """Copy the first slide master/design from another presentation file into the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        abs_path = os.path.abspath(source_path)

        # Open source read-only (ReadOnly=-1 is msoTrue)
        source_pres = app.Presentations.Open(abs_path, ReadOnly=-1)
        try:
            # Clone each design from source into active presentation
            designs_copied = 0
            for i in range(1, source_pres.Designs.Count + 1):
                pres.Designs.Clone(source_pres.Designs(i))
                designs_copied += 1
        finally:
            source_pres.Close()

        return json.dumps({
            "status": "copied",
            "source_path": abs_path,
            "designs_count": designs_copied,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 6 — Advanced Operations (18 tools)
# ═══════════════════════════════════════════════════════════════════════════════


@mcp.tool()
def find_and_replace(find_text: str, replace_text: str, match_case: bool = False) -> str:
    """Find and replace text across all slides in the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        replacements_count = 0
        slides_affected = set()

        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for sh in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(sh)
                try:
                    if shape.HasTextFrame:
                        text = shape.TextFrame.TextRange.Text
                        if match_case:
                            count = text.count(find_text)
                        else:
                            count = text.lower().count(find_text.lower())
                        if count > 0:
                            if match_case:
                                new_text = text.replace(find_text, replace_text)
                            else:
                                # Case-insensitive replace
                                import re
                                new_text = re.sub(re.escape(find_text), replace_text, text, flags=re.IGNORECASE)
                            shape.TextFrame.TextRange.Text = new_text
                            replacements_count += count
                            slides_affected.add(si)
                except Exception:
                    continue

        return json.dumps({
            "status": "completed",
            "replacements_count": replacements_count,
            "slides_affected": len(slides_affected),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def extract_all_text(include_notes: bool = True) -> str:
    """Extract all text from every slide (and optionally speaker notes) in the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        results = []

        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            texts = []
            for sh in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(sh)
                try:
                    if shape.HasTextFrame:
                        t = shape.TextFrame.TextRange.Text
                        if t.strip():
                            texts.append(t)
                except Exception:
                    continue

            notes = ""
            if include_notes:
                try:
                    notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
                except Exception:
                    pass

            results.append({
                "slide_index": si,
                "texts": texts,
                "notes": notes,
            })

        return json.dumps(results, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def get_presentation_outline() -> str:
    """Get a hierarchical outline of the active presentation showing slides, layouts, and shapes."""
    try:
        app = get_app()
        pres = get_pres(app)
        slides_list = []

        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            layout_name = ""
            try:
                layout_name = slide.CustomLayout.Name
            except Exception:
                pass

            shapes_list = []
            for sh in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(sh)
                text_preview = ""
                try:
                    if shape.HasTextFrame:
                        full_text = shape.TextFrame.TextRange.Text
                        text_preview = full_text[:100]
                except Exception:
                    pass

                shape_type = 0
                try:
                    shape_type = int(shape.Type)
                except Exception:
                    pass

                shapes_list.append({
                    "name": shape.Name,
                    "type": shape_type,
                    "text_preview": text_preview,
                })

            slides_list.append({
                "slide_index": si,
                "layout": layout_name,
                "shapes": shapes_list,
            })

        return json.dumps({"slides": slides_list}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def merge_presentations(file_paths_json: str, insert_at: int = 0) -> str:
    """Merge slides from multiple presentation files into the active presentation.

    file_paths_json: JSON array of file paths, e.g. '["C:/a.pptx","C:/b.pptx"]'.
    insert_at: slide index after which to insert (0 = end of presentation).
    """
    try:
        app = get_app()
        pres = get_pres(app)
        file_paths = json.loads(file_paths_json)
        total_inserted = 0
        insert_position = insert_at if insert_at > 0 else pres.Slides.Count

        for fp in file_paths:
            abs_path = os.path.abspath(fp)
            before_count = pres.Slides.Count
            pres.Slides.InsertFromFile(abs_path, insert_position)
            added = pres.Slides.Count - before_count
            total_inserted += added
            insert_position += added

        return json.dumps({
            "status": "merged",
            "files_merged": len(file_paths),
            "total_slides_inserted": total_inserted,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def apply_template(template_path: str) -> str:
    """Apply a PowerPoint template (.potx/.pptx) to the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        abs_path = os.path.abspath(template_path)
        pres.ApplyTemplate(abs_path)
        return json.dumps({
            "status": "applied",
            "template": abs_path,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def bulk_format_text(criteria_json: str) -> str:
    """Apply formatting to all matching text across the presentation.

    criteria_json: JSON with {find_text, font_size, font_name, font_color, bold, italic}.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        criteria = json.loads(criteria_json)
        find_text = criteria.get("find_text", "")
        matches_formatted = 0

        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for sh in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(sh)
                try:
                    if not shape.HasTextFrame:
                        continue
                    tr = shape.TextFrame.TextRange
                    for ri in range(1, tr.Runs.Count + 1):
                        run = tr.Runs(ri)
                        if find_text.lower() in run.Text.lower():
                            font = run.Font
                            if "font_size" in criteria and criteria["font_size"] is not None:
                                font.Size = criteria["font_size"]
                            if "font_name" in criteria and criteria["font_name"] is not None:
                                font.Name = criteria["font_name"]
                            if "font_color" in criteria and criteria["font_color"] is not None:
                                font.Color.RGB = _parse_color(criteria["font_color"])
                            if "bold" in criteria and criteria["bold"] is not None:
                                font.Bold = -1 if criteria["bold"] else 0
                            if "italic" in criteria and criteria["italic"] is not None:
                                font.Italic = -1 if criteria["italic"] else 0
                            matches_formatted += 1
                except Exception:
                    continue

        return json.dumps({
            "status": "formatted",
            "matches_formatted": matches_formatted,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def add_animation(slide_index: int, shape_name: str, effect: str = "appear",
                  duration: float = 0.5, delay: float = 0, trigger: str = "on_click") -> str:
    """Add an animation effect to a shape on a slide.

    effect: one of appear, fade, fly_in, wipe, split, wheel, grow_shrink.
    trigger: on_click, with_previous, after_previous.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        effect_key = effect.lower().replace(" ", "_")
        effect_id = ANIMATION_MAP.get(effect_key)
        if effect_id is None:
            raise ValueError(f"Unknown effect '{effect}'. Available: {list(ANIMATION_MAP.keys())}")

        trigger_map = {"on_click": 1, "with_previous": 2, "after_previous": 3}
        trigger_const = trigger_map.get(trigger.lower().replace(" ", "_"), 1)

        anim_effect = slide.TimeLine.MainSequence.AddEffect(shape, effect_id, trigger=trigger_const)
        anim_effect.Timing.Duration = duration
        anim_effect.Timing.TriggerDelayTime = delay

        return json.dumps({
            "status": "added",
            "effect": effect_key,
            "duration": duration,
            "delay": delay,
            "trigger": trigger,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def remove_animation(slide_index: int, shape_name: str) -> str:
    """Remove all animation effects from a specific shape on a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        seq = slide.TimeLine.MainSequence
        removed = 0

        # Iterate backwards to safely delete
        for i in range(seq.Count, 0, -1):
            effect = seq(i)
            try:
                if effect.Shape.Name == shape_name:
                    effect.Delete()
                    removed += 1
            except Exception:
                continue

        return json.dumps({
            "status": "removed",
            "shape_name": shape_name,
            "count": removed,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def get_animations(slide_index: int) -> str:
    """Get all animation effects on a slide with their properties."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        seq = slide.TimeLine.MainSequence
        animations = []

        for i in range(1, seq.Count + 1):
            effect = seq(i)
            shape_name = ""
            try:
                shape_name = effect.Shape.Name
            except Exception:
                pass

            effect_type = 0
            try:
                effect_type = int(effect.EffectType)
            except Exception:
                pass

            duration = 0
            delay = 0
            trigger_type = 0
            try:
                duration = effect.Timing.Duration
                delay = effect.Timing.TriggerDelayTime
                trigger_type = int(effect.Timing.TriggerType)
            except Exception:
                pass

            animations.append({
                "index": i,
                "shape_name": shape_name,
                "effect_type": effect_type,
                "duration": duration,
                "delay": delay,
                "trigger_type": trigger_type,
            })

        return json.dumps({
            "slide_index": slide_index,
            "animations": animations,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def reorder_animations(slide_index: int, order_json: str) -> str:
    """Reorder animation effects on a slide by specifying shape names in desired order.

    order_json: JSON array of shape names, e.g. '["Title","Subtitle","Image1"]'.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        desired_order = json.loads(order_json)
        seq = slide.TimeLine.MainSequence

        # Reorder by moving effects to match the desired sequence
        current_pos = 1
        for name in desired_order:
            for i in range(current_pos, seq.Count + 1):
                try:
                    if seq(i).Shape.Name == name:
                        if i != current_pos:
                            seq(i).MoveTo(current_pos)
                        current_pos += 1
                        break
                except Exception:
                    continue

        return json.dumps({
            "status": "reordered",
            "slide_index": slide_index,
            "order": desired_order,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def bulk_speaker_notes(notes_json: str) -> str:
    """Set speaker notes on multiple slides at once.

    notes_json: JSON array of {slide_index, notes}, e.g. '[{"slide_index":1,"notes":"Hello"}]'.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        items = json.loads(notes_json)
        updated = 0

        for item in items:
            si = item["slide_index"]
            notes_text = item["notes"]
            slide = get_slide(pres, si)
            slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes_text
            updated += 1

        return json.dumps({
            "status": "updated",
            "count": updated,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def clone_formatting(slide_index: int, source_shape: str, target_shapes_json: str) -> str:
    """Clone formatting from a source shape to multiple target shapes on the same slide.

    target_shapes_json: JSON array of shape names, e.g. '["Shape2","Shape3"]'.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        src = get_shape(slide, source_shape)
        targets = json.loads(target_shapes_json)
        updated = 0

        # Read source formatting
        src_font = {}
        try:
            if src.HasTextFrame:
                f = src.TextFrame.TextRange.Font
                src_font["Name"] = f.Name
                src_font["Size"] = f.Size
                src_font["Bold"] = f.Bold
                src_font["Italic"] = f.Italic
                try:
                    src_font["ColorRGB"] = f.Color.RGB
                except Exception:
                    pass
        except Exception:
            pass

        src_fill = {}
        try:
            fill = src.Fill
            src_fill["Type"] = int(fill.Type)
            if int(fill.Type) == 1:  # msoFillSolid
                src_fill["ForeColorRGB"] = fill.ForeColor.RGB
        except Exception:
            pass

        src_line = {}
        try:
            line = src.Line
            src_line["Weight"] = line.Weight
            src_line["Visible"] = int(line.Visible)
            try:
                src_line["ForeColorRGB"] = line.ForeColor.RGB
            except Exception:
                pass
        except Exception:
            pass

        for tname in targets:
            tgt = get_shape(slide, tname)
            try:
                if src_font and tgt.HasTextFrame:
                    f = tgt.TextFrame.TextRange.Font
                    if "Name" in src_font:
                        f.Name = src_font["Name"]
                    if "Size" in src_font:
                        f.Size = src_font["Size"]
                    if "Bold" in src_font:
                        f.Bold = src_font["Bold"]
                    if "Italic" in src_font:
                        f.Italic = src_font["Italic"]
                    if "ColorRGB" in src_font:
                        f.Color.RGB = src_font["ColorRGB"]
            except Exception:
                pass
            try:
                if src_fill.get("Type") == 1:
                    tgt.Fill.Solid()
                    tgt.Fill.ForeColor.RGB = src_fill["ForeColorRGB"]
            except Exception:
                pass
            try:
                if src_line:
                    tgt.Line.Visible = src_line.get("Visible", -1)
                    if "Weight" in src_line:
                        tgt.Line.Weight = src_line["Weight"]
                    if "ForeColorRGB" in src_line:
                        tgt.Line.ForeColor.RGB = src_line["ForeColorRGB"]
            except Exception:
                pass
            updated += 1

        return json.dumps({
            "status": "cloned",
            "source_shape": source_shape,
            "targets_updated": updated,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def search_shapes(query: str, search_text: bool = True, search_names: bool = True) -> str:
    """Search for shapes across all slides by text content and/or shape name."""
    try:
        app = get_app()
        pres = get_pres(app)
        matches = []
        query_lower = query.lower()

        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for sh in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(sh)
                shape_name = shape.Name
                text = ""
                try:
                    if shape.HasTextFrame:
                        text = shape.TextFrame.TextRange.Text
                except Exception:
                    pass

                match_type = []
                if search_names and query_lower in shape_name.lower():
                    match_type.append("name")
                if search_text and text and query_lower in text.lower():
                    match_type.append("text")

                if match_type:
                    matches.append({
                        "slide_index": si,
                        "shape_name": shape_name,
                        "match_type": match_type,
                        "text_preview": text[:100] if text else "",
                    })

        return json.dumps({
            "query": query,
            "matches": matches,
            "total_found": len(matches),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def rename_shape(slide_index: int, old_name: str, new_name: str) -> str:
    """Rename a shape on a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, old_name)
        shape.Name = new_name
        return json.dumps({
            "status": "renamed",
            "old_name": old_name,
            "new_name": new_name,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def lock_shape(slide_index: int, shape_name: str, lock: bool = True) -> str:
    """Lock or unlock a shape's aspect ratio and available lock properties."""
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        shape.LockAspectRatio = -1 if lock else 0

        # Apply additional locks if available
        try:
            locks = shape.Locks
            locks.LockPosition = lock
            locks.LockSize = lock
        except Exception:
            pass

        return json.dumps({
            "status": "locked" if lock else "unlocked",
            "shape_name": shape_name,
            "lock_aspect_ratio": lock,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def add_section(name: str, before_slide: int = 0) -> str:
    """Add a named section to the presentation.

    before_slide: 1-based slide index. 0 means add at the end.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        slide_idx = before_slide if before_slide > 0 else pres.Slides.Count + 1
        section_index = pres.SectionProperties.AddSection(slide_idx, name)
        return json.dumps({
            "status": "added",
            "name": name,
            "section_index": section_index,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def get_sections() -> str:
    """Get all sections in the active presentation with their properties."""
    try:
        app = get_app()
        pres = get_pres(app)
        sp = pres.SectionProperties
        sections = []

        for i in range(1, sp.Count + 1):
            sections.append({
                "section_index": i,
                "name": sp.Name(i),
                "slides_count": sp.SlidesCount(i),
                "first_slide": sp.FirstSlide(i),
            })

        return json.dumps({"sections": sections}, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


@mcp.tool()
def delete_section(section_index: int, delete_slides: bool = False) -> str:
    """Delete a section from the presentation.

    section_index: 1-based section index.
    delete_slides: if True, also delete the slides in that section.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        pres.SectionProperties.Delete(section_index, delete_slides)
        return json.dumps({
            "status": "deleted",
            "section_index": section_index,
            "slides_deleted": delete_slides,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 7 — Analysis & Export (13 tools)
# ═══════════════════════════════════════════════════════════════════════════════


# ---------------------------------------------------------------------------
# Tool 93: get_presentation_stats
# ---------------------------------------------------------------------------

@mcp.tool()
def get_presentation_stats() -> str:
    """Get comprehensive statistics about the active presentation.

    Returns slide count, total shapes, images, text boxes, tables, charts,
    and total word count across all slides.
    """
    try:
        app = get_app()
        pres = get_pres(app)

        total_shapes = 0
        total_images = 0
        total_text_boxes = 0
        total_tables = 0
        total_charts = 0
        total_word_count = 0

        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for j in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(j)
                total_shapes += 1
                type_int = int(_safe_attr(shape, "Type", 0))

                if type_int == 13:
                    total_images += 1
                if type_int == 17:
                    total_text_boxes += 1
                if type_int == 19:
                    total_tables += 1
                if type_int == 3:
                    total_charts += 1

                try:
                    if shape.HasTextFrame:
                        text = shape.TextFrame.TextRange.Text
                        words = text.split()
                        total_word_count += len(words)
                        # Also count as text box if it has text and isn't already counted
                        if type_int not in (13, 19, 3) and text.strip():
                            pass  # type 17 already counted above
                except Exception:
                    pass

        return json.dumps({
            "slide_count": pres.Slides.Count,
            "total_shapes": total_shapes,
            "total_images": total_images,
            "total_text_boxes": total_text_boxes,
            "total_tables": total_tables,
            "total_charts": total_charts,
            "total_word_count": total_word_count,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 94: export_slide_image
# ---------------------------------------------------------------------------

@mcp.tool()
def export_slide_image(slide_index: int, output_path: str, format: str = "png", width: int = 1920) -> str:
    """Export a single slide as an image file.

    slide_index: 1-based slide index.
    output_path: File path for the exported image.
    format: Image format (png, jpg, bmp, gif, tif). Default 'png'.
    width: Image width in pixels. Default 1920.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        slide = get_slide(pres, slide_index)

        abs_path = os.path.abspath(output_path)
        slide.Export(abs_path, format.upper(), width)

        return json.dumps({
            "status": "exported",
            "slide_index": slide_index,
            "path": abs_path,
            "format": format.lower(),
            "width": width,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 95: export_all_slides_images
# ---------------------------------------------------------------------------

@mcp.tool()
def export_all_slides_images(output_dir: str, format: str = "png", width: int = 1920) -> str:
    """Export all slides as individual image files.

    output_dir: Directory to save the images (created if it doesn't exist).
    format: Image format (png, jpg, bmp, gif, tif). Default 'png'.
    width: Image width in pixels. Default 1920.
    """
    try:
        app = get_app()
        pres = get_pres(app)

        abs_dir = os.path.abspath(output_dir)
        os.makedirs(abs_dir, exist_ok=True)

        exported = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            filename = f"slide_{i}.{format.lower()}"
            filepath = os.path.join(abs_dir, filename)
            slide.Export(filepath, format.upper(), width)
            exported.append(filepath)

        return json.dumps({
            "status": "exported",
            "count": len(exported),
            "output_dir": abs_dir,
            "format": format.lower(),
            "files": exported,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 96: export_pdf
# ---------------------------------------------------------------------------

@mcp.tool()
def export_pdf(output_path: str, slides_range: str = "", quality: str = "high") -> str:
    """Export the presentation as a PDF file.

    output_path: File path for the exported PDF.
    slides_range: Optional slide range (e.g. '1-5'). Empty for all slides.
    quality: 'high' (default) or 'low'.
    """
    try:
        app = get_app()
        pres = get_pres(app)

        abs_path = os.path.abspath(output_path)
        quality_val = 0 if quality.lower() == "high" else 1  # 0=Standard, 1=Minimum

        if slides_range:
            # Use ExportAsFixedFormat for range support
            # ppFixedFormatTypePDF = 2
            try:
                # Parse range
                parts = slides_range.split("-")
                range_from = int(parts[0].strip())
                range_to = int(parts[1].strip()) if len(parts) > 1 else range_from

                pres.ExportAsFixedFormat(
                    abs_path,
                    2,  # ppFixedFormatTypePDF
                    quality_val,
                    0,  # ppFixedFormatIntentScreen
                    False,  # msoCTrue for frame slides
                    1,  # ppPrintHandoutHorizontalFirst
                    1,  # ppPrintOutputSlides
                    range_from,
                    range_to,
                )
            except Exception:
                # Fallback to SaveAs
                pres.SaveAs(abs_path, 32)  # ppSaveAsPDF = 32
        else:
            try:
                pres.ExportAsFixedFormat(abs_path, 2, quality_val)
            except Exception:
                pres.SaveAs(abs_path, 32)

        return json.dumps({
            "status": "exported",
            "path": abs_path,
            "quality": quality,
            "slides_range": slides_range if slides_range else "all",
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 97: get_fonts_used
# ---------------------------------------------------------------------------

@mcp.tool()
def get_fonts_used() -> str:
    """Get a list of all unique fonts used in the presentation.

    Scans all text frames and individual text runs across every slide.
    """
    try:
        app = get_app()
        pres = get_pres(app)

        fonts = set()
        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for j in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(j)
                try:
                    if shape.HasTextFrame:
                        tr = shape.TextFrame.TextRange
                        # Check overall font
                        font_name = _safe_attr(tr.Font, "Name", "")
                        if font_name:
                            fonts.add(font_name)
                        # Check individual runs
                        try:
                            for r in range(1, tr.Runs().Count + 1):
                                run_font = _safe_attr(tr.Runs(r).Font, "Name", "")
                                if run_font:
                                    fonts.add(run_font)
                        except Exception:
                            pass
                except Exception:
                    pass

        return json.dumps({
            "fonts": sorted(list(fonts)),
            "count": len(fonts),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 98: get_linked_files
# ---------------------------------------------------------------------------

@mcp.tool()
def get_linked_files() -> str:
    """Get a list of all linked/embedded external files in the presentation.

    Checks OLE objects (Type 7), pictures (Type 13), and media (Type 16)
    for linked source paths.
    """
    try:
        app = get_app()
        pres = get_pres(app)

        linked = []
        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for j in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(j)
                type_int = int(_safe_attr(shape, "Type", 0))
                if type_int in (7, 13, 16):
                    source_path = ""
                    try:
                        source_path = shape.LinkFormat.SourceFullName
                    except Exception:
                        pass
                    if source_path:
                        linked.append({
                            "slide_index": si,
                            "shape_name": _safe_attr(shape, "Name", ""),
                            "type": type_int,
                            "source_path": source_path,
                        })

        return json.dumps({
            "linked_files": linked,
            "count": len(linked),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 99: check_accessibility
# ---------------------------------------------------------------------------

@mcp.tool()
def check_accessibility() -> str:
    """Check the presentation for accessibility issues.

    Checks for: images without alt text, shapes with no text and no alt text,
    and reports reading order per slide.
    """
    try:
        app = get_app()
        pres = get_pres(app)

        issues = []
        total_images = 0
        images_without_alt = 0
        slide_reading_orders = []

        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            reading_order = []

            for j in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(j)
                shape_name = _safe_attr(shape, "Name", "")
                type_int = int(_safe_attr(shape, "Type", 0))
                alt_text = _safe_attr(shape, "AlternativeText", "")

                reading_order.append({
                    "order": j,
                    "name": shape_name,
                    "type": type_int,
                })

                # Check images without alt text
                if type_int == 13:
                    total_images += 1
                    if not alt_text.strip():
                        images_without_alt += 1
                        issues.append({
                            "slide_index": si,
                            "shape_name": shape_name,
                            "issue": "Image without alt text",
                        })

                # Check shapes with no text and no alt text
                has_text = False
                try:
                    if shape.HasTextFrame:
                        text = shape.TextFrame.TextRange.Text
                        if text.strip():
                            has_text = True
                except Exception:
                    pass

                if not has_text and not alt_text.strip() and type_int != 13:
                    issues.append({
                        "slide_index": si,
                        "shape_name": shape_name,
                        "issue": "Shape with no text and no alt text",
                    })

            slide_reading_orders.append({
                "slide_index": si,
                "reading_order": reading_order,
            })

        return json.dumps({
            "issues": issues,
            "slide_count": pres.Slides.Count,
            "images_without_alt_text": images_without_alt,
            "total_images": total_images,
            "reading_orders": slide_reading_orders,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 100: get_slide_thumbnails_base64
# ---------------------------------------------------------------------------

@mcp.tool()
def get_slide_thumbnails_base64(slide_indices_json: str = "", width: int = 320) -> str:
    """Export slide thumbnails as base64-encoded PNG strings.

    slide_indices_json: JSON array of 1-based slide indices, e.g. '[1,3,5]'.
                        Empty string means all slides.
    width: Thumbnail width in pixels. Default 320.
    """
    try:
        import base64
        import tempfile

        app = get_app()
        pres = get_pres(app)

        if slide_indices_json and slide_indices_json.strip():
            indices = json.loads(slide_indices_json)
        else:
            indices = list(range(1, pres.Slides.Count + 1))

        thumbnails = []
        for idx in indices:
            slide = get_slide(pres, idx)
            tmp_file = os.path.join(tempfile.gettempdir(), f"ppt_thumb_{idx}.png")
            try:
                slide.Export(tmp_file, "PNG", width)
                with open(tmp_file, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode("utf-8")
                thumbnails.append({
                    "slide_index": idx,
                    "base64_image": b64,
                    "format": "png",
                })
            finally:
                if os.path.exists(tmp_file):
                    os.remove(tmp_file)

        return json.dumps({
            "thumbnails": thumbnails,
            "count": len(thumbnails),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 101: compare_slides
# ---------------------------------------------------------------------------

@mcp.tool()
def compare_slides(slide_a: int, slide_b: int) -> str:
    """Compare two slides and show differences.

    slide_a: 1-based index of the first slide.
    slide_b: 1-based index of the second slide.
    Returns shape counts, text content, layout info, and differences.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        sa = get_slide(pres, slide_a)
        sb = get_slide(pres, slide_b)

        def slide_info(slide):
            shapes = {}
            texts = {}
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                name = _safe_attr(shape, "Name", "")
                shapes[name] = shape_to_dict(shape)
                try:
                    if shape.HasTextFrame:
                        texts[name] = shape.TextFrame.TextRange.Text
                except Exception:
                    pass
            return {
                "slide_dict": slide_to_dict(slide),
                "shapes": shapes,
                "texts": texts,
                "shape_names": set(shapes.keys()),
            }

        info_a = slide_info(sa)
        info_b = slide_info(sb)

        names_a = info_a["shape_names"]
        names_b = info_b["shape_names"]

        differences = []

        # Shapes only in A
        for name in sorted(names_a - names_b):
            differences.append({
                "type": "shape_removed",
                "shape_name": name,
                "detail": f"Shape '{name}' exists in slide {slide_a} but not in slide {slide_b}",
            })

        # Shapes only in B
        for name in sorted(names_b - names_a):
            differences.append({
                "type": "shape_added",
                "shape_name": name,
                "detail": f"Shape '{name}' exists in slide {slide_b} but not in slide {slide_a}",
            })

        # Text differences in common shapes
        for name in sorted(names_a & names_b):
            text_a = info_a["texts"].get(name, "")
            text_b = info_b["texts"].get(name, "")
            if text_a != text_b:
                differences.append({
                    "type": "text_changed",
                    "shape_name": name,
                    "slide_a_text": text_a,
                    "slide_b_text": text_b,
                })

        # Layout difference
        if info_a["slide_dict"]["layout"] != info_b["slide_dict"]["layout"]:
            differences.append({
                "type": "layout_changed",
                "slide_a_layout": info_a["slide_dict"]["layout_name"],
                "slide_b_layout": info_b["slide_dict"]["layout_name"],
            })

        return json.dumps({
            "slide_a": {
                "index": slide_a,
                "shapes_count": len(info_a["shapes"]),
                "layout": info_a["slide_dict"]["layout_name"],
            },
            "slide_b": {
                "index": slide_b,
                "shapes_count": len(info_b["shapes"]),
                "layout": info_b["slide_dict"]["layout_name"],
            },
            "differences": differences,
            "differences_count": len(differences),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 102: snapshot_to_json
# ---------------------------------------------------------------------------

@mcp.tool()
def snapshot_to_json() -> str:
    """Export a full JSON snapshot of the entire presentation.

    Includes presentation metadata, and for each slide: slide info,
    all shapes (shape_to_dict), and slide notes.
    """
    try:
        app = get_app()
        pres = get_pres(app)

        slides_data = []
        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            shapes_list = []
            for j in range(1, slide.Shapes.Count + 1):
                shapes_list.append(shape_to_dict(slide.Shapes(j)))

            notes_text = ""
            try:
                notes_text = slide.NotesPage.Shapes(2).TextFrame.TextRange.Text
            except Exception:
                pass

            slides_data.append({
                "slide": slide_to_dict(slide),
                "shapes": shapes_list,
                "notes": notes_text,
            })

        metadata = {
            "name": _safe_attr(pres, "Name", ""),
            "path": _safe_attr(pres, "FullName", ""),
            "slide_count": pres.Slides.Count,
            "slide_width": round(_points_to_inches(float(_safe_attr(pres.PageSetup, "SlideWidth", 0))), 4),
            "slide_height": round(_points_to_inches(float(_safe_attr(pres.PageSetup, "SlideHeight", 0))), 4),
        }

        return json.dumps({
            "metadata": metadata,
            "slides": slides_data,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 103: get_color_usage
# ---------------------------------------------------------------------------

@mcp.tool()
def get_color_usage() -> str:
    """Analyze all colors used in the presentation.

    Collects colors from shape fills, lines, and text fonts.
    Returns a dictionary mapping hex colors to their usage count and locations.
    """
    try:
        app = get_app()
        pres = get_pres(app)

        def bgr_to_hex(bgr_int):
            """Convert a BGR-packed integer to #RRGGBB hex string."""
            try:
                bgr = int(bgr_int)
                r = bgr & 0xFF
                g = (bgr >> 8) & 0xFF
                b = (bgr >> 16) & 0xFF
                return f"#{r:02X}{g:02X}{b:02X}"
            except Exception:
                return None

        color_map = {}

        def record_color(hex_color, slide_idx, shape_name, prop):
            if hex_color is None:
                return
            if hex_color not in color_map:
                color_map[hex_color] = {"count": 0, "locations": []}
            color_map[hex_color]["count"] += 1
            color_map[hex_color]["locations"].append({
                "slide": slide_idx,
                "shape": shape_name,
                "property": prop,
            })

        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for j in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(j)
                shape_name = _safe_attr(shape, "Name", "")

                # Fill color
                try:
                    if shape.Fill.Visible:
                        hex_c = bgr_to_hex(shape.Fill.ForeColor.RGB)
                        record_color(hex_c, si, shape_name, "fill")
                except Exception:
                    pass

                # Line color
                try:
                    if shape.Line.Visible:
                        hex_c = bgr_to_hex(shape.Line.ForeColor.RGB)
                        record_color(hex_c, si, shape_name, "line")
                except Exception:
                    pass

                # Text font color
                try:
                    if shape.HasTextFrame:
                        tr = shape.TextFrame.TextRange
                        try:
                            hex_c = bgr_to_hex(tr.Font.Color.RGB)
                            record_color(hex_c, si, shape_name, "text")
                        except Exception:
                            pass
                        # Check individual runs for different colors
                        try:
                            for r in range(1, tr.Runs().Count + 1):
                                hex_c = bgr_to_hex(tr.Runs(r).Font.Color.RGB)
                                record_color(hex_c, si, shape_name, f"text_run_{r}")
                        except Exception:
                            pass
                except Exception:
                    pass

        return json.dumps({
            "colors": color_map,
            "unique_color_count": len(color_map),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 104: validate_presentation
# ---------------------------------------------------------------------------

@mcp.tool()
def validate_presentation() -> str:
    """Validate the presentation for common issues.

    Checks for: empty slides, very long text (>1000 chars), oversized images,
    missing titles, and inconsistent fonts (>5 unique). Returns issues,
    warnings, and a quality score (1-100).
    """
    try:
        app = get_app()
        pres = get_pres(app)

        issues = []
        warnings = []
        all_fonts = set()
        slide_count = pres.Slides.Count
        score = 100

        for si in range(1, slide_count + 1):
            slide = pres.Slides(si)
            shape_count = slide.Shapes.Count

            # Check empty slides
            if shape_count == 0:
                issues.append({
                    "slide_index": si,
                    "issue": "Empty slide (no shapes)",
                })
                score -= 5

            # Check for title placeholder (type 1 = title, 15 = title)
            has_title = False
            for j in range(1, shape_count + 1):
                shape = slide.Shapes(j)
                try:
                    ph_type = shape.PlaceholderFormat.Type
                    if ph_type in (1, 15):
                        has_title = True
                except Exception:
                    pass

                type_int = int(_safe_attr(shape, "Type", 0))

                # Check for very long text
                try:
                    if shape.HasTextFrame:
                        text = shape.TextFrame.TextRange.Text
                        if len(text) > 1000:
                            warnings.append({
                                "slide_index": si,
                                "shape_name": _safe_attr(shape, "Name", ""),
                                "warning": f"Very long text ({len(text)} characters)",
                            })
                            score -= 2

                        # Collect fonts
                        font_name = _safe_attr(shape.TextFrame.TextRange.Font, "Name", "")
                        if font_name:
                            all_fonts.add(font_name)
                        try:
                            for r in range(1, shape.TextFrame.TextRange.Runs().Count + 1):
                                fn = _safe_attr(shape.TextFrame.TextRange.Runs(r).Font, "Name", "")
                                if fn:
                                    all_fonts.add(fn)
                        except Exception:
                            pass
                except Exception:
                    pass

                # Check oversized images
                if type_int == 13:
                    try:
                        w = float(_safe_attr(shape, "Width", 0))
                        h = float(_safe_attr(shape, "Height", 0))
                        slide_w = float(pres.PageSetup.SlideWidth)
                        slide_h = float(pres.PageSetup.SlideHeight)
                        if w > slide_w * 1.5 or h > slide_h * 1.5:
                            warnings.append({
                                "slide_index": si,
                                "shape_name": _safe_attr(shape, "Name", ""),
                                "warning": "Oversized image (exceeds 1.5x slide dimensions)",
                            })
                            score -= 3
                    except Exception:
                        pass

            if not has_title and shape_count > 0:
                warnings.append({
                    "slide_index": si,
                    "warning": "Missing title placeholder",
                })
                score -= 2

        # Check font consistency
        if len(all_fonts) > 5:
            issues.append({
                "issue": f"Inconsistent fonts: {len(all_fonts)} unique fonts used (>5)",
                "fonts": sorted(list(all_fonts)),
            })
            score -= 5

        score = max(1, min(100, score))

        return json.dumps({
            "issues": issues,
            "warnings": warnings,
            "score": score,
            "slide_count": slide_count,
            "unique_fonts": len(all_fonts),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ---------------------------------------------------------------------------
# Tool 105: get_text_by_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def get_text_by_slide() -> str:
    """Get all text content organized by slide.

    For each slide, collects text from all shapes that have a text frame.
    Returns a list of slides with their shape texts.
    """
    try:
        app = get_app()
        pres = get_pres(app)

        slides_text = []
        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            texts = []

            for j in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(j)
                try:
                    if shape.HasTextFrame:
                        text = shape.TextFrame.TextRange.Text
                        if text.strip():
                            texts.append({
                                "shape_name": _safe_attr(shape, "Name", ""),
                                "text": text,
                            })
                except Exception:
                    pass

            slides_text.append({
                "slide_index": si,
                "slide_name": _safe_attr(slide, "Name", ""),
                "texts": texts,
            })

        return json.dumps({
            "slides": slides_text,
            "slide_count": len(slides_text),
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    mcp.run()
