"""
PowerPoint MCP Server
Controls local Microsoft PowerPoint via COM automation.
Install: pip install mcp pywin32
Run:     python server.py
"""
import json
import os
import re
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
    "object": 16,
    "vertical_title_text": 9,
    "vertical_text": 10,
    "custom": 13,
    "two_objects": 29,
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
    # Basic shapes
    "rectangle": 1,
    "rounded_rectangle": 5,
    "oval": 9,
    "diamond": 4,
    "triangle": 7,
    "right_triangle": 6,
    "parallelogram": 2,
    "trapezoid": 3,
    "pentagon": 56,
    "hexagon": 10,
    "octagon": 11,
    "cross": 12,
    "cube": 14,
    "can": 13,
    "donut": 18,
    "no_symbol": 19,
    # Arrows
    "right_arrow": 33,
    "left_arrow": 34,
    "up_arrow": 35,
    "down_arrow": 36,
    "left_right_arrow": 37,
    "up_down_arrow": 38,
    "bent_arrow": 41,
    "u_turn_arrow": 42,
    "chevron": 52,
    "notched_right_arrow": 50,
    # Stars and banners
    "star_4": 91,
    "star_5": 92,
    "star_6": 93,
    "star_8": 94,
    "star_16": 95,
    "star_24": 96,
    "star_32": 97,
    "explosion_1": 89,
    "explosion_2": 90,
    "ribbon_up": 97,
    "ribbon_down": 98,
    # Flowchart
    "flowchart_process": 109,
    "flowchart_decision": 110,
    "flowchart_data": 111,
    "flowchart_document": 114,
    "flowchart_terminator": 116,
    "flowchart_preparation": 117,
    "flowchart_manual_input": 118,
    "flowchart_connector": 120,
    # Special
    "heart": 21,
    "lightning": 22,
    "sun": 23,
    "moon": 24,
    "smiley_face": 17,
    "brace_left": 87,
    "brace_right": 88,
    "bracket_left": 85,
    "bracket_right": 86,
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
    "cover": 1284,
    "uncover": 1542,
    "random_bars": 769,
    "blinds": 513,
    "clock": 3849,
    "ripple": 3850,
    "honeycomb": 3851,
    "glitter": 3852,
    "vortex": 3853,
    "shred": 3854,
    "flash": 3855,
    "fly_through": 3861,
}

ANIMATION_MAP = {
    "appear": 1,
    "fade": 10,
    "fly_in": 2,
    "wipe": 22,
    "split": 13,
    "wheel": 21,
    "grow_shrink": 50,
    "bounce": 26,
    "swivel": 15,
    "spiral_in": 55,
    "expand": 50,
    "float_up": 42,
    "float_down": 36,
    "zoom": 53,
    "rise_up": 37,
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
# ERROR TAXONOMY
# ═══════════════════════════════════════════════════════════════════════════════

class PPTError(Exception):
    """Base error for PowerPoint MCP operations."""
    code = "PPT_ERROR"


class ValidationError(PPTError):
    """Invalid input from the caller."""
    code = "VALIDATION_ERROR"


class NotFoundError(PPTError):
    """Requested resource (file, shape, presentation) does not exist."""
    code = "NOT_FOUND"


class BoundsError(PPTError):
    """Index out of valid range."""
    code = "OUT_OF_BOUNDS"


class COMError(PPTError):
    """COM / HRESULT failure from PowerPoint."""
    code = "COM_ERROR"


class ReadOnlyError(PPTError):
    """Attempted mutation on a read-only presentation."""
    code = "READ_ONLY"


# Common HRESULT → human-readable messages
_HRESULT_MAP = {
    -2147352567: "Member not found (method/property does not exist on this object)",
    -2147024894: "File not found",
    -2147024891: "Access denied",
    -2147352565: "Type mismatch",
    -2146827284: "Method or property not supported on this shape type",
    -2147188160: "PowerPoint is busy or modal dialog is open",
    -2004287453: "Invalid parameter or enum value",
    -2147467259: "Unspecified COM error",
    -2147024882: "Out of memory",
    -2147418113: "Call was rejected by callee (COM server busy)",
}


# ═══════════════════════════════════════════════════════════════════════════════
# RESPONSE ENVELOPE & VALIDATION HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _ok(data: dict) -> str:
    """Wrap a successful result dict with status='ok'."""
    if "status" not in data:
        data["status"] = "ok"
    return json.dumps(data, indent=2)


def _ok_list(items: list, label: str = "items") -> str:
    """Wrap a list result in {status, count, <label>}."""
    return json.dumps({
        "status": "ok",
        "count": len(items),
        label: items,
    }, indent=2)


def _err(exc: Exception) -> str:
    """Wrap an exception into a structured error response."""
    code = "UNKNOWN_ERROR"
    message = str(exc)

    if isinstance(exc, PPTError):
        code = exc.code
    elif isinstance(exc, json.JSONDecodeError):
        code = "VALIDATION_ERROR"
        message = f"Invalid JSON: {exc}"
    else:
        # Check for COM HRESULT codes
        hresult = getattr(exc, "hresult", None)
        if hresult is None:
            # Try to extract from pywintypes.com_error args
            args = getattr(exc, "args", ())
            if args and isinstance(args[0], int):
                hresult = args[0]
        if hresult and hresult in _HRESULT_MAP:
            code = "COM_ERROR"
            message = f"{_HRESULT_MAP[hresult]} (HRESULT {hresult}): {exc}"
        elif "com_error" in type(exc).__name__.lower():
            code = "COM_ERROR"

    return json.dumps({"error": message, "code": code}, indent=2)


def _validate_positive_dimensions(**kwargs) -> None:
    """Raise ValidationError if any dimension is <= 0."""
    for name, val in kwargs.items():
        if val is not None and val <= 0:
            raise ValidationError(f"'{name}' must be > 0, got {val}")


def _validate_file_exists(path: str) -> str:
    """Return absolute path or raise NotFoundError."""
    abs_path = os.path.abspath(path)
    if not os.path.isfile(abs_path):
        raise NotFoundError(f"File not found: {abs_path}")
    return abs_path


def _validate_json_list(raw: str, label: str) -> list:
    """Parse JSON and assert it is a list."""
    data = json.loads(raw)
    if not isinstance(data, list):
        raise ValidationError(f"'{label}' must be a JSON array, got {type(data).__name__}")
    return data


def _validate_json_dict(raw: str, label: str, required_keys: tuple = ()) -> dict:
    """Parse JSON and assert it is a dict with required keys."""
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValidationError(f"'{label}' must be a JSON object, got {type(data).__name__}")
    for key in required_keys:
        if key not in data:
            raise ValidationError(f"'{label}' missing required key '{key}'")
    return data


def _validate_table_bounds(table, row: int, col: int) -> None:
    """Raise BoundsError if row/col exceed table dimensions."""
    max_rows = table.Rows.Count
    max_cols = table.Columns.Count
    if row < 1 or row > max_rows:
        raise BoundsError(f"Row {row} out of range (1..{max_rows})")
    if col < 1 or col > max_cols:
        raise BoundsError(f"Column {col} out of range (1..{max_cols})")


def _validate_color(color_str: str) -> None:
    """Validate color format before parsing."""
    if not color_str or not color_str.strip():
        raise ValidationError("Color string cannot be empty")
    s = color_str.strip()
    if s.startswith("#"):
        hex_part = s.lstrip("#")
        if len(hex_part) != 6 or not all(c in "0123456789abcdefABCDEF" for c in hex_part):
            raise ValidationError(f"Invalid hex color: '{color_str}'. Expected '#RRGGBB' with valid hex digits.")
    elif "," in s:
        parts = [p.strip() for p in s.split(",")]
        if len(parts) != 3:
            raise ValidationError(f"Invalid RGB color: '{color_str}'. Expected 'R,G,B' with 3 components.")
        for i, p in enumerate(parts):
            try:
                v = int(p)
            except ValueError:
                raise ValidationError(f"Invalid RGB component '{p}' in color '{color_str}'.")
            if v < 0 or v > 255:
                raise ValidationError(f"RGB component {v} out of range (0..255) in color '{color_str}'.")
    else:
        raise ValidationError(f"Unrecognised color format: '{color_str}'. Use '#RRGGBB' or 'R,G,B'.")


def _validate_url(url: str) -> None:
    """Validate URL has a protocol prefix."""
    if not url or not url.strip():
        raise ValidationError("URL cannot be empty")
    if not re.match(r'^(https?|file)://', url, re.IGNORECASE):
        raise ValidationError(f"URL must start with http://, https://, or file://, got '{url}'")


def _validate_placeholder_index(slide, idx: int) -> None:
    """Raise BoundsError if placeholder index is out of range."""
    count = slide.Shapes.Placeholders.Count
    if idx < 1 or idx > count:
        raise BoundsError(f"Placeholder index {idx} out of range (1..{count})")


def _require_writable(pres) -> None:
    """Raise ReadOnlyError if the presentation is read-only."""
    try:
        if pres.ReadOnly:
            raise ReadOnlyError("Presentation is read-only. Open it with read_only=False to modify.")
    except ReadOnlyError:
        raise
    except Exception:
        pass  # ReadOnly property may not be available; assume writable


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
        raise BoundsError(f"Slide index {slide_index} out of range (1..{count}).")
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
            raise BoundsError(f"Shape index {idx} out of range (1..{count}).")
        return slide.Shapes(idx)
    except (ValueError, TypeError):
        pass

    # Fall back to name search
    name = str(shape_id)
    for i in range(1, slide.Shapes.Count + 1):
        if slide.Shapes(i).Name == name:
            return slide.Shapes(i)

    raise NotFoundError(f"Shape '{shape_id}' not found on slide.")


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
        return _err(e)


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
        return _err(e)


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
                return _err(NotFoundError(f"Template not found: {abs_path}"))
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
        return _err(e)


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
            return _err(NotFoundError(f"File not found: {abs_path}"))

        pres = app.Presentations.Open(abs_path, ReadOnly=read_only)
        return json.dumps({
            "status": "opened",
            "name": pres.Name,
            "path": pres.FullName,
            "slide_count": pres.Slides.Count,
            "read_only": bool(pres.ReadOnly),
        }, indent=2)
    except Exception as e:
        return _err(e)


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
        return _err(e)


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
            return _err(ValidationError(f"Unknown format '{format}'. Supported: {', '.join(SAVE_FORMAT_MAP.keys())}"))

        pres.SaveAs(abs_path, format_id)
        return json.dumps({
            "status": "saved_as",
            "path": abs_path,
            "format": fmt_lower,
        }, indent=2)
    except Exception as e:
        return _err(e)


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
        save_warning = ""
        if save:
            try:
                pres.Save()
            except Exception as save_err:
                save_warning = f"Save failed (file may not have been saved to disk): {save_err}"
        pres.Close()
        result = {"status": "closed", "name": name}
        if save_warning:
            result["warning"] = save_warning
        return json.dumps(result, indent=2)
    except Exception as e:
        return _err(e)


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
        return _ok_list(results, "presentations")
    except Exception as e:
        return _err(e)


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
            return _err(NotFoundError(f"Presentation '{name_or_index}' not found."))

        # Activate by bringing its first window to the front
        pres.Windows(1).Activate()
        return json.dumps({
            "status": "switched",
            "name": pres.Name,
        }, indent=2)
    except Exception as e:
        return _err(e)


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
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 11: set_presentation_properties
# ---------------------------------------------------------------------------

@mcp.tool()
def set_presentation_properties(properties_json: str) -> str:
    """Set built-in document properties (Title, Author, Subject, etc.) from a JSON object string."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        props = _validate_json_dict(properties_json, "properties_json")
        updated = []
        for key, value in props.items():
            try:
                pres.BuiltInDocumentProperties(key).Value = str(value)
                updated.append(key)
            except Exception as prop_err:
                return _err(PPTError(f"Failed to set property '{key}': {prop_err}"))

        return json.dumps({
            "status": "updated",
            "properties_set": updated,
        }, indent=2)
    except json.JSONDecodeError as je:
        return _err(je)
    except Exception as e:
        return _err(e)


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
            return _err(ValidationError("Cannot export — presentation has no slides."))

        abs_path = os.path.abspath(output_path)
        fmt_lower = format.lower()

        format_id = SAVE_FORMAT_MAP.get(fmt_lower)
        if format_id is None:
            return _err(ValidationError(f"Unknown format '{format}'. Supported: {', '.join(SAVE_FORMAT_MAP.keys())}"))

        pres.SaveAs(abs_path, format_id)
        return json.dumps({
            "status": "exported",
            "path": abs_path,
            "format": fmt_lower,
        }, indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 13: set_slide_size
# ---------------------------------------------------------------------------

@mcp.tool()
def set_slide_size(
    width_inches: float = 13.333,
    height_inches: float = 7.5,
    orientation: str = "landscape",
) -> str:
    """Set slide dimensions in inches and orientation ('landscape' or 'portrait'). width/height must be > 0."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        _validate_positive_dimensions(width_inches=width_inches, height_inches=height_inches)
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
        return _err(e)


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

        return _ok_list(masters, "masters")
    except Exception as e:
        return _err(e)


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
        return _ok_list(slides, "slides")
    except Exception as e:
        return _err(e)


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
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 17: add_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def add_slide(layout: str = "blank", index: int = 0) -> str:
    """Add a new slide with a specified layout (see LAYOUT_MAP). index=0 means append at the end."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)

        layout_enum = LAYOUT_MAP.get(layout.lower())
        if layout_enum is None:
            return _err(ValidationError(f"Unknown layout '{layout}'. Supported: {', '.join(LAYOUT_MAP.keys())}"))

        insert_idx = index if index > 0 else pres.Slides.Count + 1
        new_slide = pres.Slides.Add(insert_idx, layout_enum)

        return json.dumps(slide_to_dict(new_slide), indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 18: duplicate_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def duplicate_slide(slide_index: int) -> str:
    """Duplicate a slide. The copy is inserted immediately after the original."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        dup_range = slide.Duplicate()
        new_slide = dup_range(1)
        return json.dumps(slide_to_dict(new_slide), indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 19: delete_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def delete_slide(slide_index: int) -> str:
    """Delete a slide by 1-based index."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        slide.Delete()
        return json.dumps({
            "status": "deleted",
            "slide_index": slide_index,
        }, indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 20: move_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def move_slide(slide_index: int, new_index: int) -> str:
    """Move a slide from one position to another (1-based indices)."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        get_slide(pres, slide_index)  # validate index
        pres.Slides(slide_index).MoveTo(new_index)
        return json.dumps({
            "status": "moved",
            "from": slide_index,
            "to": new_index,
        }, indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 21: copy_slide
# ---------------------------------------------------------------------------

@mcp.tool()
def copy_slide(source_index: int, target_pres_name: str = "", target_index: int = 0) -> str:
    """Copy a slide within the same presentation or to another open presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
                return _err(NotFoundError(f"Target presentation '{target_pres_name}' not found."))

            insert_at = target_index if target_index > 0 else target_pres.Slides.Count + 1
            target_pres.Slides.InsertFromFile(pres.FullName, insert_at - 1, source_index, source_index)
            return json.dumps({
                "status": "copied",
                "target_presentation": target_pres.Name,
                "new_slide_index": insert_at,
            }, indent=2)
    except Exception as e:
        return _err(e)


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
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 23: set_slide_notes
# ---------------------------------------------------------------------------

@mcp.tool()
def set_slide_notes(slide_index: int, notes: str) -> str:
    """Set the speaker notes for a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes
        return json.dumps({
            "status": "updated",
            "slide_index": slide_index,
        }, indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 24: set_slide_layout
# ---------------------------------------------------------------------------

@mcp.tool()
def set_slide_layout(slide_index: int, layout: str) -> str:
    """Change the layout of an existing slide. layout must be a key from LAYOUT_MAP."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)

        layout_enum = LAYOUT_MAP.get(layout.lower())
        if layout_enum is None:
            return _err(ValidationError(f"Unknown layout '{layout}'. Supported: {', '.join(LAYOUT_MAP.keys())}"))

        slide.Layout = layout_enum
        return json.dumps({
            "status": "updated",
            "slide_index": slide_index,
            "layout": layout,
        }, indent=2)
    except Exception as e:
        return _err(e)


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
    """Set the transition effect for a slide. effect must be a key from TRANSITION_MAP. duration in seconds."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 26: bulk_add_slides
# ---------------------------------------------------------------------------

@mcp.tool()
def bulk_add_slides(slides_json: str) -> str:
    """Add multiple slides from a JSON array of {layout, index} objects."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slides_spec = _validate_json_list(slides_json, "slides_json")
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

        return _ok_list(results, "slides")
    except json.JSONDecodeError as je:
        return _err(je)
    except Exception as e:
        return _err(e)


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
        _require_writable(pres)
        new_order = _validate_json_list(order_json, "order_json")

        # Validate all indices
        count = pres.Slides.Count
        for idx in new_order:
            if not isinstance(idx, int) or idx < 1 or idx > count:
                raise BoundsError(f"Slide index {idx} out of range (1..{count}).")

        # Capture SlideIDs in the desired order BEFORE any moves
        desired_ids = [pres.Slides(idx).SlideID for idx in new_order]

        # Move each slide to its target position, finding by SlideID
        for target_pos, sid in enumerate(desired_ids, start=1):
            # Find current position of slide with this SlideID
            for i in range(1, pres.Slides.Count + 1):
                if pres.Slides(i).SlideID == sid:
                    if i != target_pos:
                        pres.Slides(i).MoveTo(target_pos)
                    break

        return _ok({
            "status": "reordered",
            "new_order": new_order,
        })
    except (PPTError, json.JSONDecodeError) as e:
        return _err(e)
    except Exception as e:
        return _err(e)


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
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 29: set_slide_background
# ---------------------------------------------------------------------------

@mcp.tool()
def set_slide_background(slide_index: int, color: str = "", image_path: str = "") -> str:
    """Set the background of a slide to a solid color ('#RRGGBB' or 'R,G,B') or an image file path."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        if color:
            _validate_color(color)
        slide = get_slide(pres, slide_index)

        slide.FollowMasterBackground = 0  # msoFalse

        if color:
            slide.Background.Fill.Solid()
            slide.Background.Fill.ForeColor.RGB = _parse_color(color)
        elif image_path:
            abs_path = os.path.abspath(image_path)
            if not os.path.isfile(abs_path):
                return _err(NotFoundError(f"Image not found: {abs_path}"))
            slide.Background.Fill.UserPicture(abs_path)
        else:
            return _err(ValidationError("Provide either 'color' or 'image_path'."))

        return json.dumps({
            "status": "updated",
            "slide_index": slide_index,
        }, indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 30: bulk_set_transitions
# ---------------------------------------------------------------------------

@mcp.tool()
def bulk_set_transitions(settings_json: str) -> str:
    """Apply transitions to multiple slides from a JSON array of {slide_index, effect, duration, advance_time} objects."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        settings = _validate_json_list(settings_json, "settings_json")
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

        return _ok_list(results, "transitions")
    except (PPTError, json.JSONDecodeError) as e:
        return _err(e)
    except Exception as e:
        return _err(e)


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
        return _ok_list(shapes, "shapes")
    except Exception as e:
        return _err(e)


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
        return _err(e)


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
    """Add a text box to a slide with optional font formatting and alignment. Positions/sizes in inches; width/height must be > 0."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        _validate_positive_dimensions(width=width, height=height)
        if font_color:
            _validate_color(font_color)
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
        return _err(e)


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
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        changed = []
        tr = shape.TextFrame.TextRange
        if text:
            tr.Text = text
            changed.append("text")
        if font_size > 0:
            tr.Font.Size = font_size
            changed.append("font_size")
        if font_name:
            tr.Font.Name = font_name
            changed.append("font_name")
        if font_color:
            _validate_color(font_color)
            tr.Font.Color.RGB = _parse_color(font_color)
            changed.append("font_color")
        if bold != -1:
            tr.Font.Bold = -1 if bold == 1 else 0
            changed.append("bold")
        if italic != -1:
            tr.Font.Italic = -1 if italic == 1 else 0
            changed.append("italic")
        if alignment:
            align_map = {"left": 1, "center": 2, "right": 3, "justify": 4}
            align_val = align_map.get(alignment.lower())
            if align_val is not None:
                tr.ParagraphFormat.Alignment = align_val
                changed.append("alignment")

        result = shape_to_dict(shape)
        result["changed"] = changed
        return json.dumps(result, indent=2)
    except Exception as e:
        return _err(e)


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
    """Add an auto-shape to a slide. shape_type must be a key from AUTOSHAPE_MAP. width/height must be > 0 (inches)."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        _validate_positive_dimensions(width=width, height=height)
        if fill_color:
            _validate_color(fill_color)
        if line_color:
            _validate_color(line_color)
        slide = get_slide(pres, slide_index)

        type_id = AUTOSHAPE_MAP.get(shape_type.lower())
        if type_id is None:
            return _err(ValidationError(f"Unknown shape_type '{shape_type}'. Supported: {', '.join(AUTOSHAPE_MAP.keys())}"))

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
        return _err(e)


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
        _require_writable(pres)
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
        return _err(e)


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
    """Update shape properties (inches). Only non-default values are applied. width/height must be > 0 if provided."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        changed = []
        if left >= 0:
            shape.Left = _inches_to_points(left)
            changed.append("left")
        if top >= 0:
            shape.Top = _inches_to_points(top)
            changed.append("top")
        if width >= 0:
            _validate_positive_dimensions(width=width)
            shape.Width = _inches_to_points(width)
            changed.append("width")
        if height >= 0:
            _validate_positive_dimensions(height=height)
            shape.Height = _inches_to_points(height)
            changed.append("height")
        if rotation >= 0:
            shape.Rotation = rotation
            changed.append("rotation")
        if fill_color:
            _validate_color(fill_color)
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = _parse_color(fill_color)
            changed.append("fill_color")
        if line_color:
            _validate_color(line_color)
            shape.Line.ForeColor.RGB = _parse_color(line_color)
            changed.append("line_color")
        if name:
            shape.Name = name
            changed.append("name")

        result = shape_to_dict(shape)
        result["changed"] = changed
        return json.dumps(result, indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 38: delete_shape
# ---------------------------------------------------------------------------

@mcp.tool()
def delete_shape(slide_index: int, shape_name: str) -> str:
    """Delete a shape from a slide by name or index."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)
        shape.Delete()
        return json.dumps({
            "status": "deleted",
            "shape_name": shape_name,
        }, indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 39: group_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def group_shapes(slide_index: int, shape_names_json: str, group_name: str = "") -> str:
    """Group multiple shapes together. shape_names_json is a JSON array of shape names."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        names = _validate_json_list(shape_names_json, "shape_names_json")

        import win32com.client
        names_array = win32com.client.VARIANT(0x2008, names)  # VT_ARRAY | VT_BSTR
        group = slide.Shapes.Range(names_array).Group()

        if group_name:
            group.Name = group_name

        return json.dumps(shape_to_dict(group), indent=2)
    except json.JSONDecodeError as je:
        return _err(je)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 40: ungroup_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def ungroup_shapes(slide_index: int, group_name: str) -> str:
    """Ungroup a grouped shape into its individual shapes."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 41: add_hyperlink
# ---------------------------------------------------------------------------

@mcp.tool()
def add_hyperlink(slide_index: int, shape_name: str, url: str, tooltip: str = "") -> str:
    """Add a hyperlink to a shape. URL must start with http:// or https://."""
    try:
        _validate_url(url)
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


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
        _require_writable(pres)
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
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 43: align_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def align_shapes(slide_index: int, shape_names_json: str, alignment: str = "center") -> str:
    """Align multiple shapes. Alignment: left, center, right, top, middle, bottom."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        names = _validate_json_list(shape_names_json, "shape_names_json")

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
        return _err(je)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 44: distribute_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def distribute_shapes(slide_index: int, shape_names_json: str, direction: str = "horizontal") -> str:
    """Distribute shapes evenly. Direction: horizontal or vertical."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        names = _validate_json_list(shape_names_json, "shape_names_json")

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
        return _err(je)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 45: duplicate_shape
# ---------------------------------------------------------------------------

@mcp.tool()
def duplicate_shape(slide_index: int, shape_name: str, offset_x: float = 0.5, offset_y: float = 0.5) -> str:
    """Duplicate a shape and offset the copy by the given inches."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        new_shape = shape.Duplicate()
        new_shape.Left = shape.Left + _inches_to_points(offset_x)
        new_shape.Top = shape.Top + _inches_to_points(offset_y)

        return json.dumps(shape_to_dict(new_shape), indent=2)
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------------------
# Tool 46: set_shape_z_order
# ---------------------------------------------------------------------------

@mcp.tool()
def set_shape_z_order(slide_index: int, shape_name: str, action: str = "front") -> str:
    """Change the z-order of a shape. Actions: front, back, forward, backward."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


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

        return _ok_list(results, "shapes")
    except json.JSONDecodeError as je:
        return _err(je)
    except Exception as e:
        return _err(e)


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
    """Set text and formatting on a slide placeholder by 1-based index."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        _validate_placeholder_index(slide, placeholder_index)

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
        return _err(e)


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
    """Insert an image from a local file path onto a slide. Positions in inches."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        abs_path = _validate_file_exists(image_path)
        slide = get_slide(pres, slide_index)
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
        return _err(e)


@mcp.tool()
def insert_image_from_url(
    slide_index: int,
    url: str,
    left: float,
    top: float,
    width: float = 0,
    height: float = 0,
) -> str:
    """Download an image from a URL and insert it onto a slide. Positions in inches."""
    try:
        import urllib.request
        import tempfile

        _validate_url(url)
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


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
    """Add a table to a slide. rows >= 1, cols >= 1. Optionally fill with data from a 2D JSON array. Positions in inches."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        _validate_positive_dimensions(width=width, height=height)
        if rows < 1:
            raise ValidationError(f"'rows' must be >= 1, got {rows}")
        if cols < 1:
            raise ValidationError(f"'cols' must be >= 1, got {cols}")
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
        return _err(e)


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
    """Modify the text and formatting of a single table cell (1-based row/col)."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        table = shape.Table
        _validate_table_bounds(table, row, col)
        cell = table.Cell(row, col)
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
        return _err(e)


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
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, shape_name)

        data = _validate_json_list(data_json, "data_json")
        table = shape.Table
        max_rows = table.Rows.Count
        max_cols = table.Columns.Count
        rows_filled = 0
        cols_filled = 0

        for r_idx, row_data in enumerate(data):
            if r_idx + 1 > max_rows:
                break
            rows_filled = max(rows_filled, r_idx + 1)
            for c_idx, cell_val in enumerate(row_data):
                if c_idx + 1 > max_cols:
                    break
                cols_filled = max(cols_filled, c_idx + 1)
                table.Cell(r_idx + 1, c_idx + 1).Shape.TextFrame.TextRange.Text = str(cell_val)

        return json.dumps({
            "status": "filled",
            "rows": rows_filled,
            "cols": cols_filled,
        }, indent=2)
    except (PPTError, json.JSONDecodeError) as e:
        return _err(e)
    except Exception as e:
        return _err(e)


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
    """Format a table with banding, header styling, and alternating row colors. Colors as '#RRGGBB' or 'R,G,B'."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        if header_color:
            _validate_color(header_color)
        if row_color1:
            _validate_color(row_color1)
        if row_color2:
            _validate_color(row_color2)
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
        return _err(e)


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
        try:
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
            try:
                ws.Range(f"A1:{last_col_letter}{total_rows}").Select()
            except Exception:
                pass  # Range selection not critical — chart auto-detects data
        finally:
            try:
                wb.Close(True)
            except Exception:
                pass

        if title:
            chart.HasTitle = -1  # msoTrue
            chart.ChartTitle.Text = title

        return json.dumps(shape_to_dict(shape), indent=2)
    except Exception as e:
        return _err(e)


@mcp.tool()
def modify_chart(
    slide_index: int,
    shape_name: str,
    title: str = "",
    has_legend: bool = True,
    legend_position: str = "",
) -> str:
    """Modify chart title and legend properties. legend_position: bottom, top, left, right."""
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
        return _err(e)


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
        try:
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
        finally:
            try:
                wb.Close(True)
            except Exception:
                pass

        return json.dumps({
            "status": "updated",
            "shape_name": shape.Name,
        }, indent=2)
    except Exception as e:
        return _err(e)


@mcp.tool()
def insert_video(
    slide_index: int,
    video_path: str,
    left: float,
    top: float,
    width: float,
    height: float,
) -> str:
    """Insert a video file onto a slide. Positions in inches; width/height must be > 0."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        _validate_positive_dimensions(width=width, height=height)
        abs_path = _validate_file_exists(video_path)
        slide = get_slide(pres, slide_index)
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
        return _err(e)


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
        _require_writable(pres)
        abs_path = _validate_file_exists(audio_path)
        slide = get_slide(pres, slide_index)
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
        return _err(e)


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
    """Insert an OLE embedded object (e.g., Excel, Word, PDF) onto a slide. Positions in inches; width/height must be > 0."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        _validate_positive_dimensions(width=width, height=height)
        abs_path = _validate_file_exists(file_path)
        slide = get_slide(pres, slide_index)
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
        return _err(e)


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
        _require_writable(pres)
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
        return _err(e)


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
        _require_writable(pres)
        _validate_file_exists(new_image_path)
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
        return _err(e)


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
        return _err(e)


@mcp.tool()
def apply_theme(theme_path: str) -> str:
    """Apply a .thmx theme file to the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        abs_path = _validate_file_exists(theme_path)
        pres.ApplyTheme(abs_path)
        return json.dumps({
            "status": "applied",
            "theme_path": abs_path,
        }, indent=2)
    except Exception as e:
        return _err(e)


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
        return _err(e)


@mcp.tool()
def set_theme_color(slot: str, color: str) -> str:
    """Set a theme color slot by name (e.g. 'Accent1', 'Text1') to a new color ('#RRGGBB' or 'R,G,B'). Valid slots: background1, text1, background2, text2, accent1-6, hyperlink, followedhyperlink."""
    try:
        _validate_color(color)
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)

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
        return _err(e)


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
        return _err(e)


@mcp.tool()
def set_theme_fonts(major_font: str = "", minor_font: str = "") -> str:
    """Set the major (headings) and/or minor (body) theme fonts."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


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
        return _err(e)


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
    """Modify font formatting of a placeholder in a master layout. font_size in points."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


@mcp.tool()
def set_background(
    slide_index: int = 0,
    color: str = "",
    image_path: str = "",
    gradient_color1: str = "",
    gradient_color2: str = "",
) -> str:
    """Set background of a slide (by 1-based index) or the slide master (index=0). Supports solid color ('#RRGGBB'), image, or two-color gradient."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        if color:
            _validate_color(color)
        if gradient_color1:
            _validate_color(gradient_color1)
        if gradient_color2:
            _validate_color(gradient_color2)

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
        return _err(e)


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
        return _ok_list(placeholders, "placeholders")
    except Exception as e:
        return _err(e)


@mcp.tool()
def add_custom_layout(master_index: int = 1, name: str = "Custom Layout") -> str:
    """Add a new custom layout to a slide master."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


@mcp.tool()
def copy_master_from(source_path: str) -> str:
    """Copy the first slide master/design from another presentation file into the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        abs_path = _validate_file_exists(source_path)

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
        return _err(e)


# ═══════════════════════════════════════════════════════════════════════════════
# PHASE 6 — Advanced Operations (18 tools)
# ═══════════════════════════════════════════════════════════════════════════════


@mcp.tool()
def find_and_replace(find_text: str, replace_text: str, match_case: bool = False) -> str:
    """Find and replace text across all slides in the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


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

        return _ok_list(results, "slides")
    except Exception as e:
        return _err(e)


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
        return _err(e)


@mcp.tool()
def merge_presentations(file_paths_json: str, insert_at: int = 0) -> str:
    """Merge slides from multiple presentation files into the active presentation.

    file_paths_json: JSON array of file paths, e.g. '["C:/a.pptx","C:/b.pptx"]'.
    insert_at: slide index after which to insert (0 = end of presentation).
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        file_paths = _validate_json_list(file_paths_json, "file_paths_json")
        total_inserted = 0
        insert_position = insert_at if insert_at > 0 else pres.Slides.Count

        for fp in file_paths:
            abs_path = _validate_file_exists(fp)
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
        return _err(e)


@mcp.tool()
def apply_template(template_path: str) -> str:
    """Apply a PowerPoint template (.potx/.pptx) to the active presentation."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        abs_path = _validate_file_exists(template_path)
        pres.ApplyTemplate(abs_path)
        return json.dumps({
            "status": "applied",
            "template": abs_path,
        }, indent=2)
    except Exception as e:
        return _err(e)


@mcp.tool()
def bulk_format_text(criteria_json: str) -> str:
    """Apply formatting to all matching text across the presentation.

    criteria_json: JSON object with {find_text, font_size, font_name, font_color, bold, italic}.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        criteria = _validate_json_dict(criteria_json, "criteria_json")
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
        return _err(e)


@mcp.tool()
def add_animation(slide_index: int, shape_name: str, effect: str = "appear",
                  duration: float = 0.5, delay: float = 0, trigger: str = "on_click") -> str:
    """Add an animation effect to a shape on a slide.

    effect: one of appear, fade, fly_in, wipe, split, wheel, grow_shrink, bounce, swivel, spiral_in, expand, float_up, float_down, zoom, rise_up.
    trigger: on_click, with_previous, after_previous.
    duration/delay in seconds.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


@mcp.tool()
def remove_animation(slide_index: int, shape_name: str) -> str:
    """Remove all animation effects from a specific shape on a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


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
        return _err(e)


@mcp.tool()
def reorder_animations(slide_index: int, order_json: str) -> str:
    """Reorder animation effects on a slide by specifying shape names in desired order.

    order_json: JSON array of shape names, e.g. '["Title","Subtitle","Image1"]'.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        desired_order = _validate_json_list(order_json, "order_json")
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
        return _err(e)


@mcp.tool()
def bulk_speaker_notes(notes_json: str) -> str:
    """Set speaker notes on multiple slides at once.

    notes_json: JSON array of {slide_index, notes}, e.g. '[{"slide_index":1,"notes":"Hello"}]'.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        items = _validate_json_list(notes_json, "notes_json")
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
        return _err(e)


@mcp.tool()
def clone_formatting(slide_index: int, source_shape: str, target_shapes_json: str) -> str:
    """Clone formatting from a source shape to multiple target shapes on the same slide.

    target_shapes_json: JSON array of shape names, e.g. '["Shape2","Shape3"]'.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


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
        return _err(e)


@mcp.tool()
def rename_shape(slide_index: int, old_name: str, new_name: str) -> str:
    """Rename a shape on a slide."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide = get_slide(pres, slide_index)
        shape = get_shape(slide, old_name)
        shape.Name = new_name
        return json.dumps({
            "status": "renamed",
            "old_name": old_name,
            "new_name": new_name,
        }, indent=2)
    except Exception as e:
        return _err(e)


@mcp.tool()
def lock_shape(slide_index: int, shape_name: str, lock: bool = True) -> str:
    """Lock or unlock a shape's aspect ratio and available lock properties."""
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
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
        return _err(e)


@mcp.tool()
def add_section(name: str, before_slide: int = 0) -> str:
    """Add a named section to the presentation.

    before_slide: 1-based slide index. 0 means add at the end.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        slide_idx = before_slide if before_slide > 0 else pres.Slides.Count + 1
        section_index = pres.SectionProperties.AddSection(slide_idx, name)
        return json.dumps({
            "status": "added",
            "name": name,
            "section_index": section_index,
        }, indent=2)
    except Exception as e:
        return _err(e)


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
        return _err(e)


@mcp.tool()
def delete_section(section_index: int, delete_slides: bool = False) -> str:
    """Delete a section from the presentation.

    section_index: 1-based section index.
    delete_slides: if True, also delete the slides in that section.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        pres.SectionProperties.Delete(section_index, delete_slides)
        return json.dumps({
            "status": "deleted",
            "section_index": section_index,
            "slides_deleted": delete_slides,
        }, indent=2)
    except Exception as e:
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


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
        return _err(e)


# ═══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════
# DESIGN SYSTEM — Grid Engine, Typography, Palettes, Asset Catalog
# ═══════════════════════════════════════════════════════════════════════════════

# Slide dimensions (16:9 widescreen)
_SLIDE_W = 10.0
_SLIDE_H = 5.625
_MARGIN = 0.5
_USABLE_W = _SLIDE_W - 2 * _MARGIN  # 9.0
_USABLE_H = _SLIDE_H - 2 * _MARGIN  # 4.625
_GRID_COLS = 12
_GUTTER = 0.15
_COL_W = (_USABLE_W - (_GRID_COLS - 1) * _GUTTER) / _GRID_COLS  # ~0.6125

# Footer constants
_FOOTER_H = 0.375
_FOOTER_TOP = _SLIDE_H - _FOOTER_H  # 5.25

# Typography scale (points)
TYPO = {
    "hero": 36,
    "heading": 28,
    "subheading": 14,
    "section_label": 10,
    "stat": 26,
    "body": 11,
    "caption": 9,
    "footer": 7.5,
}

# Color palettes — each defines roles
PALETTES = {
    "dark_executive": {
        "bg": "#060F1A",
        "surface": "#0D1F30",
        "surface_alt": "#0A1525",
        "primary": "#0A7E8C",
        "secondary": "#00BCD4",
        "accent": "#F57C00",
        "accent2": "#E91E63",
        "text": "#FFFFFF",
        "text_dim": "#8899AA",
        "text_muted": "#556677",
        "border": "#1A3A50",
        "highlight_bg": "#0A7E8C",
        "highlight_text": "#FFFFFF",
        "footer_bg": "#040C14",
        "footer_line": "#0A7E8C",
        "footer_brand": "#0A7E8C",
        "footer_text": "#3A5060",
        "top_bar": "#00BCD4",
        "top_bar2": "#0A7E8C",
    },
    "midnight_blue": {
        "bg": "#0A0E1A",
        "surface": "#141B2D",
        "surface_alt": "#0F1422",
        "primary": "#4A6CF7",
        "secondary": "#7B93FF",
        "accent": "#FF6B6B",
        "accent2": "#FFD93D",
        "text": "#FFFFFF",
        "text_dim": "#8892A4",
        "text_muted": "#4A5568",
        "border": "#1E2A3A",
        "highlight_bg": "#4A6CF7",
        "highlight_text": "#FFFFFF",
        "footer_bg": "#060810",
        "footer_line": "#4A6CF7",
        "footer_brand": "#4A6CF7",
        "footer_text": "#3A4560",
        "top_bar": "#7B93FF",
        "top_bar2": "#4A6CF7",
    },
    "light_corporate": {
        "bg": "#FFFFFF",
        "surface": "#F5F7FA",
        "surface_alt": "#EDF0F5",
        "primary": "#0A7E8C",
        "secondary": "#004D61",
        "accent": "#F57C00",
        "accent2": "#E91E63",
        "text": "#1A2332",
        "text_dim": "#5A6B7D",
        "text_muted": "#8899AA",
        "border": "#D0D8E0",
        "highlight_bg": "#0A7E8C",
        "highlight_text": "#FFFFFF",
        "footer_bg": "#F0F2F5",
        "footer_line": "#D0D8E0",
        "footer_brand": "#0A7E8C",
        "footer_text": "#8899AA",
        "top_bar": "#0A7E8C",
        "top_bar2": "#004D61",
    },
}


def _grid_pos(col_start: int, col_span: int, row_top: float, row_height: float) -> dict:
    """Calculate position from grid coordinates.

    col_start: 1-based column (1..12)
    col_span: number of columns to span
    row_top: top position in inches from slide top
    row_height: height in inches
    Returns: {left, top, width, height} in inches
    """
    left = _MARGIN + (col_start - 1) * (_COL_W + _GUTTER)
    width = col_span * _COL_W + (col_span - 1) * _GUTTER
    return {"left": round(left, 4), "top": row_top, "width": round(width, 4), "height": row_height}


def _card_positions(count: int, top: float, height: float, margin: float = 0.5, gap: float = 0.2) -> list:
    """Calculate evenly-spaced card positions across the usable width.

    Returns list of {left, top, width, height} dicts.
    """
    total_gap = (count - 1) * gap
    card_w = (_USABLE_W - total_gap) / count
    positions = []
    for i in range(count):
        left = margin + i * (card_w + gap)
        positions.append({"left": round(left, 4), "top": top, "width": round(card_w, 4), "height": height})
    return positions


def _get_palette(name: str) -> dict:
    """Get a palette by name, defaulting to dark_executive."""
    return PALETTES.get(name, PALETTES["dark_executive"])


# Asset catalog scanner
_ASSET_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")


def _scan_assets() -> dict:
    """Scan the assets directory and return categorized file paths."""
    catalog = {"icons": {}, "backgrounds": {}, "infographics": {}}
    if not os.path.isdir(_ASSET_DIR):
        return catalog
    for root, _dirs, files in os.walk(_ASSET_DIR):
        rel = os.path.relpath(root, _ASSET_DIR).lower()
        for f in files:
            if not f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".svg")):
                continue
            full = os.path.join(root, f)
            name = os.path.splitext(f)[0].lower()
            if "icon" in rel:
                catalog["icons"][name] = full
            elif "background" in rel or "bg" in rel:
                catalog["backgrounds"][name] = full
            elif "infographic" in rel or "info" in rel:
                catalog["infographics"][name] = full
            else:
                catalog["icons"][name] = full  # default to icons
    return catalog


# ═══════════════════════════════════════════════════════════════════════════════
# COMPOUND SLIDE BUILDER HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _build_top_bar(slide, pal: dict):
    """Add the thin accent bar at the very top of the slide."""
    slide.Shapes.AddShape(1, 0, 0, _inches_to_points(_SLIDE_W), _inches_to_points(0.035))
    bar1 = slide.Shapes(slide.Shapes.Count)
    bar1.Fill.Solid()
    bar1.Fill.ForeColor.RGB = _parse_color(pal["top_bar"])
    bar1.Line.Visible = 0

    slide.Shapes.AddShape(1, 0, _inches_to_points(0.035), _inches_to_points(_SLIDE_W), _inches_to_points(0.015))
    bar2 = slide.Shapes(slide.Shapes.Count)
    bar2.Fill.Solid()
    bar2.Fill.ForeColor.RGB = _parse_color(pal["top_bar2"])
    bar2.Line.Visible = 0


def _build_badge(slide, text: str, pal: dict):
    """Add the section badge (e.g., 'SITUATION | 02') top-right."""
    l = _inches_to_points(8.05)
    t = _inches_to_points(0.18)
    w = _inches_to_points(1.65)
    h = _inches_to_points(0.4)
    shape = slide.Shapes.AddShape(5, l, t, w, h)  # rounded_rectangle
    shape.Fill.Solid()
    shape.Fill.ForeColor.RGB = _parse_color(pal["primary"])
    shape.Line.Visible = 0
    tr = shape.TextFrame.TextRange
    tr.Text = text
    tr.Font.Size = TYPO["section_label"]
    tr.Font.Name = "Segoe UI"
    tr.Font.Bold = -1
    tr.Font.Color.RGB = _parse_color("#FFFFFF")
    tr.ParagraphFormat.Alignment = 2  # center


def _build_footer(slide, pal: dict, brand: str = "Deloitte.", meta: str = "", page: str = ""):
    """Add the branded footer bar at the bottom."""
    # Footer background
    ft = slide.Shapes.AddShape(1, 0, _inches_to_points(_FOOTER_TOP), _inches_to_points(_SLIDE_W), _inches_to_points(_FOOTER_H))
    ft.Fill.Solid()
    ft.Fill.ForeColor.RGB = _parse_color(pal["footer_bg"])
    ft.Line.Visible = 0

    # Top line
    line = slide.Shapes.AddLine(0, _inches_to_points(_FOOTER_TOP), _inches_to_points(_SLIDE_W), _inches_to_points(_FOOTER_TOP))
    line.Line.ForeColor.RGB = _parse_color(pal["footer_line"])
    line.Line.Weight = 0.5

    # Brand name
    tb = slide.Shapes.AddTextbox(1, _inches_to_points(0.4), _inches_to_points(_FOOTER_TOP + 0.05), _inches_to_points(1.2), _inches_to_points(0.25))
    tb.TextFrame.TextRange.Text = brand
    tb.TextFrame.TextRange.Font.Size = 10
    tb.TextFrame.TextRange.Font.Name = "Segoe UI"
    tb.TextFrame.TextRange.Font.Bold = -1
    tb.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["footer_brand"])

    # Meta text
    if meta:
        tb2 = slide.Shapes.AddTextbox(1, _inches_to_points(2.0), _inches_to_points(_FOOTER_TOP + 0.06), _inches_to_points(5.5), _inches_to_points(0.22))
        tb2.TextFrame.TextRange.Text = meta
        tb2.TextFrame.TextRange.Font.Size = TYPO["footer"]
        tb2.TextFrame.TextRange.Font.Name = "Segoe UI"
        tb2.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["footer_text"])

    # Page number
    if page:
        tb3 = slide.Shapes.AddTextbox(1, _inches_to_points(8.5), _inches_to_points(_FOOTER_TOP + 0.05), _inches_to_points(1.2), _inches_to_points(0.22))
        tb3.TextFrame.TextRange.Text = page
        tb3.TextFrame.TextRange.Font.Size = TYPO["caption"]
        tb3.TextFrame.TextRange.Font.Name = "Segoe UI"
        tb3.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["footer_text"])
        tb3.TextFrame.TextRange.ParagraphFormat.Alignment = 3  # right


def _build_title_block(slide, title: str, subtitle: str, pal: dict):
    """Add title + subtitle + divider line."""
    # Title
    tb = slide.Shapes.AddTextbox(1, _inches_to_points(0.55), _inches_to_points(0.15), _inches_to_points(7.0), _inches_to_points(0.6))
    tb.TextFrame.TextRange.Text = title
    tb.TextFrame.TextRange.Font.Size = TYPO["heading"]
    tb.TextFrame.TextRange.Font.Name = "Segoe UI"
    tb.TextFrame.TextRange.Font.Bold = -1
    tb.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["text"])

    # Subtitle
    tb2 = slide.Shapes.AddTextbox(1, _inches_to_points(0.55), _inches_to_points(0.78), _inches_to_points(7.0), _inches_to_points(0.3))
    tb2.TextFrame.TextRange.Text = subtitle
    tb2.TextFrame.TextRange.Font.Size = TYPO["subheading"]
    tb2.TextFrame.TextRange.Font.Name = "Segoe UI"
    tb2.TextFrame.TextRange.Font.Italic = -1
    tb2.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["secondary"])

    # Divider line (spans subtitle width)
    slide.Shapes.AddLine(
        _inches_to_points(0.55), _inches_to_points(1.12),
        _inches_to_points(7.55), _inches_to_points(1.12),
    ).Line.ForeColor.RGB = _parse_color(pal["primary"])


def _build_kpi_card(slide, pos: dict, number: str, label: str, pal: dict, is_highlight: bool = False, dot_color: str = ""):
    """Build a single KPI card at the given position."""
    bg = pal["highlight_bg"] if is_highlight else pal["surface"]
    border = pal["secondary"] if is_highlight else pal["primary"]
    num_color = pal["highlight_text"] if is_highlight else pal["secondary"]
    lbl_color = pal["highlight_text"] if is_highlight else pal["text_dim"]

    l, t, w, h = pos["left"], pos["top"], pos["width"], pos["height"]

    # Card background
    shape = slide.Shapes.AddShape(5, _inches_to_points(l), _inches_to_points(t), _inches_to_points(w), _inches_to_points(h))
    shape.Fill.Solid()
    shape.Fill.ForeColor.RGB = _parse_color(bg)
    shape.Line.ForeColor.RGB = _parse_color(border)
    shape.Line.Weight = 1

    # Indicator dot
    if dot_color:
        dot = slide.Shapes.AddShape(9, _inches_to_points(l + 0.15), _inches_to_points(t + 0.12), _inches_to_points(0.12), _inches_to_points(0.12))
        dot.Fill.Solid()
        dot.Fill.ForeColor.RGB = _parse_color(dot_color)
        dot.Line.Visible = 0

    # Big number — auto-shrink if text is wide (>7 chars)
    stat_size = TYPO["stat"] if len(number) <= 7 else TYPO["stat"] - 4
    tb = slide.Shapes.AddTextbox(1, _inches_to_points(l + 0.1), _inches_to_points(t + 0.28), _inches_to_points(w - 0.2), _inches_to_points(0.5))
    tb.TextFrame.TextRange.Text = number
    tb.TextFrame.TextRange.Font.Size = stat_size
    tb.TextFrame.TextRange.Font.Name = "Segoe UI"
    tb.TextFrame.TextRange.Font.Bold = -1
    tb.TextFrame.TextRange.Font.Color.RGB = _parse_color(num_color)
    tb.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # center

    # Label
    tb2 = slide.Shapes.AddTextbox(1, _inches_to_points(l + 0.1), _inches_to_points(t + 0.78), _inches_to_points(w - 0.2), _inches_to_points(0.3))
    tb2.TextFrame.TextRange.Text = label
    tb2.TextFrame.TextRange.Font.Size = TYPO["caption"]
    tb2.TextFrame.TextRange.Font.Name = "Segoe UI"
    tb2.TextFrame.TextRange.Font.Color.RGB = _parse_color(lbl_color)
    tb2.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # center


def _build_data_rows(slide, rows: list, top: float, pal: dict):
    """Build key-value data rows with alternating backgrounds.

    rows: list of (label, value, is_highlight) tuples.
    Returns the Y position after the last row.
    """
    y = top
    row_h = 0.26
    for i, (label, value, highlight) in enumerate(rows):
        # Alternating background
        if i % 2 == 0:
            bg = slide.Shapes.AddShape(1, _inches_to_points(0.55), _inches_to_points(y), _inches_to_points(_USABLE_W), _inches_to_points(row_h))
            bg.Fill.Solid()
            bg.Fill.ForeColor.RGB = _parse_color(pal["surface_alt"])
            bg.Line.Visible = 0

        lbl_color = pal["accent"] if highlight else pal["text_dim"]
        val_color = pal["accent"] if highlight else pal["text"]

        # Label
        tb = slide.Shapes.AddTextbox(1, _inches_to_points(0.65), _inches_to_points(y + 0.02), _inches_to_points(5.0), _inches_to_points(0.22))
        tb.TextFrame.TextRange.Text = label
        tb.TextFrame.TextRange.Font.Size = TYPO["caption"]
        tb.TextFrame.TextRange.Font.Name = "Segoe UI"
        tb.TextFrame.TextRange.Font.Color.RGB = _parse_color(lbl_color)
        if highlight:
            tb.TextFrame.TextRange.Font.Bold = -1

        # Value
        tb2 = slide.Shapes.AddTextbox(1, _inches_to_points(5.5), _inches_to_points(y + 0.02), _inches_to_points(4.0), _inches_to_points(0.22))
        tb2.TextFrame.TextRange.Text = value
        tb2.TextFrame.TextRange.Font.Size = TYPO["caption"]
        tb2.TextFrame.TextRange.Font.Name = "Segoe UI"
        tb2.TextFrame.TextRange.Font.Bold = -1
        tb2.TextFrame.TextRange.Font.Color.RGB = _parse_color(val_color)
        tb2.TextFrame.TextRange.ParagraphFormat.Alignment = 3  # right

        y += row_h + 0.01
    return y


def _build_callout(slide, text: str, top: float, pal: dict):
    """Build an insight callout box with accent left-border."""
    l = _inches_to_points(0.55)
    t = _inches_to_points(top)
    w = _inches_to_points(_USABLE_W)
    h = _inches_to_points(0.5)

    # Background
    shape = slide.Shapes.AddShape(5, l, t, w, h)
    shape.Fill.Solid()
    shape.Fill.ForeColor.RGB = _parse_color(pal["surface"])
    shape.Line.ForeColor.RGB = _parse_color(pal["border"])
    shape.Line.Weight = 0.75

    # Accent left bar
    bar = slide.Shapes.AddShape(1, l, t, _inches_to_points(0.05), h)
    bar.Fill.Solid()
    bar.Fill.ForeColor.RGB = _parse_color(pal["accent"])
    bar.Line.Visible = 0

    # Text
    tb = slide.Shapes.AddTextbox(1, _inches_to_points(0.75), _inches_to_points(top + 0.03), _inches_to_points(_USABLE_W - 0.3), _inches_to_points(0.44))
    tb.TextFrame.TextRange.Text = text
    tb.TextFrame.TextRange.Font.Size = TYPO["body"] - 1
    tb.TextFrame.TextRange.Font.Name = "Segoe UI"
    tb.TextFrame.TextRange.Font.Italic = -1
    tb.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["text_dim"])
    tb.TextFrame.WordWrap = -1


def _clear_slide(pres, slide_index: int):
    """Delete all shapes on a slide."""
    slide = get_slide(pres, slide_index)
    while slide.Shapes.Count > 0:
        slide.Shapes(slide.Shapes.Count).Delete()
    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# COMPOUND SLIDE BUILDER TOOLS (Phase 8)
# ═══════════════════════════════════════════════════════════════════════════════


@mcp.tool()
def build_kpi_slide(
    slide_index: int,
    title: str,
    subtitle: str,
    kpis_json: str,
    badge: str = "",
    rows_json: str = "",
    callout: str = "",
    palette: str = "dark_executive",
    footer_meta: str = "",
    page: str = "",
    clear: bool = True,
) -> str:
    """Build a complete KPI dashboard slide from structured data.

    slide_index: 1-based slide index to rebuild.
    title: main heading text.
    subtitle: secondary text below title.
    kpis_json: JSON array of {number, label, dot_color?, highlight?} objects (1-6 cards).
    badge: section badge text, e.g. 'SITUATION | 02'. Empty to skip.
    rows_json: optional JSON array of [label, value, highlight?] for data table rows.
    callout: optional insight text for bottom callout box. Empty to skip.
    palette: color palette name (dark_executive, midnight_blue, light_corporate).
    footer_meta: footer metadata text. Empty to skip footer.
    page: page number text, e.g. '2 / 8'. Empty to skip.
    clear: if True, delete all existing shapes before building. Default True.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        pal = _get_palette(palette)
        kpis = _validate_json_list(kpis_json, "kpis_json")

        slide = _clear_slide(pres, slide_index) if clear else get_slide(pres, slide_index)

        # Background
        slide.FollowMasterBackground = 0
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = _parse_color(pal["bg"])

        # Decorative circles (background geometry)
        c1 = slide.Shapes.AddShape(9, _inches_to_points(7.5), _inches_to_points(-1.0), _inches_to_points(4.0), _inches_to_points(4.0))
        c1.Fill.Solid()
        c1.Fill.ForeColor.RGB = _parse_color(pal["surface_alt"])
        c1.Line.ForeColor.RGB = _parse_color(pal["primary"])
        c1.Line.Weight = 0.75
        c1.ZOrder(1)  # send to back

        c2 = slide.Shapes.AddShape(9, _inches_to_points(8.2), _inches_to_points(-0.3), _inches_to_points(2.6), _inches_to_points(2.6))
        c2.Fill.Solid()
        c2.Fill.ForeColor.RGB = _parse_color(pal["bg"])
        c2.Line.ForeColor.RGB = _parse_color(pal["primary"])
        c2.Line.Weight = 0.5
        c2.ZOrder(1)

        # Top bar
        _build_top_bar(slide, pal)

        # Badge
        if badge:
            _build_badge(slide, badge, pal)

        # Title block
        _build_title_block(slide, title, subtitle, pal)

        # KPI cards
        card_count = len(kpis)
        positions = _card_positions(card_count, top=1.3, height=1.15)
        dot_colors = [pal["accent"], pal["secondary"], pal.get("accent2", pal["accent"]), "#FFFFFF"]
        for i, kpi in enumerate(kpis):
            dot = kpi.get("dot_color", dot_colors[i % len(dot_colors)])
            _build_kpi_card(
                slide, positions[i],
                number=kpi["number"],
                label=kpi["label"],
                pal=pal,
                is_highlight=kpi.get("highlight", False),
                dot_color=dot,
            )

        # Data rows
        next_y = 2.65
        if rows_json:
            rows = _validate_json_list(rows_json, "rows_json")
            parsed_rows = [(r[0], r[1], r[2] if len(r) > 2 else False) for r in rows]
            # Section label
            lbl = slide.Shapes.AddTextbox(1, _inches_to_points(0.55), _inches_to_points(2.65), _inches_to_points(4.0), _inches_to_points(0.25))
            lbl.TextFrame.TextRange.Text = "CURRENT STATE BREAKDOWN"
            lbl.TextFrame.TextRange.Font.Size = TYPO["section_label"]
            lbl.TextFrame.TextRange.Font.Name = "Segoe UI"
            lbl.TextFrame.TextRange.Font.Bold = -1
            lbl.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["primary"])
            next_y = _build_data_rows(slide, parsed_rows, top=2.92, pal=pal)

        # Callout
        if callout:
            callout_top = max(next_y + 0.08, 4.55)
            _build_callout(slide, callout, callout_top, pal)

        # Footer
        if footer_meta or page:
            _build_footer(slide, pal, meta=footer_meta, page=page)

        return _ok({
            "status": "built",
            "slide_index": slide_index,
            "tool": "build_kpi_slide",
            "kpi_count": card_count,
            "palette": palette,
        })
    except (PPTError, json.JSONDecodeError) as e:
        return _err(e)
    except Exception as e:
        return _err(e)


@mcp.tool()
def build_title_slide(
    slide_index: int,
    title: str,
    subtitle: str,
    metadata_json: str = "",
    palette: str = "dark_executive",
    footer_text: str = "",
    clear: bool = True,
) -> str:
    """Build a polished title/cover slide.

    slide_index: 1-based slide index.
    title: hero title text (large).
    subtitle: tagline below the title.
    metadata_json: optional JSON array of [label, value] pairs (e.g. [["CLIENT","SITE"],["DATE","Feb 2026"]]).
    palette: color palette name.
    footer_text: bottom confidential notice. Empty to skip.
    clear: if True, delete all existing shapes first.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        pal = _get_palette(palette)

        slide = _clear_slide(pres, slide_index) if clear else get_slide(pres, slide_index)

        # Background
        slide.FollowMasterBackground = 0
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = _parse_color(pal["bg"])

        # Decorative circles
        c1 = slide.Shapes.AddShape(9, _inches_to_points(7.5), _inches_to_points(0.5), _inches_to_points(4.0), _inches_to_points(4.0))
        c1.Fill.Solid()
        c1.Fill.ForeColor.RGB = _parse_color(pal["surface_alt"])
        c1.Line.ForeColor.RGB = _parse_color(pal["primary"])
        c1.Line.Weight = 1
        c1.ZOrder(1)

        c2 = slide.Shapes.AddShape(9, _inches_to_points(8.0), _inches_to_points(1.1), _inches_to_points(2.8), _inches_to_points(2.8))
        c2.Fill.Solid()
        c2.Fill.ForeColor.RGB = _parse_color(pal["bg"])
        c2.Line.ForeColor.RGB = _parse_color(pal["primary"])
        c2.Line.Weight = 0.75
        c2.ZOrder(1)

        # Top bars
        _build_top_bar(slide, pal)

        # Left accent bar
        lb = slide.Shapes.AddShape(1, 0, 0, _inches_to_points(0.06), _inches_to_points(_SLIDE_H))
        lb.Fill.Solid()
        lb.Fill.ForeColor.RGB = _parse_color(pal["primary"])
        lb.Line.Visible = 0

        # Hero title
        tb = slide.Shapes.AddTextbox(1, _inches_to_points(0.8), _inches_to_points(1.2), _inches_to_points(7.0), _inches_to_points(1.2))
        tb.TextFrame.TextRange.Text = title
        tb.TextFrame.TextRange.Font.Size = TYPO["hero"]
        tb.TextFrame.TextRange.Font.Name = "Segoe UI"
        tb.TextFrame.TextRange.Font.Bold = -1
        tb.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["text"])

        # Subtitle
        tb2 = slide.Shapes.AddTextbox(1, _inches_to_points(0.8), _inches_to_points(2.45), _inches_to_points(7.0), _inches_to_points(0.4))
        tb2.TextFrame.TextRange.Text = subtitle
        tb2.TextFrame.TextRange.Font.Size = 18
        tb2.TextFrame.TextRange.Font.Name = "Segoe UI"
        tb2.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["secondary"])

        # Divider
        slide.Shapes.AddLine(
            _inches_to_points(0.8), _inches_to_points(3.0),
            _inches_to_points(6.3), _inches_to_points(3.0),
        ).Line.ForeColor.RGB = _parse_color(pal["primary"])

        # Metadata grid
        if metadata_json:
            meta = _validate_json_list(metadata_json, "metadata_json")
            col_w = 3.2
            y = 3.2
            for i, item in enumerate(meta):
                label = item[0] if len(item) > 0 else ""
                value = item[1] if len(item) > 1 else ""
                col = i % 2
                row = i // 2
                x = 0.8 + col * col_w
                cy = y + row * 0.42

                # Label
                lbl = slide.Shapes.AddTextbox(1, _inches_to_points(x), _inches_to_points(cy), _inches_to_points(2.8), _inches_to_points(0.16))
                lbl.TextFrame.TextRange.Text = label.upper()
                lbl.TextFrame.TextRange.Font.Size = 8
                lbl.TextFrame.TextRange.Font.Name = "Segoe UI"
                lbl.TextFrame.TextRange.Font.Bold = -1
                lbl.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["primary"])

                # Value
                val = slide.Shapes.AddTextbox(1, _inches_to_points(x), _inches_to_points(cy + 0.15), _inches_to_points(2.8), _inches_to_points(0.22))
                val.TextFrame.TextRange.Text = value
                val.TextFrame.TextRange.Font.Size = 11
                val.TextFrame.TextRange.Font.Name = "Segoe UI"
                val.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["text_dim"])

        # Footer text
        if footer_text:
            ft = slide.Shapes.AddTextbox(1, _inches_to_points(0.8), _inches_to_points(5.05), _inches_to_points(5.0), _inches_to_points(0.25))
            ft.TextFrame.TextRange.Text = footer_text
            ft.TextFrame.TextRange.Font.Size = TYPO["footer"]
            ft.TextFrame.TextRange.Font.Name = "Segoe UI"
            ft.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["text_muted"])

        return _ok({
            "status": "built",
            "slide_index": slide_index,
            "tool": "build_title_slide",
            "palette": palette,
        })
    except (PPTError, json.JSONDecodeError) as e:
        return _err(e)
    except Exception as e:
        return _err(e)


@mcp.tool()
def build_comparison_slide(
    slide_index: int,
    title: str,
    subtitle: str = "",
    left_json: str = "",
    right_json: str = "",
    badge: str = "",
    callout: str = "",
    palette: str = "dark_executive",
    footer_meta: str = "",
    page: str = "",
    clear: bool = True,
) -> str:
    """Build a two-column comparison slide.

    left_json: JSON object {heading, items: [{text, bold?}]}.
    right_json: JSON object {heading, items: [{text, bold?}]}.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        pal = _get_palette(palette)

        slide = _clear_slide(pres, slide_index) if clear else get_slide(pres, slide_index)

        # Background
        slide.FollowMasterBackground = 0
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = _parse_color(pal["bg"])

        _build_top_bar(slide, pal)
        if badge:
            _build_badge(slide, badge, pal)
        _build_title_block(slide, title, subtitle, pal)

        # Two columns
        col_w = 4.3
        col_gap = 0.4
        col_left_x = _MARGIN
        col_right_x = _MARGIN + col_w + col_gap

        for col_x, data_json, is_right in [(col_left_x, left_json, False), (col_right_x, right_json, True)]:
            if not data_json:
                continue
            data = json.loads(data_json)
            heading = data.get("heading", "")
            items = data.get("items", [])

            # Column heading
            hd = slide.Shapes.AddTextbox(1, _inches_to_points(col_x), _inches_to_points(1.3), _inches_to_points(col_w), _inches_to_points(0.3))
            hd.TextFrame.TextRange.Text = heading
            hd.TextFrame.TextRange.Font.Size = TYPO["section_label"]
            hd.TextFrame.TextRange.Font.Name = "Segoe UI"
            hd.TextFrame.TextRange.Font.Bold = -1
            hd.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["primary"] if not is_right else pal["secondary"])

            # Items
            y = 1.65
            for item in items:
                txt = item.get("text", item) if isinstance(item, dict) else str(item)
                is_bold = item.get("bold", False) if isinstance(item, dict) else False

                tb = slide.Shapes.AddTextbox(1, _inches_to_points(col_x + 0.15), _inches_to_points(y), _inches_to_points(col_w - 0.3), _inches_to_points(0.25))
                tb.TextFrame.TextRange.Text = txt
                tb.TextFrame.TextRange.Font.Size = TYPO["body"]
                tb.TextFrame.TextRange.Font.Name = "Segoe UI"
                tb.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["text"] if is_bold else pal["text_dim"])
                if is_bold:
                    tb.TextFrame.TextRange.Font.Bold = -1
                tb.TextFrame.WordWrap = -1
                y += 0.3

        # Vertical divider between columns
        mid_x = col_left_x + col_w + col_gap / 2
        slide.Shapes.AddLine(
            _inches_to_points(mid_x), _inches_to_points(1.3),
            _inches_to_points(mid_x), _inches_to_points(4.4),
        ).Line.ForeColor.RGB = _parse_color(pal["border"])

        if callout:
            _build_callout(slide, callout, 4.55, pal)
        if footer_meta or page:
            _build_footer(slide, pal, meta=footer_meta, page=page)

        return _ok({
            "status": "built",
            "slide_index": slide_index,
            "tool": "build_comparison_slide",
            "palette": palette,
        })
    except (PPTError, json.JSONDecodeError) as e:
        return _err(e)
    except Exception as e:
        return _err(e)


@mcp.tool()
def build_timeline_slide(
    slide_index: int,
    title: str,
    subtitle: str = "",
    phases_json: str = "",
    badge: str = "",
    callout: str = "",
    palette: str = "dark_executive",
    footer_meta: str = "",
    page: str = "",
    clear: bool = True,
) -> str:
    """Build a horizontal timeline/phases slide.

    phases_json: JSON array of {label, title, description?, timeline?, investment?} objects.
    """
    try:
        app = get_app()
        pres = get_pres(app)
        _require_writable(pres)
        pal = _get_palette(palette)
        phases = _validate_json_list(phases_json, "phases_json") if phases_json else []

        slide = _clear_slide(pres, slide_index) if clear else get_slide(pres, slide_index)

        # Background
        slide.FollowMasterBackground = 0
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = _parse_color(pal["bg"])

        _build_top_bar(slide, pal)
        if badge:
            _build_badge(slide, badge, pal)
        _build_title_block(slide, title, subtitle, pal)

        # Phase cards
        if phases:
            count = len(phases)
            positions = _card_positions(count, top=1.4, height=2.8, gap=0.15)

            # Connecting line through all phases
            first_center = positions[0]["left"] + positions[0]["width"] / 2
            last_center = positions[-1]["left"] + positions[-1]["width"] / 2
            line = slide.Shapes.AddLine(
                _inches_to_points(first_center), _inches_to_points(1.7),
                _inches_to_points(last_center), _inches_to_points(1.7),
            )
            line.Line.ForeColor.RGB = _parse_color(pal["primary"])
            line.Line.Weight = 2

            for i, (phase, pos) in enumerate(zip(phases, positions)):
                is_last = (i == count - 1)
                l, t, w, h = pos["left"], pos["top"], pos["width"], pos["height"]

                # Phase number circle
                cx = l + w / 2 - 0.18
                circle = slide.Shapes.AddShape(9, _inches_to_points(cx), _inches_to_points(t), _inches_to_points(0.36), _inches_to_points(0.36))
                circle.Fill.Solid()
                circle.Fill.ForeColor.RGB = _parse_color(pal["highlight_bg"] if is_last else pal["surface"])
                circle.Line.ForeColor.RGB = _parse_color(pal["primary"])
                circle.Line.Weight = 1.5
                circle.TextFrame.TextRange.Text = phase.get("label", str(i + 1))
                circle.TextFrame.TextRange.Font.Size = 10
                circle.TextFrame.TextRange.Font.Name = "Segoe UI"
                circle.TextFrame.TextRange.Font.Bold = -1
                circle.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["text"])
                circle.TextFrame.TextRange.ParagraphFormat.Alignment = 2

                # Card body
                card = slide.Shapes.AddShape(5, _inches_to_points(l), _inches_to_points(t + 0.5), _inches_to_points(w), _inches_to_points(h - 0.5))
                card.Fill.Solid()
                card.Fill.ForeColor.RGB = _parse_color(pal["highlight_bg"] if is_last else pal["surface"])
                card.Line.ForeColor.RGB = _parse_color(pal["primary"])
                card.Line.Weight = 0.75

                # Phase title
                tb = slide.Shapes.AddTextbox(1, _inches_to_points(l + 0.1), _inches_to_points(t + 0.6), _inches_to_points(w - 0.2), _inches_to_points(0.3))
                tb.TextFrame.TextRange.Text = phase.get("title", "")
                tb.TextFrame.TextRange.Font.Size = 11
                tb.TextFrame.TextRange.Font.Name = "Segoe UI"
                tb.TextFrame.TextRange.Font.Bold = -1
                tb.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["text"])
                tb.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                tb.TextFrame.WordWrap = -1

                # Description
                desc = phase.get("description", "")
                if desc:
                    tb2 = slide.Shapes.AddTextbox(1, _inches_to_points(l + 0.1), _inches_to_points(t + 1.0), _inches_to_points(w - 0.2), _inches_to_points(0.8))
                    tb2.TextFrame.TextRange.Text = desc
                    tb2.TextFrame.TextRange.Font.Size = 8
                    tb2.TextFrame.TextRange.Font.Name = "Segoe UI"
                    tb2.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["text_dim"])
                    tb2.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    tb2.TextFrame.WordWrap = -1

                # Timeline / Investment
                timeline = phase.get("timeline", "")
                invest = phase.get("investment", "")
                extra = []
                if timeline:
                    extra.append(timeline)
                if invest:
                    extra.append(invest)
                if extra:
                    tb3 = slide.Shapes.AddTextbox(1, _inches_to_points(l + 0.1), _inches_to_points(t + h - 0.55), _inches_to_points(w - 0.2), _inches_to_points(0.4))
                    tb3.TextFrame.TextRange.Text = "\n".join(extra)
                    tb3.TextFrame.TextRange.Font.Size = 8
                    tb3.TextFrame.TextRange.Font.Name = "Segoe UI"
                    tb3.TextFrame.TextRange.Font.Color.RGB = _parse_color(pal["secondary"])
                    tb3.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    tb3.TextFrame.WordWrap = -1

        if callout:
            _build_callout(slide, callout, 4.55, pal)
        if footer_meta or page:
            _build_footer(slide, pal, meta=footer_meta, page=page)

        return _ok({
            "status": "built",
            "slide_index": slide_index,
            "tool": "build_timeline_slide",
            "phase_count": len(phases),
            "palette": palette,
        })
    except (PPTError, json.JSONDecodeError) as e:
        return _err(e)
    except Exception as e:
        return _err(e)


@mcp.tool()
def list_palettes() -> str:
    """List all available design palettes with their color roles."""
    try:
        result = []
        for name, pal in PALETTES.items():
            result.append({
                "name": name,
                "colors": {k: v for k, v in pal.items()},
            })
        return _ok_list(result, "palettes")
    except Exception as e:
        return _err(e)


@mcp.tool()
def list_assets() -> str:
    """Scan the assets directory and list available icons, backgrounds, and infographics."""
    try:
        catalog = _scan_assets()
        return _ok({
            "status": "ok",
            "asset_dir": _ASSET_DIR,
            "icons": list(catalog["icons"].keys()),
            "backgrounds": list(catalog["backgrounds"].keys()),
            "infographics": list(catalog["infographics"].keys()),
            "total": sum(len(v) for v in catalog.values()),
        })
    except Exception as e:
        return _err(e)


# ═══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    mcp.run()
