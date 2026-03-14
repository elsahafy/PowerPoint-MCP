# Contributing to PowerPoint MCP Server

Thank you for your interest in contributing! This guide will help you get started.

## Getting Started

1. **Fork** the repository on GitHub.
2. **Clone** your fork locally:
   ```bash
   git clone git@github.com:<your-username>/PowerPoint-MCP.git
   cd PowerPoint-MCP
   ```
3. **Install dependencies**:
   ```bash
   pip install mcp pywin32
   ```
4. **Open Microsoft PowerPoint** on your Windows machine (required for COM automation).

## Development Environment

- **OS**: Windows (COM automation requires a local PowerPoint installation)
- **Python**: 3.10+
- **Dependencies**: `mcp`, `pywin32`

## Project Structure

```
powerpoint/
  server.py          # Single-file MCP server (all 105 tools)
  PLAN.md            # Implementation plan and tool reference
  tests/
    test_phase1.py   # Integration tests per phase
    ...
```

## How to Contribute

### Reporting Bugs

- Open an issue with a clear title and description.
- Include the tool name, input parameters, and the error output (JSON).
- Mention your PowerPoint version and Windows version.

### Suggesting Features

- Open an issue describing the new tool or enhancement.
- Explain the use case and expected behavior.
- Reference the relevant PowerPoint COM API if possible.

### Submitting Changes

1. Create a feature branch from `main`:
   ```bash
   git checkout -b feature/your-feature-name
   ```
2. Make your changes in `server.py`.
3. Follow the existing code patterns:
   - Every tool uses `@mcp.tool()` decorator
   - Every tool wraps its body in `try/except`
   - Errors return `json.dumps({"error": str(e)}, indent=2)`
   - Success returns `json.dumps({...}, indent=2)`
   - Use existing helpers (`get_app`, `get_pres`, `get_slide`, `get_shape`, etc.)
4. Test your changes with PowerPoint running:
   ```bash
   python tests/test_phaseN.py
   ```
5. Commit with a clear message:
   ```bash
   git commit -m "Add tool_name: brief description"
   ```
6. Push and open a Pull Request against `main`.

## Code Style

- **Single file**: All tools live in `server.py`. Do not split into modules.
- **Docstrings**: Every tool function must have a docstring describing its purpose and parameters.
- **COM safety**: Always use `get_app()` to obtain the PowerPoint instance. Never cache COM objects across calls.
- **MsoTriState**: Use `-1` (msoTrue) and `0` (msoFalse), never Python `True`/`False` for COM booleans.
- **Colors**: Use `_parse_color()` for all color handling (accepts `#RRGGBB` or `R,G,B`).
- **Units**: Positions and sizes are in **inches** at the tool API level, converted to points internally via `_inches_to_points()`.
- **Error handling**: Catch exceptions at the tool level and return JSON error objects. Never let exceptions propagate to the MCP framework.

## Adding a New Tool

1. Add the `@mcp.tool()` function in the appropriate phase section of `server.py`.
2. Follow the signature pattern: parameters with sensible defaults, return `str`.
3. Add a try/except block returning JSON on both success and failure.
4. Update the tool count in `PLAN.md` if applicable.
5. Write a test case in the corresponding `tests/test_phaseN.py`.

## Testing

Integration tests require PowerPoint to be installed and running on Windows:

```bash
python tests/test_phase1.py
python tests/test_phase2.py
# ... etc.
```

Tests create temporary presentations and close them without saving. They should not leave any side effects.

## Pull Request Guidelines

- Keep PRs focused on a single change (one tool, one bug fix, etc.).
- Include a description of what changed and why.
- Ensure all existing tests still pass.
- Add tests for new tools.

## Code of Conduct

Be respectful, constructive, and collaborative. We welcome contributors of all experience levels.

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](LICENSE).
