from __future__ import annotations

from pathlib import Path


def containing_app_bundle(app_file: Path) -> Path | None:
    """Return the enclosing macOS ``.app`` bundle for an app source file."""
    app_file = Path(app_file).resolve()
    for parent in app_file.parents:
        if parent.suffix.casefold() == ".app":
            return parent
    return None


def dedicated_product_master_bundle(app_file: Path) -> Path | None:
    """Return Orchid's separately registered Product Master app, if bundled.

    macOS can refuse to create a useful second GUI instance of the same app
    bundle, even when ``open -n`` accepts the request.  Product Master needs a
    distinct bundle identity so Launch Services creates its own foreground
    window instead of attempting to reopen the dashboard process.
    """
    main_bundle = containing_app_bundle(app_file)
    if main_bundle is None:
        return None
    candidate = main_bundle / "Contents" / "Resources" / "Orchid Product Master.app"
    executable = candidate / "Contents" / "MacOS" / "OrchidProductMaster"
    return candidate if executable.is_file() else None


def product_master_launch_command(
    app_file: Path,
    python_executable: str,
    *,
    is_frozen: bool = False,
    platform: str = "",
) -> tuple[list[str], bool]:
    """Return the direct Product Master command using Orchid's live runtime.

    Product Master is a separate process so a catalog save cannot block the
    purchase-packet dashboard.  It must *not* be launched by macOS Launch
    Services through a nested app bundle: that extra handoff was the source of
    the repeated "Opening Product Master" dead end on the user's Mac.  The
    dashboard's already-running Python executable has the correct virtual
    environment and can start the editor directly, on every platform.

    ``False`` means the returned process is the actual editor process, not an
    intermediate launcher.
    """
    app_file = Path(app_file).resolve()
    if is_frozen:
        return [str(python_executable), "--product-master"], False
    return [str(python_executable), str(app_file), "--product-master"], False
