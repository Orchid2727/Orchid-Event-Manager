from __future__ import annotations

"""Scrolling support for the stable Python 3.13 / Tk 8.6 runtime.

CustomTkinter 5.2.2 already installs the correct macOS ``<MouseWheel>`` bindings for
``CTkScrollableFrame``. Previous Orchid releases replaced those bindings with global
routers, which prevented the normal handler from receiving trackpad gestures on some
Macs. Version 3.7 deliberately leaves CustomTkinter's native bindings untouched.
"""

from typing import Iterable

import customtkinter as ctk


def install_native_scroll_support(
    root,
    frames: Iterable[ctk.CTkScrollableFrame] | None = None,
) -> None:
    """Keep CustomTkinter's built-in wheel/trackpad behavior active.

    The function remains for source compatibility with existing screens. We only make
    sure each canvas has a small increment for smoother movement; no global event
    bindings and no monkey patches are installed.
    """
    if frames:
        for frame in frames:
            register_scrollable(root, frame, reinstall=False)


def register_scrollable(root, frame: ctk.CTkScrollableFrame, reinstall: bool = True) -> None:
    """Register a frame without overriding CustomTkinter's own event bindings."""
    if frame is None:
        return
    try:
        canvas = getattr(frame, "_parent_canvas", None)
        if canvas is not None and canvas.winfo_exists():
            canvas.configure(yscrollincrement=1, xscrollincrement=1)
    except Exception:
        pass
