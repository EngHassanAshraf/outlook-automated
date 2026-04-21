"""
progress_ui.py
--------------
Rich-based terminal UI with a compact pinned-bottom panel.

Layout (always visible, fixed height):

  ╭─ Latest Activity ──────────────────────────────────────────╮
  │  ☑  Daily Report — MVHYDEPARK                              │
  │  ☑  Weekly Technical Status — MV3                          │
  │  ⚠  Collision: 09-02-2026(1).pdf                           │
  │  ☑  Daily Report — MVNC                                    │
  │  ☑  Daily Report — MVHYDEPARK                              │
  ╰────────────────────────────────────────────────────────────╯
  ╭────────────────────────────────────────────────────────────╮
  │ Processing emails ████████░░░░  57.0%  46/80  2.3/s 0:00:32│
  ╰────────────────────────────────────────────────────────────╯

Public API
----------
    ui = ProgressUI(total=80)
    ui.start()
    ui.notify("☑  Daily Report — MVHYDEPARK")
    ui.warn("⚠  Collision: 09-02-2026(1).pdf")
    ui.error("✗  Failed to move mail")
    ui.update(current=46)
    ui.complete("🎊  All done!  (46 processed)")
    ui.stop()
    ui.reset(total=50)   # new session
"""

import logging
from typing import Optional
from collections import deque

from rich.console import Console
from rich.live import Live
from rich.layout import Layout
from rich.panel import Panel
from rich.progress import (
    Progress,
    BarColumn,
    TextColumn,
    TimeElapsedColumn,
    MofNCompleteColumn,
    SpinnerColumn,
    ProgressColumn,
    Task,
    TaskID,
)
from rich.text import Text
from rich.table import Table
from rich import box


_LOG_LINES = 20


class _SpeedColumn(ProgressColumn):
    """Renders processing speed, safely handles None before first tick."""

    def render(self, task: Task) -> Text:
        speed = task.speed
        if speed is None:
            return Text("  —/s  ", style="dim")
        return Text(f"  {speed:.1f}/s  ", style="dim")


class ProgressUI:
    """
    Compact two-panel live UI pinned to the bottom of the terminal.

    Uses Rich Live with refresh_per_second=15 and calls live.update()
    on every state change so the display is always current regardless
    of how fast the processing loop runs.
    """

    def __init__(self, total: Optional[int] = None, log_lines: int = _LOG_LINES):
        self._total = total
        self._current = 0
        self._log_lines = log_lines
        self._log_buffer: deque[tuple[str, str]] = deque(maxlen=log_lines)
        self._done = False
        self._completion_msg: Optional[str] = None

        self._console = Console(highlight=False)
        self._progress: Progress = self._build_progress()
        self._task_id: Optional[TaskID] = None
        self._live: Optional[Live] = None

    # ------------------------------------------------------------------ #
    # Lifecycle                                                            #
    # ------------------------------------------------------------------ #

    def start(self) -> None:
        """Start the live display. Call once before the processing loop."""
        self._done = False
        self._completion_msg = None
        self._current = 0
        self._log_buffer.clear()

        self._progress = self._build_progress()
        self._task_id = self._progress.add_task("Processing emails", total=self._total)

        self._live = Live(
            self._render(),
            console=self._console,
            refresh_per_second=15,
            transient=False,
            vertical_overflow="visible",
            auto_refresh=True,   # background refresh keeps bar animating
        )
        self._live.__enter__()   # use __enter__/__exit__ so we control timing

    def stop(self) -> None:
        """Stop the live display. Safe to call even if never started."""
        if self._live is not None:
            self._live.__exit__(None, None, None)
            self._live = None

    def reset(self, total: Optional[int] = None) -> None:
        """Reset for a new processing session. Calls stop() internally."""
        self.stop()
        self._total = total
        self._current = 0
        self._done = False
        self._completion_msg = None
        self._log_buffer.clear()

    # ------------------------------------------------------------------ #
    # Progress                                                             #
    # ------------------------------------------------------------------ #

    def update(self, current: int) -> None:
        """
        Advance the bar to *current* (absolute, not delta).

        Args:
            current: Items processed so far.
        """
        if self._live is None or self._task_id is None:
            return
        advance = current - self._current
        self._current = current
        self._progress.update(self._task_id, advance=advance)
        self._live.update(self._render())

    def complete(self, message: str = "✅  Done!") -> None:
        """Switch to the success state with a completion message."""
        if self._live is None or self._task_id is None:
            return
        self._done = True
        self._completion_msg = message
        if self._total is not None:
            remaining = self._total - self._current
            if remaining > 0:
                self._progress.update(self._task_id, advance=remaining)
        self._live.update(self._render())

    # ------------------------------------------------------------------ #
    # Log panel                                                            #
    # ------------------------------------------------------------------ #

    def notify(self, message: str) -> None:
        """Push a normal activity line (white)."""
        self._push(message, "white")

    def warn(self, message: str) -> None:
        """Push a warning line (yellow)."""
        self._push(message, "yellow")

    def error(self, message: str) -> None:
        """Push an error line (bold red)."""
        self._push(message, "bold red")

    # ------------------------------------------------------------------ #
    # Internal                                                             #
    # ------------------------------------------------------------------ #

    def _push(self, message: str, style: str) -> None:
        self._log_buffer.append((message, style))
        if self._live is not None:
            self._live.update(self._render())

    def _build_progress(self) -> Progress:
        if self._total is None:
            columns = [
                SpinnerColumn(spinner_name="dots", style="cyan"),
                TextColumn("[bold cyan]{task.description}"),
                TimeElapsedColumn(),
            ]
        else:
            columns = [
                TextColumn("[bold cyan]{task.description}"),
                BarColumn(
                    bar_width=None,
                    style="grey37",
                    complete_style="cyan",
                    finished_style="bright_green",
                ),
                TextColumn("[bold white]{task.percentage:>5.1f}%"),
                MofNCompleteColumn(),
                _SpeedColumn(),
                TimeElapsedColumn(),
            ]
        return Progress(*columns, expand=True)

    def _render(self) -> Layout:
        log_height = self._log_lines + 2   # content rows + top/bottom border
        progress_height = 3                # 1 content row + 2 borders

        layout = Layout()
        layout.split_column(
            Layout(name="log",      size=log_height),
            Layout(name="progress", size=progress_height),
        )

        # ── Log panel ──────────────────────────────────────────────────
        grid = Table.grid(padding=(0, 0))
        grid.add_column(no_wrap=True)

        entries = list(self._log_buffer)
        # Pad from the top so new entries appear at the bottom
        for _ in range(self._log_lines - len(entries)):
            grid.add_row(Text(""))
        for msg, style in entries:
            grid.add_row(Text(f"  {msg}", style=style, no_wrap=True))

        layout["log"].update(Panel(
            grid,
            title="[bold dim]Latest Activity[/]",
            border_style="dim",
            box=box.ROUNDED,
            padding=(0, 10),
        ))

        # ── Progress panel ─────────────────────────────────────────────
        if self._done and self._completion_msg:
            content: Any = Text(f"  {self._completion_msg}", style="bold bright_green")
            border = "bright_green"
        else:
            content = self._progress
            border = "cyan"

        layout["progress"].update(Panel(
            content,
            border_style=border,
            box=box.ROUNDED,
            padding=(0, 10),
        ))

        return layout


# ---------------------------------------------------------------------------
# Silent handler — absorbs console log records while the live UI is active
# so raw timestamps never appear above the Rich panels.
# ---------------------------------------------------------------------------

class SuppressConsoleHandler(logging.Handler):
    """No-op handler that silently absorbs log records."""

    def emit(self, record: logging.LogRecord) -> None:
        pass
