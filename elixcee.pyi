"""
Type stubs for elixcee — Excel VBA emulator (Rust / PyO3).

Row and column numbers are always 1-based (VBA / Excel convention).
"""

from __future__ import annotations

from typing import Any, Optional

# ── ExcelError ────────────────────────────────────────────────────────────────

class ExcelError:
    """Represents an Excel cell error value (#N/A, #VALUE!, #DIV/0!, etc.).

    Returned by :meth:`Vm.get_cell` and :meth:`Vm.cells` for error cells, and
    accepted by :meth:`Vm.set_cell` to store an error value.
    """

    code: str
    """The error string, e.g. ``"#N/A"``, ``"#VALUE!"``, ``"#DIV/0!"``."""

    def __init__(self, code: str) -> None: ...

# ── Vm ────────────────────────────────────────────────────────────────────────

class Vm:
    """A virtual Excel workbook / VBA interpreter.

    All row/column coordinates are **1-based** (matching VBA's ``Cells(row, col)``).
    """

    def __init__(self, on_msgbox: str = "skip") -> None:
        """Create a new VM.

        Parameters
        ----------
        on_msgbox:
            ``"skip"`` (default) silently ignores ``MsgBox`` calls.
            ``"error"`` raises :exc:`RuntimeError` when a ``MsgBox`` is hit.
        """
        ...

    # ── VBA execution ──────────────────────────────────────────────────────────

    def run(self, vba_code: str, macro_name: str) -> None:
        """Parse and execute *macro_name* inside *vba_code*.

        Raises :exc:`SyntaxError` on parse failure, :exc:`RuntimeError` on
        runtime error.
        """
        ...

    # ── Cell access ────────────────────────────────────────────────────────────

    def set_cell(self, row: int, col: int, value: Any) -> None:
        """Write *value* into the cell at (``row``, ``col``) (1-based)."""
        ...

    def get_cell(self, row: int, col: int) -> Any:
        """Return the value at (``row``, ``col``).  Returns ``None`` for empty cells."""
        ...

    def get_cell_number_format(self, row: int, col: int) -> str | None:
        """Return the active sheet's resolved number-format code for a cell.

        E.g. ``"m/d/yyyy"`` for a date-formatted cell. Returns ``None`` for a
        cell with no format, the General format, or a sheet with no
        source-file styles (e.g. one created purely via ``set_sheet()``).
        Lets a caller detect a date-formatted cell (whose value otherwise
        comes back as a raw Excel serial number, e.g. ``45366``) and convert
        it itself.
        """
        ...

    def cells(self) -> dict[tuple[int, int], Any]:
        """Return all non-empty cells of the active sheet as ``{(row, col): value}``."""
        ...

    # ── Formula support ────────────────────────────────────────────────────────

    def set_cell_formula(self, row: int, col: int, formula: str) -> None:
        """Store *formula* (e.g. ``"=SUM(A1:A3)"``) on a cell and evaluate it immediately."""
        ...

    def set_cell_formula_batch(
        self, formulas: dict[tuple[int, int], str]
    ) -> None:
        """Set multiple cell formulas at once.

        Parameters
        ----------
        formulas:
            Mapping of ``(row, col)`` → formula string (e.g. ``"=SUM(A1:A3)"``).
        """
        ...

    def recalculate(self) -> None:
        """Re-evaluate all cells that have a stored formula.

        Useful after writing raw values with :meth:`set_cell` when you want
        dependent formula cells to update.
        """
        ...

    # ── Sheet management ───────────────────────────────────────────────────────

    def set_sheet(self, name: str, index: int | None = None) -> None:
        """Switch the active sheet to *name* (creates it if absent).

        *index* (0-based) places a newly-created sheet at that position among
        the existing sheets instead of appending it at the end; ignored if
        *name* already exists, and clamped rather than erroring if it's past
        the current sheet count.
        """
        ...

    def delete_sheet(self, name: str) -> None:
        """Delete the sheet named *name*. Raises ``ValueError`` if it doesn't exist."""
        ...

    def active_sheet(self) -> str:
        """Return the name of the currently active sheet."""
        ...

    def sheet_names(self) -> list[str]:
        """Return all sheet names in this workbook."""
        ...

    def get_sheet(self, name: str) -> dict[tuple[int, int], Any]:
        """Return all non-empty cells in the named sheet as ``{(row, col): value}``."""
        ...

    # ── Bulk worksheet range/row access ──────────────────────────────────────────
    #
    # A Python-native API for common row/range operations — not a claim of
    # openpyxl compatibility (different return-type contract, no ``Cell``
    # objects). All methods take *sheet* as a keyword; ``None`` (the default)
    # means the active sheet, and passing an explicit sheet name never changes
    # which sheet is active.

    def get_range(self, addr: str, sheet: str | None = None) -> list[list[Any]]:
        """Read a rectangular range (e.g. ``"A1:C5"``), 1-based A1 notation.

        Returns a row-major nested list, ``None`` for empty cells — same
        per-cell typing as :meth:`get_cell`.

        Parameters
        ----------
        addr:
            A single-area A1 range, e.g. ``"A1:C5"`` or a bare cell like ``"B2"``.
        sheet:
            Sheet to read from. Defaults to the active sheet.

        Raises ``ValueError`` on a multi-area, malformed, or reversed address,
        or an unknown *sheet* name.
        """
        ...

    def set_range(
        self, addr: str, values: list[list[Any]], sheet: str | None = None
    ) -> None:
        """Write a rectangular range (e.g. ``"A1:C2"``), 1-based A1 notation.

        *values* must be a strictly rectangular (non-ragged) nested sequence
        whose shape exactly matches *addr*'s row×col shape. ``None`` means an
        empty cell. A string value starting with ``"="`` is stored literally,
        never promoted to a formula — use :meth:`set_cell_formula`/
        :meth:`set_cell_formula_batch` for that. Every value is converted and
        the shape is checked **before** any cell is touched: a validation
        failure leaves every existing cell unchanged.

        Writing into a non-anchor cell of a merged range, or into a protected
        sheet, is **not** blocked — this matches :meth:`set_cell`'s existing
        behavior.

        Parameters
        ----------
        addr:
            A single-area A1 range, e.g. ``"A1:C2"``.
        values:
            A rectangular nested sequence matching *addr*'s shape.
        sheet:
            Sheet to write to. Defaults to the active sheet.

        Raises ``ValueError`` on a bad address, ragged/mismatched shape, or an
        unknown *sheet* name; ``TypeError`` on an unsupported value type.
        """
        ...

    # ── Variables ──────────────────────────────────────────────────────────────

    def variables(self) -> dict[str, Any]:
        """Return all VBA module-level variables as ``{name: value}``."""
        ...

    # Named ranges are registered via ``Range("A1:B3").Name = "MyData"`` in VBA
    # and are then usable anywhere a range address is expected.
    named_ranges: dict[str, str]
    """Workbook-level named ranges: ``{lowercase_name: address_string}``."""

    # ── I/O ───────────────────────────────────────────────────────────────────

    def save_workbook(self, path: str) -> None:
        """Save all sheets to *path*.  Supports ``.xlsx`` and ``.ods``."""
        ...

    def cells_df(self) -> "pandas.DataFrame":  # type: ignore[name-defined]  # noqa: F821
        """Return the active sheet as a **pandas DataFrame** (requires pandas).

        Row indices and column indices are 1-based integers.  Missing cells are
        represented as ``None`` / ``pd.NA``.

        Raises :exc:`ImportError` if pandas is not installed.
        """
        ...

# ── Module-level functions ────────────────────────────────────────────────────

def run_macro(
    vba_code: str,
    macro_name: str,
    on_msgbox: str = "skip",
) -> dict[tuple[int, int], Any]:
    """Run *macro_name* and return all resulting cells as ``{(row, col): value}``.

    Parameters
    ----------
    vba_code:
        Full VBA source containing the target Sub.
    macro_name:
        Name of the Sub to execute.
    on_msgbox:
        ``"skip"`` (default) or ``"error"``.
    """
    ...

def load_workbook(
    path: str,
    sheet: Optional[str] = None,
    on_msgbox: str = "skip",
) -> Vm:
    """Load an ``.xlsx``, ``.xlsm``, or ``.ods`` file into a new :class:`Vm`.

    The VBA source is **not** extracted from the file — pass it separately to
    :meth:`Vm.run`.

    Parameters
    ----------
    path:
        Path to the spreadsheet file.
    sheet:
        Sheet name to set as active.  Defaults to the first sheet.
    on_msgbox:
        ``"skip"`` (default) or ``"error"``.
    """
    ...

def hello() -> str:
    """Return a greeting string (smoke-test helper)."""
    ...
