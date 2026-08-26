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

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        """Rename a sheet.

        Renaming the active sheet is supported (it stays active under the new
        name). Renaming a sheet to itself, or to a different casing of its
        own name, succeeds.

        Parameters
        ----------
        old_name:
            The sheet's current name (case-insensitive).
        new_name:
            The new name.

        Raises ``ValueError`` if *old_name* doesn't exist, *new_name* is empty
        or whitespace-only, *new_name* (case-insensitively) already names a
        *different* existing sheet, or the sheet is protected.
        """
        ...

    def move_sheet(self, name: str, new_index: int) -> None:
        """Move a sheet to an absolute 0-based position among the workbook's sheets.

        Unlike openpyxl's ``Worksheet.move_sheet(offset)`` (a relative
        offset), *new_index* here is an absolute target position (0 =
        first), matching :meth:`set_sheet`'s own *index* convention.
        Out-of-range values are clamped to the nearest end rather than
        raising. Does not check sheet protection — real Excel's per-sheet
        protection does not gate tab reordering.

        Raises ``ValueError`` if *name* doesn't exist.
        """
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

    def append_row(self, values: list[Any], sheet: str | None = None) -> int:
        """Write one row just past the sheet's used range.

        Uses the true max used row (row 1 if the sheet is empty/all-empty),
        so this is correct on a sparse sheet. Returns the 1-based row number
        written. Same validate-then-commit and active-sheet-preservation
        guarantees as :meth:`set_range`.

        Raises ``ValueError`` if *values* is empty or *sheet* is unknown;
        ``TypeError`` on an unsupported value type.
        """
        ...

    def iter_rows(
        self,
        min_row: int = 1,
        max_row: int | None = None,
        min_col: int = 1,
        max_col: int | None = None,
        sheet: str | None = None,
    ) -> list[list[Any]]:
        """Values-only iteration over a rectangular region, 1-based bounds.

        *max_row*/*max_col* default to the sheet's used range. On a sheet
        with no non-empty cells at all **and** no explicit *max_row*, returns
        ``[]`` rather than one row of ``None``\\ s.

        Returns plain nested lists — this does **not** claim openpyxl
        ``Cell``-object compatibility (no ``.value``/``.style``/etc attached,
        just the values).
        """
        ...

    def iter_cols(
        self,
        min_row: int = 1,
        max_row: int | None = None,
        min_col: int = 1,
        max_col: int | None = None,
        sheet: str | None = None,
    ) -> list[list[Any]]:
        """Values-only, column-major iteration over a rectangular region —
        the transposed sibling of :meth:`iter_rows`. Each returned inner
        list is one column's values, top to bottom.

        *max_row*/*max_col* default to the sheet's used range. On a sheet
        with no non-empty cells at all **and** no explicit *max_col*, returns
        ``[]`` rather than one column of ``None``\\ s.

        Returns plain nested lists — this does **not** claim openpyxl
        ``Cell``-object compatibility (no ``.value``/``.style``/etc attached,
        just the values).
        """
        ...

    def max_row(self, sheet: str | None = None) -> int | None:
        """Highest used row number, or ``None`` for a sheet with zero
        non-empty cells (never ``0``)."""
        ...

    def max_column(self, sheet: str | None = None) -> int | None:
        """Highest used column number, or ``None`` for a sheet with zero
        non-empty cells (never ``0``)."""
        ...

    def calculate_dimension(self, sheet: str | None = None) -> str | None:
        """The used range as an A1-style string (e.g. ``"B2:D10"``), or
        ``None`` for a sheet with zero non-empty cells (never ``"A1:A1"``).

        Min-anchored, not A1-anchored: if the only populated cell is ``C3``,
        this returns ``"C3:C3"``, not ``"A1:C3"``.
        """
        ...

    def insert_rows(
        self, idx: int, amount: int = 1, sheet: str | None = None
    ) -> None:
        """Insert *amount* blank rows before 1-based row *idx*, shifting *idx*
        and everything below it down. Mirrors openpyxl's
        ``Worksheet.insert_rows(idx, amount=1)``.

        Does **not** shift merged ranges, hidden-row markers, cell
        styles/number formats, or formula cell-reference text — a
        pre-existing limitation of the underlying VBA engine, now reachable
        from Python.

        Parameters
        ----------
        idx:
            1-based row number to insert before.
        amount:
            Number of rows to insert.
        sheet:
            Sheet to modify. Defaults to the active sheet; never changes
            which sheet is active.

        Raises ``ValueError`` if *idx*/*amount* is 0 or exceeds 1,048,576, or
        *sheet* is unknown.
        """
        ...

    def delete_rows(
        self, idx: int, amount: int = 1, sheet: str | None = None
    ) -> None:
        """Delete *amount* rows starting at 1-based row *idx*, shifting
        everything below the deleted band up. Mirrors openpyxl's
        ``Worksheet.delete_rows(idx, amount=1)``.

        Same fidelity gap as :meth:`insert_rows`.

        Raises ``ValueError`` if *idx*/*amount* is 0 or exceeds 1,048,576, or
        *sheet* is unknown.
        """
        ...

    def insert_cols(
        self, idx: int, amount: int = 1, sheet: str | None = None
    ) -> None:
        """Insert *amount* blank columns before 1-based column *idx*,
        shifting *idx* and everything to its right, right. Mirrors
        openpyxl's ``Worksheet.insert_cols(idx, amount=1)``.

        Same fidelity gap as :meth:`insert_rows`.

        Raises ``ValueError`` if *idx*/*amount* is 0 or exceeds 16,384
        (``XFD``), or *sheet* is unknown.
        """
        ...

    def delete_cols(
        self, idx: int, amount: int = 1, sheet: str | None = None
    ) -> None:
        """Delete *amount* columns starting at 1-based column *idx*,
        shifting everything to the right of the deleted band left. Mirrors
        openpyxl's ``Worksheet.delete_cols(idx, amount=1)``.

        Same fidelity gap as :meth:`insert_rows`.

        Raises ``ValueError`` if *idx*/*amount* is 0 or exceeds 16,384
        (``XFD``), or *sheet* is unknown.
        """
        ...

    def merged_cells(self, sheet: str | None = None) -> list[str]:
        """Return every merged range on a sheet as A1-style strings (e.g.
        ``["B1:C1"]``).

        Order matches source-file/insertion order (never re-sorted) — do not
        assume alphabetical or row-major order.

        Raises ``ValueError`` if *sheet* is unknown.
        """
        ...

    def merge_cells(self, addr: str, sheet: str | None = None) -> None:
        """Creates a merge over *addr*.

        Rejects a single-cell address (nothing would actually be merged) and
        rejects a merge that would overlap an existing one on the same
        sheet. Does **not** touch cell values — whatever is in the covered
        cells (if anything) stays exactly as it was.

        Raises ``ValueError`` on a bad, oversized, or single-cell address, an
        overlapping merge, or an unknown *sheet* name.
        """
        ...

    def unmerge_cells(self, addr: str, sheet: str | None = None) -> None:
        """Removes a merge whose range exactly matches *addr*.

        An inexact/partial match is rejected rather than silently no-opping.

        Raises ``ValueError`` on a bad or oversized address, no exact match,
        or an unknown *sheet* name.
        """
        ...

    def hidden_rows(self, sheet: str | None = None) -> list[int]:
        """Every hidden row number on a sheet, as a sorted list of 1-based
        row numbers (e.g. ``[5, 6, 9]``). Expanded, not interval-form.

        Raises ``ValueError`` if *sheet* is unknown.
        """
        ...

    def hidden_columns(self, sheet: str | None = None) -> list[int]:
        """Column-axis mirror of :meth:`hidden_rows`."""
        ...

    def set_row_hidden(
        self, row: int, hidden: bool = True, sheet: str | None = None
    ) -> None:
        """Hides or unhides a single row (1-based).

        Hiding an already-hidden row is a no-op; unhiding an already-visible
        row is a no-op.

        Raises ``ValueError`` if *row* is 0 or exceeds Excel's own grid limit
        (1,048,576 rows), or *sheet* is unknown.
        """
        ...

    def set_column_hidden(
        self, col: int, hidden: bool = True, sheet: str | None = None
    ) -> None:
        """Column-axis mirror of :meth:`set_row_hidden`.

        Raises ``ValueError`` if *col* is 0 or exceeds Excel's own grid limit
        (16,384 columns), or *sheet* is unknown.
        """
        ...

    def sort_range(
        self,
        addr: str,
        key_col: int,
        descending: bool = False,
        header: bool = False,
        sheet: str | None = None,
    ) -> None:
        """Python-native, single-key sort of a rectangular range, in place.

        Not from openpyxl (which has no sort primitive of its own) — this
        exposes the existing VBA ``Range(addr).Sort key:=, order:=,
        header:=`` statement's exact behavior to Python.

        *header=True* excludes *addr*'s first row from the sort; it stays
        exactly where it is. Does **not** check sheet protection — matches
        :meth:`set_range`'s bulk cell-value-write precedent.

        Raises ``ValueError`` on a bad or oversized address, a *key_col*
        outside *addr*'s own column span, or an unknown *sheet* name.
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
