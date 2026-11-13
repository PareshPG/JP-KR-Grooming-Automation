"""
Writes cell edits directly into a workbook's worksheet XML inside the .xlsx
zip, and copies every OTHER part of the zip through completely unchanged.

Why this exists: a plain openpyxl load-and-save with ZERO edits was tested
against this workbook and already drops 73 internal parts - xl/connections.xml
(the Power Query connection definitions), all 22 xl/queryTables/*.xml parts,
xl/metadata.xml (which the FILTER/UNIQUE formulas in Worth it & UBA_App/
UBA_BR depend on), relationship files for over a dozen Excel Tables, and the
customXml parts - one of which (customXml/item2.xml) was confirmed by direct
inspection to be a <DataMashup> container: Excel's actual storage format for
a compiled Power Query package. This is a known, long-standing openpyxl
limitation with complex workbooks (pivot caches, external connections, and
dynamic-array metadata aren't fully round-tripped on save) - not a bug in
any one script. Every prior step in this pipeline used openpyxl's
load-edit-save pattern on the whole workbook, which silently corrupted every
output file produced so far. This module replaces that pattern for writes.

Reads elsewhere in this project continue to use openpyxl (read_only mode is
fast and never writes, so it carries none of this risk) - only the WRITE
side needed to change.
"""
import re
import zipfile
import datetime
from lxml import etree
from openpyxl.utils import get_column_letter

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"

EXCEL_EPOCH = datetime.datetime(1899, 12, 30)  # standard trick: this epoch
# naturally reproduces Excel's 1900-leap-year-bug serial numbering for any
# real-world date without special-casing. Cross-checked against openpyxl's
# own to_excel(): both give 46199.0 for 2026-06-26.


def to_excel_serial(dt):
    delta = dt - EXCEL_EPOCH
    return delta.days + delta.seconds / 86400


def col_letters_to_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n


def parse_cell_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError(f"Not a valid cell reference: {ref!r}")
    col_letters, row_num = m.group(1), int(m.group(2))
    return row_num, col_letters_to_num(col_letters)


def resolve_sheet_path(zf, sheet_name):
    """Map a sheet's display name to its internal xl/worksheets/sheetN.xml part."""
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    # XML escapes special characters in attribute values - a sheet name like
    # "Worth it & UBA_App" appears in the raw XML as "Worth it &amp; UBA_App"
    xml_escaped_name = (
        sheet_name.replace("&", "&amp;").replace("<", "&lt;")
        .replace(">", "&gt;").replace('"', "&quot;")
    )
    pattern = (
        rf'<sheet[^>]*(?:name="{re.escape(xml_escaped_name)}"[^>]*r:id="(rId\d+)"'
        rf'|r:id="(rId\d+)"[^>]*name="{re.escape(xml_escaped_name)}")'
    )
    m = re.search(pattern, wb_xml)
    if not m:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.xml")
    rid = m.group(1) or m.group(2)

    rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m2 = re.search(
        rf'(?:Id="{rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="{rid}")',
        rels_xml,
    )
    if not m2:
        raise ValueError(f"Relationship '{rid}' not found in workbook.xml.rels")
    target = m2.group(1) or m2.group(2)
    return target if target.startswith("xl/") else f"xl/{target}"


class SheetEditor:
    """Edits a single worksheet's XML tree in memory. Row/cell elements must
    stay in ascending order for Excel to accept the file, so new rows/cells
    are inserted in sorted position - but insertion position is found via an
    index built once at construction time, not by re-scanning the tree on
    every call (which made bulk edits, e.g. 12,000+ cells in one sheet,
    too slow to be usable)."""

    def __init__(self, xml_bytes):
        self.tree = etree.fromstring(xml_bytes)
        self.sheet_data = self.tree.find(f"{{{MAIN_NS}}}sheetData")
        if self.sheet_data is None:
            raise ValueError("sheetData not found - unexpected worksheet structure")

        # row_num -> row element; sorted list of row_nums kept alongside for
        # fast insertion-point lookup via bisect
        self._rows_by_num = {}
        self._row_nums_sorted = []
        for row_el in self.sheet_data:
            r = int(row_el.get("r"))
            self._rows_by_num[r] = row_el
            self._row_nums_sorted.append(r)

        # per-row: col_num -> cell element, plus sorted col list
        self._cells_by_row = {}
        self._col_nums_sorted_by_row = {}
        for r, row_el in self._rows_by_num.items():
            cells = {}
            cols_sorted = []
            for cell_el in row_el:
                c = parse_cell_ref(cell_el.get("r"))[1]
                cells[c] = cell_el
                cols_sorted.append(c)
            self._cells_by_row[r] = cells
            self._col_nums_sorted_by_row[r] = cols_sorted

    def _get_or_create_row(self, row_num):
        row_el = self._rows_by_num.get(row_num)
        if row_el is not None:
            return row_el

        import bisect
        idx = bisect.bisect_left(self._row_nums_sorted, row_num)
        new_row = etree.Element(f"{{{MAIN_NS}}}row", r=str(row_num))
        if idx < len(self._row_nums_sorted):
            next_row_el = self._rows_by_num[self._row_nums_sorted[idx]]
            next_row_el.addprevious(new_row)
        else:
            self.sheet_data.append(new_row)

        self._row_nums_sorted.insert(idx, row_num)
        self._rows_by_num[row_num] = new_row
        self._cells_by_row[row_num] = {}
        self._col_nums_sorted_by_row[row_num] = []
        return new_row

    def _get_or_create_cell(self, row_el, cell_ref, row_num, col_num):
        cells = self._cells_by_row[row_num]
        cell_el = cells.get(col_num)
        if cell_el is not None:
            return cell_el

        import bisect
        cols_sorted = self._col_nums_sorted_by_row[row_num]
        idx = bisect.bisect_left(cols_sorted, col_num)
        new_cell = etree.Element(f"{{{MAIN_NS}}}c", r=cell_ref)
        if idx < len(cols_sorted):
            next_cell_el = cells[cols_sorted[idx]]
            next_cell_el.addprevious(new_cell)
        else:
            row_el.append(new_cell)

        cols_sorted.insert(idx, col_num)
        cells[col_num] = new_cell
        return new_cell

    def set_cell(self, cell_ref, value):
        """Set a cell's value: number, datetime, or string (written as an
        inline string, t="inlineStr", specifically to avoid needing to touch
        sharedStrings.xml or renumber any other cell's string reference).
        Keeps the cell's existing style (s attribute) untouched - that's
        what makes an already-date-formatted cell keep displaying as a date
        after only its <v> content changes."""
        row_num, col_num = parse_cell_ref(cell_ref)
        row_el = self._get_or_create_row(row_num)
        cell_el = self._get_or_create_cell(row_el, cell_ref, row_num, col_num)

        for child in list(cell_el):
            cell_el.remove(child)

        if value is None:
            if "t" in cell_el.attrib:
                del cell_el.attrib["t"]
            return

        if isinstance(value, datetime.datetime):
            if "t" in cell_el.attrib:
                del cell_el.attrib["t"]
            v_el = etree.SubElement(cell_el, f"{{{MAIN_NS}}}v")
            v_el.text = repr(to_excel_serial(value))
        elif isinstance(value, (int, float)):
            if "t" in cell_el.attrib:
                del cell_el.attrib["t"]
            v_el = etree.SubElement(cell_el, f"{{{MAIN_NS}}}v")
            v_el.text = repr(value)
        else:
            cell_el.set("t", "inlineStr")
            is_el = etree.SubElement(cell_el, f"{{{MAIN_NS}}}is")
            t_el = etree.SubElement(is_el, f"{{{MAIN_NS}}}t")
            t_el.text = str(value)
            t_el.set(f"{{{XML_NS}}}space", "preserve")

    def clear_cell(self, cell_ref):
        """Remove a cell entirely - used to clear stale data beyond new
        content's extent (e.g. old IDxAGE/J:Q rows no longer needed), so
        formulas relying on blank-vs-populated to find the end of real data
        see a genuinely empty cell, not leftover content."""
        row_num, col_num = parse_cell_ref(cell_ref)
        row_el = self._rows_by_num.get(row_num)
        if row_el is None:
            return
        cell_el = self._cells_by_row[row_num].pop(col_num, None)
        if cell_el is not None:
            row_el.remove(cell_el)
            self._col_nums_sorted_by_row[row_num].remove(col_num)

    def set_cell_style_only(self, cell_ref, style_index):
        """Set ONLY a cell's style reference (the 's' attribute) - does not
        touch the cell's value, formula, or any child elements at all. This
        is what color-coding uses: it must never disturb the formula that's
        already computing the cell's value, only add a fill/font color on
        top of it."""
        row_num, col_num = parse_cell_ref(cell_ref)
        row_el = self._get_or_create_row(row_num)
        cell_el = self._get_or_create_cell(row_el, cell_ref, row_num, col_num)
        cell_el.set("s", str(style_index))

    def get_cell_style_index(self, cell_ref):
        """Read a cell's current style index (the 's' attribute, defaults
        to 0 if absent) - used to find the base style to derive a colored
        variant from."""
        row_num, col_num = parse_cell_ref(cell_ref)
        row_el = self._rows_by_num.get(row_num)
        if row_el is None:
            return 0
        cell_el = self._cells_by_row[row_num].get(col_num)
        if cell_el is None:
            return 0
        return int(cell_el.get("s", "0"))

    def clear_rows(self, start_row, end_row):
        """Remove entire <row> elements in [start_row, end_row] inclusive -
        for wiping a full row range clean before rewriting it (IDxAGE, J:Q
        rebuilds). Cleaner than clearing cells one at a time: leaves no
        empty <row> elements behind, so max_row reflects the real new
        extent rather than the largest row ever used."""
        for r in list(self._rows_by_num.keys()):
            if start_row <= r <= end_row:
                row_el = self._rows_by_num.pop(r)
                self.sheet_data.remove(row_el)
                self._row_nums_sorted.remove(r)
                self._cells_by_row.pop(r, None)
                self._col_nums_sorted_by_row.pop(r, None)

    def _recompute_dimension(self):
        """Update the <dimension> tag to match the sheet's actual used range
        after edits - Excel and openpyxl both use this as a used-range hint,
        and leaving it stale (e.g. still declaring 2263 rows after clearing
        down to 300) is exactly the kind of small inconsistency worth not
        introducing, even though it isn't relied on by the formulas here."""
        min_row = min_col = None
        max_row = max_col = None
        for row_el in self.sheet_data:
            r = int(row_el.get("r"))
            has_cells = False
            for cell_el in row_el:
                has_cells = True
                _, c = parse_cell_ref(cell_el.get("r"))
                min_col = c if min_col is None else min(min_col, c)
                max_col = c if max_col is None else max(max_col, c)
            if has_cells:
                min_row = r if min_row is None else min(min_row, r)
                max_row = r if max_row is None else max(max_row, r)

        dim_el = self.tree.find(f"{{{MAIN_NS}}}dimension")
        if dim_el is None:
            return
        if min_row is None:
            dim_el.set("ref", "A1")
        else:
            from openpyxl.utils import get_column_letter
            start = f"{get_column_letter(min_col)}{min_row}"
            end = f"{get_column_letter(max_col)}{max_row}"
            dim_el.set("ref", start if start == end else f"{start}:{end}")

    def to_bytes(self):
        self._recompute_dimension()
        return etree.tostring(self.tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_existing_cells_fast(xml_bytes, fixes):
    """Fast path for editing MANY pre-existing cells in a LARGE sheet, e.g.
    Respondent Data's Rating cleanup (up to ~17,000 candidate cells across
    2,863 rows x 677 columns - 1.9 million <c> elements in total). A full
    DOM parse of a sheet this size (etree.fromstring) alone measured at
    ~16 seconds, and building a further Python-level index on top of that
    made bulk edits impractically slow. This never creates new cells or
    rows - only rewrites the <v> (or inline string) content of cells that
    are already present - so it can safely operate as a single left-to-right
    regex substitution over the raw XML text instead of a full DOM parse.
    This is safe specifically because well-formed XML escapes '<', '>', and
    '&' in text content - a literal "</c>" byte sequence can only ever be a
    real element boundary, never part of an inline string's value.

    fixes: {cell_ref: new_value} where new_value is int/float/str (no
    datetime support needed here - this path is only used for numeric
    Rating cleanup).
    """
    text = xml_bytes.decode("utf-8")
    cell_pattern = re.compile(r'<c r="([A-Z]+\d+)"([^>]*?)(/>|>(.*?)</c>)', re.DOTALL)

    remaining = dict(fixes)  # cell_ref -> value, popped as found

    def replace(m):
        cell_ref = m.group(1)
        if cell_ref not in remaining:
            return m.group(0)
        value = remaining.pop(cell_ref)
        attrs = m.group(2)
        # strip any existing t="..." attribute - we're always writing a
        # plain number here, never a string, for this fast path
        attrs = re.sub(r'\s*t="[^"]*"', "", attrs)
        if isinstance(value, (int, float)):
            return f'<c r="{cell_ref}"{attrs}><v>{value!r}</v></c>'
        # string fallback, written as inline string
        attrs = attrs + ' t="inlineStr"'
        return f'<c r="{cell_ref}"{attrs}><is><t xml:space="preserve">{value}</t></is></c>'

    new_text = cell_pattern.sub(replace, text)

    if remaining:
        raise ValueError(
            f"{len(remaining)} target cell(s) were not found as existing "
            f"cells in the sheet XML - this fast path only edits EXISTING "
            f"cells: {list(remaining.keys())[:10]}"
        )

    return new_text.encode("utf-8")


def combine_edits(*edit_funcs):
    """Combine multiple edit functions targeting the same sheet into one -
    needed when a sheet gets edited by more than one step (e.g. Worth it &
    UBA_App gets both its J:Q rebuild AND the report-month T1 cell)."""
    def combined(editor):
        for fn in edit_funcs:
            fn(editor)
    return combined


CALC_CHAIN_PATH = "xl/calcChain.xml"


def _drop_calc_chain(items):
    """Remove xl/calcChain.xml and its two references ([Content_Types].xml,
    xl/_rels/workbook.xml.rels) from a list of (ZipInfo, bytes) items.

    Why this is necessary and safe: calcChain.xml is a pure calculation-order
    cache - which cell to recalculate before which other cell - and contains
    no actual user data. But it's tightly version-locked to the exact prior
    state of every formula in the workbook; editing cell content anywhere
    (even in cells that aren't formulas themselves) is enough for Excel to
    consider it stale, and Excel then flags the whole file as needing repair
    on open. This was confirmed directly: Excel's own repair log named
    exactly this ("Removed Records: Formula from /xl/calcChain.xml part"),
    and the file opened afterward with every formula, value, and format
    fully intact - calcChain was the ONLY problem. Excel silently and safely
    regenerates this file on open/recalculate. Dropping it here means the
    output file is internally consistent about its absence (no dangling
    Content_Types/relationship entries pointing at a missing part) rather
    than relying on Excel's repair-and-recover path to notice and fix it."""
    out = []
    for item, data in items:
        if item.filename == CALC_CHAIN_PATH:
            continue
        if item.filename == "[Content_Types].xml":
            text = data.decode("utf-8")
            text = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", text)
            data = text.encode("utf-8")
        elif item.filename == "xl/_rels/workbook.xml.rels":
            text = data.decode("utf-8")
            text = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', "", text)
            data = text.encode("utf-8")
        out.append((item, data))
    return out


class StylesEditor:
    """Edits xl/styles.xml to create colored variants of existing cell
    styles - used for the equity color-coding feature, which needs to set
    a cell's fill and font color while preserving everything else about its
    current style (number format, borders, alignment). Reuses an existing
    fill/font/xf combination if an identical one already exists, rather
    than creating duplicates on every run."""

    def __init__(self, xml_bytes):
        self.tree = etree.fromstring(xml_bytes)
        self.fonts = self.tree.find(f"{{{MAIN_NS}}}fonts")
        self.fills = self.tree.find(f"{{{MAIN_NS}}}fills")
        self.cellXfs = self.tree.find(f"{{{MAIN_NS}}}cellXfs")
        self._style_cache = {}

    def _find_or_add_fill(self, rgb):
        rgb = rgb.upper()
        for i, fill_el in enumerate(self.fills):
            pattern = fill_el.find(f"{{{MAIN_NS}}}patternFill")
            if pattern is None or pattern.get("patternType") != "solid":
                continue
            fg = pattern.find(f"{{{MAIN_NS}}}fgColor")
            if fg is not None and fg.get("rgb", "").upper().endswith(rgb):
                return i
        fill_el = etree.SubElement(self.fills, f"{{{MAIN_NS}}}fill")
        pattern = etree.SubElement(fill_el, f"{{{MAIN_NS}}}patternFill", patternType="solid")
        etree.SubElement(pattern, f"{{{MAIN_NS}}}fgColor", rgb=f"FF{rgb}")
        etree.SubElement(pattern, f"{{{MAIN_NS}}}bgColor", indexed="64")
        return len(self.fills) - 1

    def _find_or_add_font(self, base_font_index, rgb):
        rgb = rgb.upper()
        base_font = self.fonts[base_font_index]
        # Build the target font's XML: base font's properties, with color replaced
        candidate = etree.fromstring(etree.tostring(base_font))
        color_el = candidate.find(f"{{{MAIN_NS}}}color")
        if color_el is not None:
            candidate.remove(color_el)
        color_el = etree.Element(f"{{{MAIN_NS}}}color", rgb=f"FF{rgb}")
        candidate.insert(0, color_el)
        candidate_bytes = etree.tostring(candidate)

        # Search for an EXISTING font already identical to the candidate -
        # compare fonts as they actually are, not after modifying both sides
        # to the same color (which would make every font with matching
        # name/size trivially "equal" regardless of its real stored color).
        for i, font_el in enumerate(self.fonts):
            if etree.tostring(font_el) == candidate_bytes:
                return i

        self.fonts.append(candidate)
        return len(self.fonts) - 1

    def get_or_create_colored_style(self, base_xf_index, fill_rgb, font_rgb):
        """Returns a cellXfs index with the same numFmt/border/alignment as
        base_xf_index, but with the given fill and font colors applied."""
        cache_key = (base_xf_index, fill_rgb.upper(), font_rgb.upper())
        if cache_key in self._style_cache:
            return self._style_cache[cache_key]

        base_xf = self.cellXfs[base_xf_index]
        base_font_id = int(base_xf.get("fontId", "0"))

        new_fill_id = self._find_or_add_fill(fill_rgb)
        new_font_id = self._find_or_add_font(base_font_id, font_rgb)

        new_xf = etree.fromstring(etree.tostring(base_xf))
        new_xf.set("fontId", str(new_font_id))
        new_xf.set("fillId", str(new_fill_id))
        new_xf.set("applyFont", "1")
        new_xf.set("applyFill", "1")
        new_xf_bytes = etree.tostring(new_xf)

        for i, xf_el in enumerate(self.cellXfs):
            if etree.tostring(xf_el) == new_xf_bytes:
                self._style_cache[cache_key] = i
                return i

        self.cellXfs.append(new_xf)
        new_index = len(self.cellXfs) - 1
        self._style_cache[cache_key] = new_index
        return new_index

    def to_bytes(self):
        self.fonts.set("count", str(len(self.fonts)))
        self.fills.set("count", str(len(self.fills)))
        self.cellXfs.set("count", str(len(self.cellXfs)))
        return etree.tostring(self.tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_workbook_with_styles(original_path, output_path, sheet_edit_funcs, styles_edit_func):
    """Like patch_workbook, but also allows editing xl/styles.xml via a
    StylesEditor. styles_edit_func: callable(StylesEditor) -> None, called
    BEFORE the sheet edits, since sheet edits (setting a cell's style index)
    depend on styles already having been created."""
    zf = zipfile.ZipFile(original_path, "r")
    sheet_paths = {name: resolve_sheet_path(zf, name) for name in sheet_edit_funcs}
    path_to_name = {path: name for name, path in sheet_paths.items()}

    items = [(item, zf.read(item.filename)) for item in zf.infolist()]
    zf.close()
    items = _drop_calc_chain(items)

    styles_editor = None
    styles_bytes = None
    for item, data in items:
        if item.filename == "xl/styles.xml":
            styles_editor = StylesEditor(data)
            styles_edit_func(styles_editor)
            styles_bytes = styles_editor.to_bytes()
            break

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in items:
            if item.filename == "xl/styles.xml" and styles_bytes is not None:
                data = styles_bytes
            sheet_name = path_to_name.get(item.filename)
            if sheet_name:
                editor = SheetEditor(data)
                sheet_edit_funcs[sheet_name](editor)
                data = editor.to_bytes()
            out.writestr(item, data)
    return sheet_paths, styles_editor


def _escape_xml_text(s):
    """Minimal XML text escaping for inline strings - & and < must always
    be escaped; > only needs escaping in specific contexts but escaping it
    everywhere is harmless and simpler to reason about."""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_bulk_rows_xml(rows, style_map, start_row_num, extra_cells_by_row=None):
    """Construct <row> XML for a large block of data rows as one joined
    string - NOT via DOM element creation, which is far too slow at the
    scale this is needed for (hundreds of thousands to millions of cells).
    Confirmed via clean_ratings.py that regex/string-based approaches stay
    fast at this scale where DOM parsing does not.

    rows: list of row-values (each a list/tuple, one value per column,
          aligned 1:1 with style_map - use None for a blank cell).
    style_map: list of style indices, one per column, 0-indexed to match
               each row's values - reused for every row (confirmed the
               real template uses one consistent style per column across
               all data rows, not per-row styling).
    start_row_num: the Excel row number of the first row in `rows`.
    extra_cells_by_row: optional {row_num: xml_string} - pre-built cell
        XML to append AFTER this row's generated cells (at higher column
        letters, so ascending column order is preserved by simple
        concatenation). Used to carry forward existing cells - e.g. live
        summary formulas - that sit beyond the column range being
        reconstructed and must never be silently dropped.

    Blank cells (None or "") are omitted entirely rather than written as
    empty elements - standard OOXML sparse representation, and consistent
    with how genuinely blank cells already appear elsewhere in these files.
    Strings are written as inline strings (t="inlineStr"), matching the
    convention already used throughout this project to avoid ever needing
    to touch sharedStrings.xml."""
    extra_cells_by_row = extra_cells_by_row or {}
    row_chunks = []
    for i, row_values in enumerate(rows):
        r = start_row_num + i
        cell_chunks = []
        for col_idx, value in enumerate(row_values):
            if value is None:
                continue
            col_letter = get_column_letter(col_idx + 1)
            cell_ref = f"{col_letter}{r}"
            style = style_map[col_idx] if col_idx < len(style_map) else None
            style_attr = f' s="{style}"' if style else ""
            if value == "":
                cell_chunks.append(f'<c r="{cell_ref}"{style_attr} t="inlineStr"><is><t></t></is></c>')
            elif isinstance(value, bool):
                cell_chunks.append(f'<c r="{cell_ref}"{style_attr} t="b"><v>{int(value)}</v></c>')
            elif isinstance(value, (int, float)):
                cell_chunks.append(f'<c r="{cell_ref}"{style_attr}><v>{value}</v></c>')
            elif isinstance(value, datetime.datetime):
                serial = to_excel_serial(value)
                cell_chunks.append(f'<c r="{cell_ref}"{style_attr}><v>{serial}</v></c>')
            else:
                escaped = _escape_xml_text(str(value))
                cell_chunks.append(f'<c r="{cell_ref}"{style_attr} t="inlineStr"><is><t>{escaped}</t></is></c>')
        if r in extra_cells_by_row:
            cell_chunks.append(extra_cells_by_row[r])
        row_chunks.append(f'<row r="{r}">{"".join(cell_chunks)}</row>')
    return "".join(row_chunks)


def patch_workbook_bulk_replace(original_path, output_path, sheet_replacements):
    """Replace all data rows from a given starting row onward in one or
    more large sheets, preserving header rows and everything else in the
    workbook byte-for-byte. Built specifically for reconstructing
    Respondent Data / Worth it & UBA tabs from a fresh flat file each
    month, where the row count can grow and a full DOM rebuild would be
    far too slow.

    sheet_replacements: {sheet_name: (data_start_row, rows, style_map)}
    or {sheet_name: (data_start_row, rows, style_map, preserve_columns_beyond)}
    - data_start_row: first row number to replace (header rows before
      this are left completely untouched)
    - rows: list of row-values, see build_bulk_rows_xml
    - style_map: list of style indices, one per column
    - preserve_columns_beyond: optional 0-indexed column count. If given,
      any EXISTING cell in a replaced row whose column index is beyond
      this count is extracted and carried forward into the new row,
      instead of being discarded along with the rest of the old row.
      Needed because some sheets (confirmed: Worth it & UBA_App/BR) have
      a small block of live summary formulas sitting in columns beyond
      the data range being reconstructed - naively replacing the whole
      row silently destroyed these formulas on every run until this was
      added. Omit (or pass None) for sheets with no such trailing content
      (confirmed: Respondent Data, where the reconstructed row already
      covers every real column).

    Uses regex-based row removal (not DOM) for the same performance
    reason as the row construction above: matching and removing existing
    <row> elements at this scale needs to stay a linear string operation,
    not an XML tree rebuild."""
    zf = zipfile.ZipFile(original_path, "r")
    sheet_paths = {name: resolve_sheet_path(zf, name) for name in sheet_replacements}
    path_to_name = {path: name for name, path in sheet_paths.items()}

    items = [(item, zf.read(item.filename)) for item in zf.infolist()]
    zf.close()
    items = _drop_calc_chain(items)

    row_pattern = re.compile(rb'<row r="(\d+)"[^>]*(?:/>|>.*?</row>)', re.DOTALL)
    cell_pattern = re.compile(rb'<c r="([A-Z]+)(\d+)"[^>]*(?:/>|>.*?</c>)', re.DOTALL)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in items:
            sheet_name = path_to_name.get(item.filename)
            if sheet_name:
                spec = sheet_replacements[sheet_name]
                data_start_row, rows, style_map = spec[0], spec[1], spec[2]
                preserve_columns_beyond = spec[3] if len(spec) > 3 else None

                text = data.decode("utf-8")
                sheet_data_start = text.index("<sheetData>") + len("<sheetData>")
                sheet_data_end = text.index("</sheetData>")
                before = text[:sheet_data_start]
                body = text[sheet_data_start:sheet_data_end]
                after = text[sheet_data_end:]

                kept_rows = []
                extra_cells_by_row = {}
                for m in row_pattern.finditer(body.encode("utf-8")):
                    row_num = int(m.group(1))
                    if row_num < data_start_row:
                        kept_rows.append(m.group(0).decode("utf-8"))
                    elif preserve_columns_beyond:
                        # This row is being replaced - before discarding it,
                        # pull out any cell beyond the preserve boundary
                        row_xml_bytes = m.group(0)
                        preserved = []
                        for cm in cell_pattern.finditer(row_xml_bytes):
                            col_letters = cm.group(1).decode("utf-8")
                            if col_letters_to_num(col_letters) > preserve_columns_beyond:
                                preserved.append(cm.group(0).decode("utf-8"))
                        if preserved:
                            extra_cells_by_row[row_num] = "".join(preserved)

                new_rows_xml = build_bulk_rows_xml(rows, style_map, data_start_row, extra_cells_by_row)
                new_body = "".join(kept_rows) + new_rows_xml
                new_text = before + new_body + after

                # Update <dimension> to reflect the new extent - use the
                # widest of (my own columns) and (any preserved column)
                last_row = data_start_row + len(rows) - 1
                num_cols = len(style_map)
                if extra_cells_by_row:
                    widest_preserved = max(
                        col_letters_to_num(m2.group(1).decode("utf-8"))
                        for xml in extra_cells_by_row.values()
                        for m2 in [cell_pattern.search(xml.encode("utf-8"))] if m2
                    ) if any(extra_cells_by_row.values()) else num_cols
                    num_cols = max(num_cols, widest_preserved)
                last_col_letter = get_column_letter(num_cols) if num_cols else "A"
                dim_match = re.search(r'<dimension ref="([^"]+)"/>', new_text)
                if dim_match:
                    start_ref = dim_match.group(1).split(":")[0]
                    new_dim = f'<dimension ref="{start_ref}:{last_col_letter}{last_row}"/>'
                    new_text = new_text[:dim_match.start()] + new_dim + new_text[dim_match.end():]

                data = new_text.encode("utf-8")
            out.writestr(item, data)

    return sheet_paths


def patch_workbook(original_path, output_path, sheet_edit_funcs):
    """sheet_edit_funcs: {sheet_name: callable(SheetEditor)}.
    Copies the ENTIRE original .xlsx zip through unchanged except the
    specific worksheet XML parts named in sheet_edit_funcs.

    Reads everything from original_path into memory BEFORE opening
    output_path for writing - required for in-place edits where
    output_path == original_path, since opening a zip in write mode
    truncates it immediately, which would corrupt a read still in progress
    against the same file."""
    zf = zipfile.ZipFile(original_path, "r")
    sheet_paths = {name: resolve_sheet_path(zf, name) for name in sheet_edit_funcs}
    path_to_name = {path: name for name, path in sheet_paths.items()}

    items = [(item, zf.read(item.filename)) for item in zf.infolist()]
    zf.close()
    items = _drop_calc_chain(items)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in items:
            sheet_name = path_to_name.get(item.filename)
            if sheet_name:
                editor = SheetEditor(data)
                sheet_edit_funcs[sheet_name](editor)
                data = editor.to_bytes()
            out.writestr(item, data)
    return sheet_paths


def patch_workbook_fast_cells(original_path, output_path, sheet_fixes):
    """Like patch_workbook, but for the large-sheet/many-existing-cell-edits
    case: sheet_fixes is {sheet_name: {cell_ref: value}}. Uses the regex
    fast path (patch_existing_cells_fast) instead of a full DOM parse.
    Same in-place-edit safety as patch_workbook: reads everything into
    memory before opening output_path for writing."""
    zf = zipfile.ZipFile(original_path, "r")
    sheet_paths = {name: resolve_sheet_path(zf, name) for name in sheet_fixes}
    path_to_name = {path: name for name, path in sheet_paths.items()}

    items = [(item, zf.read(item.filename)) for item in zf.infolist()]
    zf.close()
    items = _drop_calc_chain(items)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in items:
            sheet_name = path_to_name.get(item.filename)
            if sheet_name:
                data = patch_existing_cells_fast(data, sheet_fixes[sheet_name])
            out.writestr(item, data)
    return sheet_paths


def set_full_calc_on_load(path):
    """Mark the workbook so Excel recalculates EVERY formula when it opens the file (<calcPr fullCalcOnLoad="1"/>).
    Without it the file keeps the template's old cached values until something forces a recalculation - on
    27 Oct 2026 a Working file went through the recalculation macro, was reported SUCCESS, and still showed the
    template's April figures. Excel clears the flag again on its next save. Only workbook.xml changes."""
    import re
    zf = zipfile.ZipFile(path, "r")
    items = [(i, zf.read(i.filename)) for i in zf.infolist()]
    zf.close()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        done = False
        for item, data in items:
            if item.filename == "xl/workbook.xml":
                text = data.decode("utf8")
                m = re.search(r"<calcPr\b[^>]*?/?>", text)
                if not m:
                    raise RuntimeError("workbook.xml has no <calcPr> element - cannot set fullCalcOnLoad")
                tag = m.group(0)
                if "fullCalcOnLoad" not in tag:
                    new = tag.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
                    text = text[:m.start()] + new + text[m.end():]
                data = text.encode("utf8")
                done = True
            out.writestr(item, data)
    if not done:
        raise RuntimeError("workbook.xml not found")