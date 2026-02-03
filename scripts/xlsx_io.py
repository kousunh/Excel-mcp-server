"""Pure Python .xlsx reader/writer using only stdlib.

Manipulates ZIP/XML directly, preserving images, charts, and all
non-modified content. Only the specific XML files that are changed
get re-serialized; everything else is passed through byte-for-byte.
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import re
import copy

# OOXML namespaces
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Register known namespaces to preserve prefixes on serialization
_KNOWN_NS = {
    '': NS,
    'r': NS_R,
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
    'xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
    'xr6': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision6',
    'xr10': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision10',
    'xr2': 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
}
for _p, _u in _KNOWN_NS.items():
    ET.register_namespace(_p, _u)


def _tag(name):
    return f'{{{NS}}}{name}'


# ---------------------------------------------------------------------------
# Cell reference utilities
# ---------------------------------------------------------------------------

def col_to_num(col_str):
    n = 0
    for c in col_str.upper():
        n = n * 26 + (ord(c) - ord('A') + 1)
    return n


def num_to_col(n):
    r = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        r = chr(rem + ord('A')) + r
    return r


def parse_cell_ref(ref):
    """Parse 'A1' -> (row, col)."""
    m = re.match(r'^([A-Z]+)(\d+)$', ref.upper().replace('$', ''))
    if not m:
        raise ValueError(f"Invalid cell reference: {ref}")
    return int(m.group(2)), col_to_num(m.group(1))


def parse_range(range_str):
    """Parse 'A1:C10' or 'A1' -> (min_col, min_row, max_col, max_row)."""
    parts = range_str.replace('$', '').upper().split(':')
    r1, c1 = parse_cell_ref(parts[0])
    if len(parts) == 2:
        r2, c2 = parse_cell_ref(parts[1])
    else:
        r2, c2 = r1, c1
    return c1, r1, c2, r2


def cell_ref(row, col):
    return f"{num_to_col(col)}{row}"


# ---------------------------------------------------------------------------
# XlsxFile
# ---------------------------------------------------------------------------

class XlsxFile:
    def __init__(self, path):
        self.path = os.path.abspath(path)
        self._entries = {}       # zip_path -> bytes
        self._compress = {}      # zip_path -> compress_type
        self._sheets = []        # [(name, zip_path), ...]
        self._shared_strings = []
        self._ss_modified = False
        self._sheet_trees = {}   # zip_path -> ET root
        self._styles_tree = None
        self._styles_modified = False
        self._modified_sheets = set()

    def open(self):
        self._register_ns_from_zip()
        with zipfile.ZipFile(self.path, 'r') as z:
            for info in z.infolist():
                self._entries[info.filename] = z.read(info.filename)
                self._compress[info.filename] = info.compress_type
        self._parse_workbook()
        self._parse_shared_strings()
        self._parse_styles()
        return self

    def save(self):
        # Serialize modified parts
        for sp in self._modified_sheets:
            if sp in self._sheet_trees:
                self._entries[sp] = _serialize(self._sheet_trees[sp])
        if self._ss_modified:
            self._serialize_ss()
        if self._styles_modified:
            self._entries['xl/styles.xml'] = _serialize(self._styles_tree)

        # Ensure sharedStrings.xml is in [Content_Types].xml
        if self._ss_modified:
            self._ensure_content_type('xl/sharedStrings.xml',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')

        tmp = self.path + '.tmp'
        with zipfile.ZipFile(tmp, 'w') as zout:
            for name, data in self._entries.items():
                info = zipfile.ZipInfo(name)
                info.compress_type = self._compress.get(name, zipfile.ZIP_DEFLATED)
                zout.writestr(info, data)
        os.replace(tmp, self.path)

    def close(self):
        self._entries.clear()
        self._sheet_trees.clear()

    # -- Sheet listing --

    @property
    def sheet_names(self):
        return [name for name, _ in self._sheets]

    def _sheet_path(self, name):
        for n, p in self._sheets:
            if n == name:
                return p
        raise ValueError(f"Sheet '{name}' not found")

    def _get_sheet_tree(self, name):
        sp = self._sheet_path(name)
        if sp not in self._sheet_trees:
            self._sheet_trees[sp] = _parse(self._entries[sp])
        return sp, self._sheet_trees[sp]

    # -- Reading values --

    def read_values(self, sheet_name, range_str):
        """Read a 2D list of values from a range."""
        _, tree = self._get_sheet_tree(sheet_name)
        c1, r1, c2, r2 = parse_range(range_str)

        # Index cells by (row, col) for fast lookup
        cells = {}
        for row_el in tree.iter(_tag('row')):
            rn = int(row_el.get('r'))
            if rn < r1 or rn > r2:
                continue
            for cell_el in row_el.iter(_tag('c')):
                ref = cell_el.get('r', '')
                try:
                    cr, cc = parse_cell_ref(ref)
                except ValueError:
                    continue
                if c1 <= cc <= c2:
                    cells[(cr, cc)] = self._cell_value(cell_el)

        return [[cells.get((r, c)) for c in range(c1, c2 + 1)]
                for r in range(r1, r2 + 1)]

    def _cell_value(self, cell_el):
        t = cell_el.get('t', '')
        v_el = cell_el.find(_tag('v'))
        is_el = cell_el.find(_tag('is'))  # inline string

        if t == 's' and v_el is not None:
            idx = int(v_el.text)
            return self._shared_strings[idx] if idx < len(self._shared_strings) else None
        elif t == 'inlineStr' and is_el is not None:
            return _inline_text(is_el)
        elif t == 'b' and v_el is not None:
            return v_el.text == '1'
        elif t == 'e':
            return f"#ERROR:{v_el.text}" if v_el is not None else None
        elif v_el is not None and v_el.text is not None:
            # Number (or date stored as number)
            try:
                fv = float(v_el.text)
                return int(fv) if fv == int(fv) else fv
            except (ValueError, OverflowError):
                return v_el.text
        return None

    # -- Writing values --

    def write_values(self, sheet_name, range_str, values_2d):
        """Write a 2D list of values to a range."""
        sp, tree = self._get_sheet_tree(sheet_name)
        c1, r1, c2, r2 = parse_range(range_str)
        sheet_data = tree.find(_tag('sheetData'))
        if sheet_data is None:
            sheet_data = ET.SubElement(tree, _tag('sheetData'))

        # Index existing rows
        row_map = {}
        for row_el in sheet_data.findall(_tag('row')):
            row_map[int(row_el.get('r'))] = row_el

        for ri, row_vals in enumerate(values_2d):
            rn = r1 + ri
            if rn > r2:
                break
            row_el = row_map.get(rn)
            if row_el is None:
                row_el = ET.SubElement(sheet_data, _tag('row'))
                row_el.set('r', str(rn))
                row_map[rn] = row_el

            # Index existing cells in this row
            cell_map = {}
            for c_el in row_el.findall(_tag('c')):
                try:
                    _, cc = parse_cell_ref(c_el.get('r', ''))
                    cell_map[cc] = c_el
                except ValueError:
                    pass

            for ci, val in enumerate(row_vals if isinstance(row_vals, list) else [row_vals]):
                cn = c1 + ci
                if cn > c2:
                    break
                ref = cell_ref(rn, cn)
                c_el = cell_map.get(cn)
                if c_el is None:
                    c_el = ET.SubElement(row_el, _tag('c'))
                    c_el.set('r', ref)

                self._set_cell_value(c_el, val)

        self._modified_sheets.add(sp)

    def _set_cell_value(self, c_el, val):
        # Remove formula if present
        f_el = c_el.find(_tag('f'))
        if f_el is not None:
            c_el.remove(f_el)

        v_el = c_el.find(_tag('v'))

        if val is None:
            if v_el is not None:
                c_el.remove(v_el)
            c_el.attrib.pop('t', None)
            return

        if v_el is None:
            v_el = ET.SubElement(c_el, _tag('v'))

        if isinstance(val, bool):
            v_el.text = '1' if val else '0'
            c_el.set('t', 'b')
        elif isinstance(val, (int, float)):
            v_el.text = str(val)
            c_el.attrib.pop('t', None)  # number is default type
        else:
            # String -> add to shared strings
            idx = self._add_shared_string(str(val))
            v_el.text = str(idx)
            c_el.set('t', 's')

    # -- Reading formats --

    def read_formats(self, sheet_name, range_str):
        """Read formatting info for cells with non-default formatting."""
        _, tree = self._get_sheet_tree(sheet_name)
        c1, r1, c2, r2 = parse_range(range_str)
        formats = []

        for row_el in tree.iter(_tag('row')):
            rn = int(row_el.get('r'))
            if rn < r1 or rn > r2:
                continue
            for cell_el in row_el.iter(_tag('c')):
                ref = cell_el.get('r', '')
                try:
                    cr, cc = parse_cell_ref(ref)
                except ValueError:
                    continue
                if cc < c1 or cc > c2:
                    continue

                s_idx = int(cell_el.get('s', '0'))
                if s_idx == 0:
                    continue  # default style

                fmt = self._xf_to_fmt(s_idx)
                if fmt:
                    fmt['cell'] = ref.replace('$', '')
                    formats.append(fmt)

        return formats

    def _xf_to_fmt(self, xf_idx):
        """Convert cellXf index to our format dict."""
        if self._styles_tree is None:
            return {}
        xfs = self._styles_tree.find(_tag('cellXfs'))
        if xfs is None:
            return {}
        xf_list = xfs.findall(_tag('xf'))
        if xf_idx >= len(xf_list):
            return {}
        xf = xf_list[xf_idx]

        fmt = {}
        # Font
        font_id = int(xf.get('fontId', '0'))
        if xf.get('applyFont') == '1' or font_id > 0:
            self._read_font(font_id, fmt)

        # Fill
        fill_id = int(xf.get('fillId', '0'))
        if fill_id > 1:  # 0=none, 1=gray125 (default)
            self._read_fill(fill_id, fmt)

        # Border
        border_id = int(xf.get('borderId', '0'))
        if border_id > 0:
            self._read_border(border_id, fmt)

        # Number format
        num_fmt_id = int(xf.get('numFmtId', '0'))
        if num_fmt_id > 0:
            nf = self._num_fmt_code(num_fmt_id)
            if nf and nf != 'General':
                fmt['numberFormat'] = nf

        # Alignment
        align_el = xf.find(_tag('alignment'))
        if align_el is not None:
            h = align_el.get('horizontal')
            v = align_el.get('vertical')
            w = align_el.get('wrapText')
            if h and h != 'general':
                fmt['textAlign'] = h
            if v and v != 'bottom':
                fmt['verticalAlign'] = v
            if w == '1':
                fmt['wrapText'] = True

        return fmt

    def _read_font(self, font_id, fmt):
        fonts = self._styles_tree.find(_tag('fonts'))
        if fonts is None:
            return
        font_list = fonts.findall(_tag('font'))
        if font_id >= len(font_list):
            return
        font = font_list[font_id]

        if font.find(_tag('b')) is not None:
            fmt['bold'] = True
        if font.find(_tag('i')) is not None:
            fmt['italic'] = True
        if font.find(_tag('u')) is not None:
            fmt['underline'] = True
        sz = font.find(_tag('sz'))
        if sz is not None:
            fmt['fontSize'] = float(sz.get('val', '11'))
        name = font.find(_tag('name'))
        if name is not None:
            fmt['fontName'] = name.get('val', '')
        color = font.find(_tag('color'))
        if color is not None:
            rgb = color.get('rgb', '')
            if rgb and len(rgb) == 8:
                rgb = rgb[2:]
            if rgb and rgb.lower() != '000000':
                fmt['fontColor'] = f'#{rgb.lower()}'

    def _read_fill(self, fill_id, fmt):
        fills = self._styles_tree.find(_tag('fills'))
        if fills is None:
            return
        fill_list = fills.findall(_tag('fill'))
        if fill_id >= len(fill_list):
            return
        pf = fill_list[fill_id].find(_tag('patternFill'))
        if pf is None:
            return
        fg = pf.find(_tag('fgColor'))
        if fg is not None:
            rgb = fg.get('rgb', '')
            if rgb and len(rgb) == 8:
                rgb = rgb[2:]
            if rgb and rgb.lower() != '000000':
                fmt['bg'] = f'#{rgb.lower()}'

    def _read_border(self, border_id, fmt):
        borders_el = self._styles_tree.find(_tag('borders'))
        if borders_el is None:
            return
        border_list = borders_el.findall(_tag('border'))
        if border_id >= len(border_list):
            return
        border = border_list[border_id]
        borders = {}
        for side in ('left', 'right', 'top', 'bottom'):
            el = border.find(_tag(side))
            if el is not None:
                style = el.get('style')
                if style and style != 'none':
                    borders[side] = style
        if borders:
            fmt['borders'] = borders

    def _num_fmt_code(self, num_fmt_id):
        # Built-in formats
        builtin = {
            0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
            9: '0%', 10: '0.00%', 11: '0.00E+00', 14: 'mm-dd-yy',
            22: 'm/d/yy h:mm'
        }
        if num_fmt_id in builtin:
            return builtin[num_fmt_id]
        # Custom formats
        nfs = self._styles_tree.find(_tag('numFmts'))
        if nfs is not None:
            for nf in nfs.findall(_tag('numFmt')):
                if int(nf.get('numFmtId', '0')) == num_fmt_id:
                    return nf.get('formatCode', '')
        return None

    # -- Writing formats --

    def apply_format(self, sheet_name, range_str, fmt):
        """Apply formatting to a range of cells."""
        sp, tree = self._get_sheet_tree(sheet_name)
        c1, r1, c2, r2 = parse_range(range_str)
        sheet_data = tree.find(_tag('sheetData'))
        if sheet_data is None:
            return

        # Ensure rows and cells exist for the range
        row_map = {}
        for row_el in sheet_data.findall(_tag('row')):
            row_map[int(row_el.get('r'))] = row_el

        # Cache: old_xf_idx -> new_xf_idx
        xf_cache = {}

        for rn in range(r1, r2 + 1):
            row_el = row_map.get(rn)
            if row_el is None:
                row_el = ET.SubElement(sheet_data, _tag('row'))
                row_el.set('r', str(rn))
                row_map[rn] = row_el

            cell_map = {}
            for c_el in row_el.findall(_tag('c')):
                try:
                    _, cc = parse_cell_ref(c_el.get('r', ''))
                    cell_map[cc] = c_el
                except ValueError:
                    pass

            for cn in range(c1, c2 + 1):
                c_el = cell_map.get(cn)
                if c_el is None:
                    c_el = ET.SubElement(row_el, _tag('c'))
                    c_el.set('r', cell_ref(rn, cn))

                old_xf = int(c_el.get('s', '0'))
                if old_xf not in xf_cache:
                    xf_cache[old_xf] = self._build_xf(old_xf, fmt)
                c_el.set('s', str(xf_cache[old_xf]))

        self._modified_sheets.add(sp)
        self._styles_modified = True

    def _build_xf(self, base_xf_idx, fmt):
        """Create a new cellXf by merging base style with new format properties."""
        xfs = self._styles_tree.find(_tag('cellXfs'))
        if xfs is None:
            xfs = ET.SubElement(self._styles_tree, _tag('cellXfs'))
            xfs.set('count', '1')
            default_xf = ET.SubElement(xfs, _tag('xf'))
            default_xf.set('numFmtId', '0')
            default_xf.set('fontId', '0')
            default_xf.set('fillId', '0')
            default_xf.set('borderId', '0')

        xf_list = xfs.findall(_tag('xf'))
        base = xf_list[base_xf_idx] if base_xf_idx < len(xf_list) else xf_list[0]

        font_id = int(base.get('fontId', '0'))
        fill_id = int(base.get('fillId', '0'))
        border_id = int(base.get('borderId', '0'))
        num_fmt_id = int(base.get('numFmtId', '0'))
        base_align = base.find(_tag('alignment'))

        # Font
        if any(k in fmt for k in ('bold', 'italic', 'underline', 'fontSize', 'fontName', 'fontColor')):
            font_id = self._merge_font(font_id, fmt)

        # Fill
        if 'backgroundColor' in fmt:
            fill_id = self._make_fill(fmt['backgroundColor'])

        # Border
        if 'borders' in fmt:
            border_id = self._merge_border(border_id, fmt['borders'])

        # Number format
        if 'numberFormat' in fmt:
            num_fmt_id = self._get_num_fmt_id(fmt['numberFormat'])

        # Alignment
        align_props = {}
        if base_align is not None:
            for attr in ('horizontal', 'vertical', 'wrapText'):
                v = base_align.get(attr)
                if v:
                    align_props[attr] = v
        if 'textAlign' in fmt:
            align_props['horizontal'] = fmt['textAlign']
        if 'verticalAlign' in fmt:
            align_props['vertical'] = fmt['verticalAlign']
        if 'wrapText' in fmt:
            align_props['wrapText'] = '1' if fmt['wrapText'] else '0'

        # Find or create matching xf
        return self._find_or_add_xf(font_id, fill_id, border_id, num_fmt_id, align_props)

    def _merge_font(self, base_font_id, fmt):
        """Create a new font by merging base font with new properties."""
        fonts_el = self._styles_tree.find(_tag('fonts'))
        base_list = fonts_el.findall(_tag('font'))
        base_font = base_list[base_font_id] if base_font_id < len(base_list) else base_list[0]

        new_font = copy.deepcopy(base_font)

        def _set_flag(tag_name, key):
            el = new_font.find(_tag(tag_name))
            if key in fmt:
                if fmt[key]:
                    if el is None:
                        ET.SubElement(new_font, _tag(tag_name))
                else:
                    if el is not None:
                        new_font.remove(el)

        _set_flag('b', 'bold')
        _set_flag('i', 'italic')
        _set_flag('u', 'underline')

        if 'fontSize' in fmt:
            sz = new_font.find(_tag('sz'))
            if sz is None:
                sz = ET.SubElement(new_font, _tag('sz'))
            sz.set('val', str(fmt['fontSize']))

        if 'fontName' in fmt:
            name = new_font.find(_tag('name'))
            if name is None:
                name = ET.SubElement(new_font, _tag('name'))
            name.set('val', fmt['fontName'])

        if 'fontColor' in fmt:
            color = new_font.find(_tag('color'))
            if color is None:
                color = ET.SubElement(new_font, _tag('color'))
            hex_c = fmt['fontColor'].lstrip('#')
            color.set('rgb', f'FF{hex_c.upper()}')

        # Add font and return index
        fonts_el.append(new_font)
        idx = len(fonts_el.findall(_tag('font'))) - 1
        fonts_el.set('count', str(idx + 1))
        self._styles_modified = True
        return idx

    def _make_fill(self, bg_color):
        fills_el = self._styles_tree.find(_tag('fills'))
        hex_c = bg_color.lstrip('#').upper()

        fill = ET.SubElement(fills_el, _tag('fill'))
        pf = ET.SubElement(fill, _tag('patternFill'))
        pf.set('patternType', 'solid')
        fg = ET.SubElement(pf, _tag('fgColor'))
        fg.set('rgb', f'FF{hex_c}')
        bg = ET.SubElement(pf, _tag('bgColor'))
        bg.set('indexed', '64')

        idx = len(fills_el.findall(_tag('fill'))) - 1
        fills_el.set('count', str(idx + 1))
        self._styles_modified = True
        return idx

    def _merge_border(self, base_border_id, borders_fmt):
        borders_el = self._styles_tree.find(_tag('borders'))
        base_list = borders_el.findall(_tag('border'))
        base = base_list[base_border_id] if base_border_id < len(base_list) else base_list[0]

        new_border = copy.deepcopy(base)

        def _set_side(side_name, config):
            el = new_border.find(_tag(side_name))
            if el is None:
                el = ET.SubElement(new_border, _tag(side_name))
            style = config.get('style', 'thin')
            el.set('style', style)
            color_el = el.find(_tag('color'))
            if color_el is None:
                color_el = ET.SubElement(el, _tag('color'))
            hex_c = config.get('color', '#000000').lstrip('#').upper()
            color_el.set('rgb', f'FF{hex_c}')

        for position, config in borders_fmt.items():
            if not config:
                continue
            if position == 'outside':
                for s in ('left', 'right', 'top', 'bottom'):
                    _set_side(s, config)
            elif position == 'inside':
                for s in ('left', 'right', 'top', 'bottom'):
                    _set_side(s, config)
            elif position in ('left', 'right', 'top', 'bottom'):
                _set_side(position, config)

        borders_el.append(new_border)
        idx = len(borders_el.findall(_tag('border'))) - 1
        borders_el.set('count', str(idx + 1))
        self._styles_modified = True
        return idx

    def _get_num_fmt_id(self, format_code):
        nfs = self._styles_tree.find(_tag('numFmts'))
        if nfs is None:
            nfs = ET.SubElement(self._styles_tree, _tag('numFmts'))
            # Insert numFmts at beginning of styleSheet
            self._styles_tree.insert(0, nfs)
            nfs.set('count', '0')

        # Check if already exists
        for nf in nfs.findall(_tag('numFmt')):
            if nf.get('formatCode') == format_code:
                return int(nf.get('numFmtId'))

        # Find next available ID (custom IDs start at 164)
        max_id = 163
        for nf in nfs.findall(_tag('numFmt')):
            fid = int(nf.get('numFmtId', '0'))
            if fid > max_id:
                max_id = fid

        new_id = max_id + 1
        new_nf = ET.SubElement(nfs, _tag('numFmt'))
        new_nf.set('numFmtId', str(new_id))
        new_nf.set('formatCode', format_code)
        nfs.set('count', str(len(nfs.findall(_tag('numFmt')))))
        self._styles_modified = True
        return new_id

    def _find_or_add_xf(self, font_id, fill_id, border_id, num_fmt_id, align_props):
        xfs = self._styles_tree.find(_tag('cellXfs'))
        xf_list = xfs.findall(_tag('xf'))

        # Check if matching xf already exists
        for i, xf in enumerate(xf_list):
            if (int(xf.get('fontId', '0')) == font_id and
                int(xf.get('fillId', '0')) == fill_id and
                int(xf.get('borderId', '0')) == border_id and
                int(xf.get('numFmtId', '0')) == num_fmt_id):
                # Check alignment
                a_el = xf.find(_tag('alignment'))
                existing = {}
                if a_el is not None:
                    for attr in ('horizontal', 'vertical', 'wrapText'):
                        v = a_el.get(attr)
                        if v:
                            existing[attr] = v
                if existing == align_props:
                    return i

        # Create new xf
        new_xf = ET.SubElement(xfs, _tag('xf'))
        new_xf.set('numFmtId', str(num_fmt_id))
        new_xf.set('fontId', str(font_id))
        new_xf.set('fillId', str(fill_id))
        new_xf.set('borderId', str(border_id))
        if font_id > 0:
            new_xf.set('applyFont', '1')
        if fill_id > 0:
            new_xf.set('applyFill', '1')
        if border_id > 0:
            new_xf.set('applyBorder', '1')
        if num_fmt_id > 0:
            new_xf.set('applyNumberFormat', '1')
        if align_props:
            new_xf.set('applyAlignment', '1')
            a = ET.SubElement(new_xf, _tag('alignment'))
            for k, v in align_props.items():
                a.set(k, v)

        idx = len(xfs.findall(_tag('xf'))) - 1
        xfs.set('count', str(idx + 1))
        return idx

    # -- Internal helpers --

    def _register_ns_from_zip(self):
        """Extract and register namespace prefixes from XML files in the ZIP."""
        try:
            with zipfile.ZipFile(self.path, 'r') as z:
                for name in z.namelist():
                    if name.endswith('.xml') or name.endswith('.rels'):
                        data = z.read(name)
                        for m in re.finditer(rb'xmlns:(\w+)=["\']([^"\']+)["\']', data):
                            prefix = m.group(1).decode('utf-8')
                            uri = m.group(2).decode('utf-8')
                            try:
                                ET.register_namespace(prefix, uri)
                            except Exception:
                                pass
        except Exception:
            pass

    def _parse_workbook(self):
        wb_tree = _parse(self._entries.get('xl/workbook.xml', b''))
        rels_data = self._entries.get('xl/_rels/workbook.xml.rels', b'')
        rels_tree = _parse(rels_data)

        rid_map = {}
        for rel in rels_tree.iter(f'{{{NS_REL}}}Relationship'):
            rid_map[rel.get('Id')] = rel.get('Target')

        for sheet in wb_tree.iter(_tag('sheet')):
            name = sheet.get('name')
            rid = sheet.get(f'{{{NS_R}}}id')
            target = rid_map.get(rid, '')
            if not target.startswith('/'):
                sp = f'xl/{target}'
            else:
                sp = target[1:]
            self._sheets.append((name, sp))

    def _parse_shared_strings(self):
        data = self._entries.get('xl/sharedStrings.xml')
        if not data:
            self._shared_strings = []
            return
        tree = _parse(data)
        self._shared_strings = []
        for si in tree.iter(_tag('si')):
            self._shared_strings.append(_inline_text(si))

    def _add_shared_string(self, s):
        # Check if already exists
        try:
            return self._shared_strings.index(s)
        except ValueError:
            pass
        self._shared_strings.append(s)
        self._ss_modified = True
        return len(self._shared_strings) - 1

    def _serialize_ss(self):
        root = ET.Element(_tag('sst'))
        root.set('count', str(len(self._shared_strings)))
        root.set('uniqueCount', str(len(self._shared_strings)))
        for s in self._shared_strings:
            si = ET.SubElement(root, _tag('si'))
            t = ET.SubElement(si, _tag('t'))
            if s and (s[0] == ' ' or s[-1] == ' '):
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = s
        self._entries['xl/sharedStrings.xml'] = _serialize(root)

    def _parse_styles(self):
        data = self._entries.get('xl/styles.xml')
        if data:
            self._styles_tree = _parse(data)
        else:
            # Create minimal styles
            self._styles_tree = ET.Element(_tag('styleSheet'))
            for coll in ('fonts', 'fills', 'borders', 'cellXfs'):
                el = ET.SubElement(self._styles_tree, _tag(coll))
                el.set('count', '1')
                if coll == 'fonts':
                    font = ET.SubElement(el, _tag('font'))
                    ET.SubElement(font, _tag('sz')).set('val', '11')
                    ET.SubElement(font, _tag('name')).set('val', 'Calibri')
                elif coll == 'fills':
                    ET.SubElement(ET.SubElement(el, _tag('fill')),
                                  _tag('patternFill')).set('patternType', 'none')
                elif coll == 'borders':
                    ET.SubElement(el, _tag('border'))
                elif coll == 'cellXfs':
                    xf = ET.SubElement(el, _tag('xf'))
                    for attr in ('numFmtId', 'fontId', 'fillId', 'borderId'):
                        xf.set(attr, '0')

    def _ensure_content_type(self, part_name, content_type):
        """Ensure a part is registered in [Content_Types].xml."""
        ct_data = self._entries.get('[Content_Types].xml')
        if not ct_data:
            return
        ns_ct = 'http://schemas.openxmlformats.org/package/2006/content-types'
        ET.register_namespace('', ns_ct)
        tree = _parse(ct_data)
        # Check if Override already exists
        for ov in tree.findall(f'{{{ns_ct}}}Override'):
            if ov.get('PartName') == f'/{part_name}':
                return
        # Add it
        ov = ET.SubElement(tree, f'{{{ns_ct}}}Override')
        ov.set('PartName', f'/{part_name}')
        ov.set('ContentType', content_type)
        self._entries['[Content_Types].xml'] = _serialize(tree)


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------

def _parse(data):
    """Parse XML bytes into ElementTree root."""
    if isinstance(data, bytes):
        return ET.fromstring(data)
    return ET.fromstring(data.encode('utf-8'))


def _serialize(root):
    """Serialize ElementTree root to bytes with XML declaration."""
    xml_str = ET.tostring(root, encoding='unicode', xml_declaration=False)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml_str).encode('utf-8')


def _inline_text(si_or_is_el):
    """Extract text from a <si> or <is> element (handles rich text)."""
    t_el = si_or_is_el.find(_tag('t'))
    if t_el is not None:
        return t_el.text or ''
    # Rich text: concatenate all <r><t> elements
    parts = []
    for r in si_or_is_el.iter(_tag('r')):
        t = r.find(_tag('t'))
        if t is not None and t.text:
            parts.append(t.text)
    return ''.join(parts) if parts else ''
