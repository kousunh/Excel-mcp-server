"""Apply formatting (font, color, borders, alignment) to Excel cells."""

import argparse
import json
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import (
    get_app, get_workbook, get_sheet, open_path,
    hex_to_rgb_int, output_json, IS_WINDOWS
)

# Excel alignment constants
H_ALIGN = {'left': -4131, 'center': -4108, 'right': -4152}
V_ALIGN = {'top': -4160, 'middle': -4108, 'bottom': -4107}

# Excel border constants
BORDER_POSITIONS = {
    'left': 7, 'right': 10, 'top': 8, 'bottom': 9,
    'inside_vertical': 11, 'inside_horizontal': 12
}
BORDER_STYLES = {
    'none': 0, 'thin': 1, 'medium': -4138, 'thick': 4,
    'double': -4119, 'dotted': -4118, 'dashed': -4115
}


def _parse_hex(hex_color):
    if hex_color.startswith('#'):
        hex_color = hex_color[1:]
    return (int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def _get_wb(workbook=None, path=None):
    if path:
        return open_path(path)
    app, err = get_app()
    if err:
        return None, False, err
    wb, err = get_workbook(app, workbook)
    return wb, False, err


def format_cells(workbook=None, path=None, cell_range=None, fmt=None, sheet=None):
    wb, was_opened, err = _get_wb(workbook, path)
    if err:
        return {"error": err}

    ws, err = get_sheet(wb, sheet)
    if err:
        return {"error": err}

    try:
        rng = ws.range(cell_range)
    except Exception as e:
        return {"error": f"Invalid range '{cell_range}': {e}"}

    try:
        # Cross-platform xlwings API
        if 'bold' in fmt:
            rng.font.bold = fmt['bold']
        if 'italic' in fmt:
            rng.font.italic = fmt['italic']
        if 'fontSize' in fmt:
            rng.font.size = fmt['fontSize']
        if 'fontName' in fmt:
            rng.font.name = fmt['fontName']
        if 'fontColor' in fmt:
            rng.font.color = _parse_hex(fmt['fontColor'])
        if 'backgroundColor' in fmt:
            rng.color = _parse_hex(fmt['backgroundColor'])
        if 'numberFormat' in fmt:
            rng.number_format = fmt['numberFormat']

        # Platform-specific: underline
        if 'underline' in fmt:
            val = 2 if fmt['underline'] else -4142
            try:
                if IS_WINDOWS:
                    rng.api.Font.Underline = val
                else:
                    rng.api.font_object.underline.set(val)
            except Exception:
                pass

        # Platform-specific: alignment
        h = fmt.get('textAlign')
        v = fmt.get('verticalAlign')
        wrap = fmt.get('wrapText')
        if h and h in H_ALIGN:
            try:
                if IS_WINDOWS:
                    rng.api.HorizontalAlignment = H_ALIGN[h]
                else:
                    rng.api.horizontal_alignment.set(H_ALIGN[h])
            except Exception:
                pass
        if v and v in V_ALIGN:
            try:
                if IS_WINDOWS:
                    rng.api.VerticalAlignment = V_ALIGN[v]
                else:
                    rng.api.vertical_alignment.set(V_ALIGN[v])
            except Exception:
                pass
        if wrap is not None:
            try:
                if IS_WINDOWS:
                    rng.api.WrapText = wrap
                else:
                    rng.api.wrap_text.set(wrap)
            except Exception:
                pass

        # Borders
        if 'borders' in fmt:
            _apply_borders(rng, fmt['borders'])

        wb.save()

        try:
            wb.app.screen_updating = True
        except Exception:
            pass

        return {"success": True, "workbook": wb.name, "sheet": ws.name, "range": cell_range}

    except Exception as e:
        return {"error": f"Failed to format: {e}"}

    finally:
        if was_opened:
            try:
                wb.close()
            except Exception:
                pass


def _apply_borders(rng, borders_config):
    for position, config in borders_config.items():
        if not config:
            continue
        style = config.get('style', 'thin')
        color = config.get('color', '#000000')
        rgb = hex_to_rgb_int(color)
        line_style = BORDER_STYLES.get(style, 1)

        if position == 'outside':
            positions = ['left', 'right', 'top', 'bottom']
        elif position == 'inside':
            positions = ['inside_vertical', 'inside_horizontal']
        elif position in BORDER_POSITIONS:
            positions = [position]
        else:
            continue

        for pos in positions:
            idx = BORDER_POSITIONS.get(pos)
            if idx is None:
                continue
            try:
                if IS_WINDOWS:
                    border = rng.api.Borders(idx)
                    border.LineStyle = line_style
                    if style != 'none':
                        border.Color = rgb
                else:
                    border = rng.api.borders[idx]
                    border.line_style.set(line_style)
                    if style != 'none':
                        border.color.set(rgb)
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workbook', default=None)
    parser.add_argument('--path', default=None)
    parser.add_argument('--range', required=True)
    parser.add_argument('--format', required=True)
    parser.add_argument('--sheet', default=None)
    args = parser.parse_args()

    if not args.workbook and not args.path:
        output_json({"error": "Either --workbook or --path is required"})
        return

    try:
        fmt = json.loads(args.format)
    except json.JSONDecodeError:
        output_json({"error": "Invalid JSON for format"})
        return

    result = format_cells(args.workbook, args.path, args.range, fmt, args.sheet)
    output_json(result)


if __name__ == "__main__":
    main()
