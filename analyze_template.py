"""
分析样例文件结构，输出 Excel builder 所需的完整坐标映射。
运行：python3 analyze_template.py > template_map.md
"""
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

SAMPLE = "/Users/gaoyiman/Desktop/Code/5000B/EstimateM/规模及工作量估计表.xlsx"

wb = openpyxl.load_workbook(SAMPLE, data_only=False)  # data_only=False 保留公式字符串

def col(idx):
    return get_column_letter(idx)

def cell_info(ws, r, c):
    cell = ws.cell(r, c)
    return {
        "addr": f"{col(c)}{r}",
        "value": cell.value,
        "font_bold": cell.font.bold if cell.font else None,
        "fill": str(cell.fill.fgColor.rgb) if cell.fill and cell.fill.fgColor else None,
        "alignment_h": cell.alignment.horizontal if cell.alignment else None,
    }

def dump_row(ws, row, max_col=24):
    cells = []
    for c in range(1, max_col + 1):
        v = ws.cell(row, c).value
        if v is not None:
            cells.append(f"  {col(c)}{row}: {repr(v)}")
    return cells

def merged_ranges(ws):
    return [str(m) for m in ws.merged_cells.ranges]

print("# EstimateM 样例文件结构分析")
print(f"\n文件：{SAMPLE}")
print(f"Sheet 列表：{wb.sheetnames}\n")

for sheet_idx, sheet_name in enumerate(wb.sheetnames, 1):
    ws = wb[sheet_name]
    max_r = ws.max_row
    max_c = ws.max_column
    print(f"\n{'='*70}")
    print(f"## Sheet {sheet_idx}：{sheet_name}")
    print(f"尺寸：{max_r} 行 × {max_c} 列")

    # 列宽
    print("\n### 列宽")
    for c in range(1, max_c + 1):
        cl = col(c)
        dim = ws.column_dimensions.get(cl)
        w = dim.width if dim else None
        if w:
            print(f"  {cl}: {w:.1f}")

    # 行高
    print("\n### 行高（非默认）")
    for r in range(1, max_r + 1):
        dim = ws.row_dimensions.get(r)
        if dim and dim.height:
            print(f"  行{r}: {dim.height:.1f}")

    # 合并单元格
    merges = merged_ranges(ws)
    print(f"\n### 合并单元格（共 {len(merges)} 个）")
    for m in merges:
        # 打印合并区的左上角值
        from openpyxl.utils.cell import range_boundaries
        min_col, min_row, max_col_m, max_row_m = range_boundaries(m)
        v = ws.cell(min_row, min_col).value
        print(f"  {m}  ← {repr(v)}")

    # 逐行内容
    print(f"\n### 单元格内容（全部 {max_r} 行）")
    for r in range(1, max_r + 1):
        cells = dump_row(ws, r, max_c)
        if cells:
            print(f"\n行 {r}:")
            for c in cells:
                print(c)

    # 公式汇总
    print(f"\n### 公式汇总")
    formula_cells = []
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("="):
                formula_cells.append(f"  {col(c)}{r}: {v}")
    if formula_cells:
        for f in formula_cells:
            print(f)
    else:
        print("  （无公式）")

    # 样式采样（表头行）
    print(f"\n### 样式采样（前5行非空单元格）")
    sampled = 0
    for r in range(1, min(10, max_r + 1)):
        for c in range(1, max_c + 1):
            cell = ws.cell(r, c)
            if cell.value is not None:
                fill_color = None
                if cell.fill and cell.fill.fgColor:
                    try:
                        fill_color = cell.fill.fgColor.rgb
                    except Exception:
                        pass
                font_color = None
                if cell.font and cell.font.color:
                    try:
                        font_color = cell.font.color.rgb
                    except Exception:
                        pass
                bold = cell.font.bold if cell.font else None
                halign = cell.alignment.horizontal if cell.alignment else None
                border_info = ""
                if cell.border:
                    sides = []
                    for side_name in ["left", "top", "right", "bottom"]:
                        side = getattr(cell.border, side_name, None)
                        if side and side.style:
                            sides.append(f"{side_name}={side.style}")
                    if sides:
                        border_info = " border=[" + ",".join(sides) + "]"
                print(f"  {col(c)}{r}: fill={fill_color} font_color={font_color} bold={bold} align={halign}{border_info}")
                sampled += 1
                if sampled >= 30:
                    break
        if sampled >= 30:
            break

print("\n\n# 分析完成")
