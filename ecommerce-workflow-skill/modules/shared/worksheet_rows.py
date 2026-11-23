"""为关键词和类目表增行，保留后续模块的格式和浮动对象位置。"""

from copy import copy

from openpyxl.utils.cell import coordinate_from_string
from openpyxl.worksheet.cell_range import CellRange


def shift_anchor_for_insert(anchor, insert_at: int, amount: int):
    """锚点使用零起始行号；字符串锚点需要由调用方重新赋值。"""
    if isinstance(anchor, str):
        column, row = coordinate_from_string(anchor)
        return f"{column}{row + amount if row >= insert_at else row}"
    for name in ("_from", "to"):
        marker = getattr(anchor, name, None)
        if marker is not None and marker.row + 1 >= insert_at:
            marker.row += amount
    return anchor


def insert_rows_preserving_merges(ws, insert_at: int, amount: int, template_row: int) -> None:
    """复制模板行样式，并整体平移原有行、合并区域和图片。"""
    if amount <= 0:
        return
    if insert_at < 1 or template_row < 1:
        raise ValueError("插入位置和模板行必须大于零")
    merges = [CellRange(str(area)) for area in ws.merged_cells.ranges]
    dimensions = {row: copy(dim) for row, dim in ws.row_dimensions.items()}
    # 解除合并会删除边缘单元格，必须先保存其边框、填充等样式。
    styles = {key: copy(cell._style) for key, cell in ws._cells.items()}
    template_styles = {col: copy(ws.cell(template_row, col)._style)
                       for col in range(1, ws.max_column + 1)}
    for area in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(area))
    ws.insert_rows(insert_at, amount)

    # 不只复制行高，还保留隐藏、分组、整行格式等属性。
    ws.row_dimensions.clear()
    for row, dim in dimensions.items():
        target = row + amount if row >= insert_at else row
        dim.index = target
        ws.row_dimensions[target] = dim
    for row in range(insert_at, insert_at + amount):
        if template_row in dimensions:
            dim = copy(dimensions[template_row])
            dim.index = row
            ws.row_dimensions[row] = dim
        for col, style in template_styles.items():
            ws.cell(row, col)._style = copy(style)

    shifted_styles = {}
    for (row, col), style in styles.items():
        target = row + amount if row >= insert_at else row
        shifted_styles[target, col] = style
        ws.cell(target, col)._style = copy(style)
    for area in merges:
        moved = copy(area)
        if area.min_row >= insert_at:
            moved.shift(row_shift=amount)
        elif area.max_row >= insert_at:
            moved.max_row += amount
        ws.merge_cells(str(moved))
    for row in range(insert_at, insert_at + amount):
        for area in merges:
            if area.min_row == area.max_row == template_row:
                candidate = CellRange(min_col=area.min_col, max_col=area.max_col,
                                      min_row=row, max_row=row)
                if not any(not candidate.isdisjoint(existing) for existing in ws.merged_cells.ranges):
                    ws.merge_cells(str(candidate))
        for col, style in template_styles.items():
            ws.cell(row, col)._style = copy(style)
    # 合并过程会重建边缘单元格，最后恢复其原样式。
    for (row, col), style in shifted_styles.items():
        ws.cell(row, col)._style = copy(style)
    for drawing in list(ws._images) + list(ws._charts):
        drawing.anchor = shift_anchor_for_insert(drawing.anchor, insert_at, amount)
