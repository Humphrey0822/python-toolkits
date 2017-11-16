# coding=utf-8
# 设置Excel表格样式
import xlwt


def set_style(name, height, bold=False, back=False):
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    borders = xlwt.Borders()  # 设置边框
    borders.left = xlwt.Borders.THIN  # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    if back:
        patterni = xlwt.Pattern()  # 为样式创建图案
        patterni.pattern = 2  # 设置底纹的图案索引，1为实心，2为50%灰色，对应为excel文件单元格格式中填充中的图案样式
        patterni.pattern_fore_colour = 0x16  # 设置底纹的前景色，对应为excel文件单元格格式中填充中的背景色
        patterni.pattern_back_colour = 0x16  # 设置底纹的背景色，对应为excel文件单元格格式中填充中的图案颜色
        style.pattern = patterni  # 为样式设置图案
    return style
