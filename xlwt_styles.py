import xlwt3 as xlwt

### Описание стиля 1(horizontal)
# перенос по словам, выравнивание
alignment = xlwt.Alignment()
alignment.wrap = 1
alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT,
# HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
alignment.vert = xlwt.Alignment.VERT_CENTER  # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
# шрифт
font = xlwt.Font()
font.name = 'Arial Cyr'
font.bold = True
# границы
borders = xlwt.Borders()
borders.left = xlwt.Borders.THIN  # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
borders.right = xlwt.Borders.THIN
borders.top = xlwt.Borders.THIN
borders.bottom = xlwt.Borders.THIN
# Создаём стиль с нашими установками
style = xlwt.XFStyle()
style.font = font
style.alignment = alignment
style.borders = borders

### Описание стиля 2(vertical)
# перенос по словам, выравнивание
alignment_v = xlwt.Alignment()
alignment_v.wrap = 1
alignment_v.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
alignment_v.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
alignment_v.rota = 90
# шрифт
font_v = xlwt.Font()
font_v.name = 'Arial Cyr'
font_v.bold = True
# границы
borders_v = xlwt.Borders()
borders_v.left = xlwt.Borders.THIN # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
borders_v.right = xlwt.Borders.THIN
borders_v.top = xlwt.Borders.THIN
borders_v.bottom = xlwt.Borders.THIN
# Создаём стиль с нашими установками
style_v = xlwt.XFStyle()
style_v.font = font_v
style_v.alignment = alignment_v
style_v.borders = borders_v

### Описание стиля 3(Editable cells)
np_style = xlwt.easyxf("protection: cell_locked false")