import xlwt

def set_style(FontName='等线',FontHeight = 11,bold=False,FontColor='black',bgColor='white'):
  style = xlwt.XFStyle()  # 赋值style为XFStyle()，初始化样式

  font = xlwt.Font()  # 为样式创建字体
  font.name = FontName  # 'Times New Roman'
  font.colour_index =  xlwt.Style.colour_map[FontColor]
  font.bold = bold
  font.height = 20*FontHeight


  borders = xlwt.Borders()
  borders.left = 1
  borders.top = 1
  borders.right = 1
  borders.bottom = 1
  borders.left_colour = 0x01  # 边框上色
  borders.right_colour = 0x01
  borders.top_colour = 0x01
  borders.bottom_colour = 0x01

  # 设置居中
  al = xlwt.Alignment()
  al.horz = 0x02  # 设置水平居中
  al.vert = 0x01  # 设置垂直居中
  al.wrap = 1

  pattern = xlwt.Pattern()
  pattern.pattern = xlwt.Pattern.SOLID_PATTERN
  pattern.pattern_fore_colour = xlwt.Style.colour_map[bgColor]

  style.pattern = pattern
  style.borders = borders
  style.alignment = al
  style.font = font
  return style