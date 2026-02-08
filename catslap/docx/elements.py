# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)

from io import BytesIO

from PIL import Image

from catslap.base.relationships import Relationships
from catslap.base.types import ContentTypes
from catslap.docx.styles import Styles
from catslap.utils import html
from catslap.utils import text as text_util
from catslap.utils.xml import XmlParserException, XmlTag, XmlParser
from catslap.docx import word_tags as WT

SIZE_TWIPS_PER_PX = 20
SIZE_TWIPS_PER_CM = 567
SIZE_WIDTH_CM = 17
SIZE_HEIGHT_CM = 24


def __get_tag_value_bool(tag, tag_name):
  tag_value = tag.get_tag(tag_name, False)
  if tag_value is None:
    return False
  value = tag_value.get_attr(WT.ATTR_VAL)
  if value is None:
    return True
  return value != '0'

def create_run(r_tag: XmlTag, text: str, runprops: dict | None, relationships: Relationships, types: ContentTypes, styles: Styles) -> XmlTag:
  """
  Creates a Word run from properties and text.

  Args:
    r_tag: Reference base tag.
    text: Text to insert.
    runprops: Style properties.
    relationships: Document relationships.
    types: Document ContentTypes.
    styles: Document styles.

  Returns:
    XmlTag representing the run or a hyperlink.

  Raises:
    XmlParserException: If image data is invalid.
  """
  image = runprops.get('image')
  if image and image.startswith('data:'):
    try:
      mediatype, encoding, data = html.extract_image_data(image)
    except html.HtmlException as e:
      raise XmlParserException(str(e))
    wd = runprops.get('width')
    hg = runprops.get('height')
    return create_image(mediatype, data, wd, hg, relationships, types)

  bold = runprops.get('bold') is True
  italic = runprops.get('italic') is True
  strike = runprops.get('strike') is True
  underline = runprops.get('underline') is True
  color = runprops.get('color')
  style = runprops.get('style')
  link = runprops.get('link')
  if link:
    style = styles.style_map.get(Styles.CFG_STYLE_LINK_URL)
  code = runprops.get('code')
  if code:
    style = styles.style_map.get(Styles.CFG_STYLE_CODE)

  rpr_tag = r_tag.get_tag(WT.TAG_RPR, False)
  if rpr_tag:
    out_rpr_tag = rpr_tag.clone(True)
    out_rpr_tag.remove_tag(WT.TAG_BOLD)
    out_rpr_tag.remove_tag(WT.TAG_BOLD_X)
    out_rpr_tag.remove_tag(WT.TAG_ITALIC)
    out_rpr_tag.remove_tag(WT.TAG_ITALIC_X)
    out_rpr_tag.remove_tag(WT.TAG_STRIKE)
    out_rpr_tag.remove_tag(WT.TAG_UNDERLINE)
    out_rpr_tag.remove_tag(WT.TAG_R_STYLE)
    out_rpr_tag.remove_tag(WT.TAG_COLOR)
    bold = bold or __get_tag_value_bool(rpr_tag, WT.TAG_BOLD)
    italic = italic or __get_tag_value_bool(rpr_tag, WT.TAG_ITALIC)
    strike = strike or __get_tag_value_bool(rpr_tag, WT.TAG_STRIKE)
    underline = underline or __get_tag_value_bool(rpr_tag, WT.TAG_UNDERLINE)
    style = style if style is not None else rpr_tag.get_tag_attr(WT.TAG_R_STYLE, WT.ATTR_VAL, False)
    color = color if color is not None else rpr_tag.get_tag_attr(WT.TAG_COLOR, WT.ATTR_VAL, False)
  else:
    out_rpr_tag = XmlTag(WT.TAG_RPR)

  out_r_tag = XmlTag(WT.TAG_R)
  out_r_tag.add_tag(out_rpr_tag)
  if bold:
    out_rpr_tag.add_tag(XmlTag(WT.TAG_BOLD))
    out_rpr_tag.add_tag(XmlTag(WT.TAG_BOLD_X))
  if italic:
    out_rpr_tag.add_tag(XmlTag(WT.TAG_ITALIC))
    out_rpr_tag.add_tag(XmlTag(WT.TAG_ITALIC_X))
  if underline:
    out_rpr_tag.add_tag(XmlTag(WT.TAG_UNDERLINE, {WT.ATTR_VAL: WT.ATTR_VAL_UNDERLINE}))
  if strike:
    out_rpr_tag.add_tag(XmlTag(WT.TAG_STRIKE))        
  if style:
    out_rpr_tag.add_tag(XmlTag(WT.TAG_R_STYLE, {WT.ATTR_VAL: style}))
  if color:
    out_rpr_tag.add_tag(XmlTag('w:color', {WT.ATTR_VAL: get_color(color)}))

  if text is None:
    text = ''
  out_t_tag = out_r_tag.add_tag(XmlTag('w:t', {'xml:space': 'preserve'}))
  out_t_tag.add_text(XmlParser.escape_entities(text))
  if link:
    relationship = relationships.add_relationship_hyperlink(link)
    hyper_tag = XmlTag(WT.TAG_HYPERLINK, {WT.ATTR_ID: relationship.rid, 'w:history': '1'})
    hyper_tag.add_tag(out_r_tag)
    return hyper_tag
  return out_r_tag

def create_image(mediatype: str, data: bytes, pxwd: int|None, pxhg: int|None, relationships: Relationships, types: ContentTypes) -> XmlTag:
  """
  Creates a run with an embedded image.

  Args:
    mediatype: Image MIME type.
    data: Image bytes.
    pxwd: Width in pixels (optional).
    pxhg: Height in pixels (optional).
    relationships: Document relationships.
    types: Document ContentTypes.

  Returns:
    XmlTag with image content.

  Raises:
    OSError: If the image cannot be written.
  """
  image_ext = mediatype[mediatype.find('/') + 1:]
  image_ref = 'image' + str(relationships.max_id + 1) + '.' + image_ext
  relationship = relationships.add_relationship_image(image_ref)
  types.add_default(image_ext, 'image/' + image_ext)
  rid = relationship.rid
  num = relationships.max_id * 2
  relationships.add_image(image_ref, data)
  if not pxwd or not pxhg:
    stream = BytesIO(data)
    img = Image.open(stream).convert("RGBA")
    stream.close()
    if pxwd and not pxhg:
      pxhg = pxwd * img.height / img.width
    elif pxhg and not pxwd:
      pxwd = pxhg * img.width / img.height
    if not pxwd:
      pxwd = img.width
    if not pxhg:
      pxhg = img.height
  dpi_wd = int(round(pxwd * SIZE_TWIPS_PER_PX))
  dpi_hg = int(round(pxhg * SIZE_TWIPS_PER_PX))

  run_tag = XmlTag('w:r')
  rpr_tag = run_tag.add_tag('w:rPr')
  rpr_tag.add_tag(XmlTag('w:rFonts', {'w:cstheme': 'minorHAnsi'}))
  rpr_tag.add_tag(XmlTag('w:noProof'))
  drawing_tag = run_tag.add_tag('w:drawing')
  inline_tag = drawing_tag.add_tag(XmlTag('wp:inline', {'distT': '0', 'distB': '0', 'distL': '0', 'distR': '0',
                                                        'wp14:anchorId': '27F8AE68', 'wp14:editId': '2DB58A57'}))
  inline_tag.add_tag(XmlTag('wp:extent', {'cx': dpi_wd, 'cy': dpi_hg}))
  inline_tag.add_tag(XmlTag('wp:effectExtent', {'l': '0', 't': '0', 'r': '0', 'b': '0'}))
  inline_tag.add_tag(XmlTag('wp:docPr', {'id': num+1, 'name': 'Imagen 19', 'descr': 'Icono&#xA;&#xA;Descripción generada automáticamente'}))
  cnv_gr_tag = inline_tag.add_tag(XmlTag('wp:cNvGraphicFramePr'))
  cnv_gr_tag.add_tag(XmlTag('a:graphicFrameLocks', {'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'noChangeAspect': '1'}))
  graph_tag = inline_tag.add_tag(XmlTag('a:graphic', {'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}))
  gdata_tag = graph_tag.add_tag(XmlTag('a:graphicData', {'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'}))
  pic_tag = gdata_tag.add_tag(XmlTag('pic:pic', {'xmlns:pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'}))

  nvpic_tag = pic_tag.add_tag(XmlTag('pic:nvPicPr'))
  nvpic_tag.add_tag(XmlTag('pic:cNvPr', {'id': num, 'name': 'Imagen 19', 'descr': 'Icono&#xA;&#xA;Descripción generada automáticamente'}))
  nvpic_tag.add_tag(XmlTag('pic:cNvPicPr'))

  blip_tag = pic_tag.add_tag(XmlTag('pic:blipFill'))
  blip_tag.add_tag(XmlTag('a:blip', {'r:embed': rid, 'cstate': 'print'}))
  stretch_tag = blip_tag.add_tag(XmlTag('a:stretch'))
  stretch_tag.add_tag(XmlTag('a:fillRect'))

  sppr_tag = pic_tag.add_tag(XmlTag('pic:spPr'))
  xfrm_tag = sppr_tag.add_tag(XmlTag('a:xfrm'))
  xfrm_tag.add_tag(XmlTag('a:off', {'x': '0', 'y': '0'}))
  xfrm_tag.add_tag(XmlTag('a:ext', {'cx': dpi_wd, 'cy': dpi_hg}))
  prst_tag = sppr_tag.add_tag(XmlTag('a:prstGeom', {'prst': 'rect'}))
  prst_tag.add_tag(XmlTag('a:avLst'))
  return run_tag

def get_css_properties(istyle, props):
  """
  Extracts CSS properties relevant for Word.

  Args:
    istyle: CSS style string.
    props: Property dictionary to fill.

  Returns:
    Updated property dictionary.
  """

  if not props:
    props = {}
  stylemap = html.parse_css(istyle)
  #-- underline / strike
  text_decoration = stylemap.get('text-decoration')
  if text_decoration:
    props['underline'] = text_decoration == 'underline'
    props['strike'] = text_decoration == 'line-through'
    #-- align
  text_align = stylemap.get('text-align')
  if text_align:
    if text_align == 'justify':
      text_align = 'both'
    props['align'] = text_align
    #-- color
  color = stylemap.get('color')
  if color:
    props['color'] = color
    #-- fondo
  bgcolor = stylemap.get('background-color')
  if bgcolor:
    props['bgcolor'] = bgcolor
    #-- italic
  italic = stylemap.get('font-style')
  if italic:
    props['italic'] = italic == 'italic'
    #-- bold
  bold = stylemap.get('font-weight')
  if bold:
    props['bold'] = bold in ['bold', '400', '600']
    #-- height
  pxhg = stylemap.get('height')
  if pxhg and pxhg.endswith("px"):
    try:
      pxhg = int(text_util.trim(pxhg[0:len(pxhg) - 2]))
      props['height'] = pxhg
    except ValueError:
      pass
      #-- width
  pxwd = stylemap.get('width')
  if pxwd and pxwd.endswith("px"):
    try:
      pxwd = int(text_util.trim(pxwd[0:len(pxwd) - 2]))
      props['width'] = pxwd
    except ValueError:
      pass
  return props

def get_html_table_properties_to_json(html_tag: XmlTag) -> dict:
  """
  Converts an HTML table into a properties dictionary.

  Args:
    html_tag: <table> tag.

  Returns:
    Dictionary with properties and rows/cells.
  """
  table_props = get_html_table_item_properties(html_tag, False)
  for tag in html_tag.elements:
    if isinstance(tag, XmlTag):
      process_html_table_properties(tag, table_props, [])
  return table_props

def get_html_table_item_properties(html_tag: XmlTag, inner: bool) -> dict:
  """
  Gets properties of an HTML table element.

  Args:
    html_tag: HTML tag (table/tr/td/th).
    inner: If True, includes cell and content properties.

  Returns:
    Property dictionary.
  """
  table_props = get_css_properties(html_tag.get_attr('style'))
  bgcolor = html_tag.get_attr('bgcolor')
  if bgcolor:
    table_props['bgcolor'] = bgcolor
  width = html_tag.get_attr('width')
  if width and not table_props.get('width'):
    table_props['width'] = width
  if inner:
    rowspan = html_tag.get_attr_int('rowspan')
    if rowspan:
      table_props['rowspan'] = rowspan
    colspan = html_tag.get_attr_int('colspan')
    if colspan:
      table_props['colspan'] = colspan
    text = html_tag.get_inner_html()
    if text:
      table_props['#text'] = text
  return table_props

def process_html_table_properties(html_tag: XmlTag, table_props: dict, row: list):
  """
  Walks an HTML tree to build table properties.

  Args:
    html_tag: Current tag.
    table_props: Accumulated dictionary.
    row: Current row (list of cells).
  """
  tag_name = html_tag.name.lower()
  if tag_name == 'caption':
    table_props['caption'] = html_tag.get_text()
    return
  if tag_name == 'tr':
    rows = table_props.get('rows')
    if not rows:
      rows = []
      table_props['rows'] = rows    
    row = []
    rows.append(row)
  if tag_name == 'th' or tag_name == 'td':
    cell = get_html_table_item_properties(html_tag, True)
    cell['cell'] = tag_name
    row.append(cell)
    return
  for tag in html_tag.elements:
    if isinstance(tag, XmlTag):
      process_html_table_properties(tag, table_props, row)


def get_px_size(value, px_size: float = 0) -> int|None:
  """
  Converts a CSS value to approximate twips.

  Args:
    value: CSS value (px or %).
    px_size: Reference size for percentages.

  Returns:
    Size in twips or None.
  """
  if value:
    value = value.strip()
    try:
      if value.endswith("px"):
        return int(float[0:-2] * SIZE_TWIPS_PER_PX)
      if value.endswith("%"):
        return int(float(value[0:-1]) * px_size / 100)
    except ValueError:
      pass
  return None


def get_px_width(value, max_size: float = 0) -> int|None:
  """
  Converts a CSS width to twips with a limit.

  Args:
    value: CSS value.
    max_size: Maximum size in twips.

  Returns:
    Width in twips or None.
  """
  return get_px_size(value, SIZE_WIDTH_CM * SIZE_TWIPS_PER_CM if max_size is None or max_size <= 0 else max_size)


def create_table(num_table, table_props, styles) -> list:
  """
  Creates a Word table from properties.

  Args:
    num_table: Table number (for caption).
    table_props: Table and cell properties.
    styles: Document styles.

  Returns:
    List of XmlTag representing the table.
  """
  out = []
  rows = table_props.get("rows")
  if not rows or len(rows) == 0:
    return out

  tbl = XmlTag(WT.TAG_TABLE)
  tbl_pr = XmlTag("w:tblPr")
  jc = XmlTag("w:jc")
  jc.set_attr(WT.ATTR_VAL, "center")
  tbl_pr.add_tag(jc)
  tbl_borders = XmlTag("w:tblBorders")
  for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
    border = XmlTag(f"w:{side}")
    border.set_attr(WT.ATTR_VAL, "single")
    border.set_attr("w:sz", "1")
    border.set_attr("w:space", "0")
    border.set_attr("w:color", "808080")
    tbl_borders.add_tag(border)
  tbl_pr.add_tag(tbl_borders)
  table_wd = get_px_width(table_props.get("width"))
  if table_wd:                
    tbl_w = XmlTag("w:tblW")
    tbl_w.set_attr(WT.ATTR_WIDTH, table_wd)
    tbl_w.set_attr(WT.ATTR_TYPE, WT.VAL_TYPE_DXA)
    tbl_pr.add_tag(tbl_w)
  tbl.add_tag(tbl_pr)

  first_row = rows[0]
  tbl_grid = XmlTag("w:tblGrid")
  for cell in first_row:
    grid_col = XmlTag("w:gridCol")
    cell_wd = get_px_width(cell.get("width"), table_wd)
    if cell_wd:
      grid_col.set_attr(WT.ATTR_WIDTH, cell_wd)
      grid_col.set_attr(WT.ATTR_TYPE, WT.VAL_TYPE_DXA)
    tbl_grid.add_tag(grid_col)
  tbl.add_tag(tbl_grid)

  numrow = 0
  for row in rows:
    numrow += 1
    tr = XmlTag(WT.TAG_TABLE_ROW)
    for cell in row:
      tc = XmlTag(WT.TAG_TABLE_CELL)
      tc_pr = XmlTag("w:tcPr")
      cell_type = cell.get("cell")
      bgcolor = get_color(cell.get('bgcolor'))
      if not bgcolor:
        if cell_type == 'th':
          bgcolor = styles.style_map.get(Styles.CFG_STYLE_TABLE_CELL_BGCOLOR)
        else:
          bgcolor = styles.style_map.get(Styles.CFG_STYLE_TABLE_CELL_BGCOLOR) if (numrow % 2) == 0 else styles.style_map.get(Styles.CFG_STYLE_TABLE_CELL_BGCOLOR2)
      if bgcolor:
        shd = XmlTag("w:shd")
        shd.set_attr(WT.ATTR_VAL, "clear")
        shd.set_attr("w:color", "auto")
        shd.set_attr("w:fill", bgcolor)
        tc_pr.add_tag(shd)
      cell_wd = get_px_width(cell.get("width"), table_wd)
      if cell_wd:
        tc_w = XmlTag("w:tcW")
        tc_w.set_attr("w:w", cell_wd)
        tc_w.set_attr("w:type", "dxa")
        tc_pr.add_tag(tc_w)
      cs = cell.get("colspan", 1)
      if cs > 1:
        grid_span = XmlTag("w:gridSpan")
        grid_span.set_attr(WT.ATTR_VAL, str(cs))
        tc_pr.add_tag(grid_span)
      rs = cell.get("rowspan", 1)
      if rs > 1:
        vmerge = XmlTag("w:vMerge")
        vmerge.set_attr(WT.ATTR_VAL, "restart")
        tc_pr.add_tag(vmerge)
      tc.add_tag(tc_pr)
      p = XmlTag(WT.TAG_P)
      p_pr = XmlTag(WT.TAG_PPR)
      align = cell.get('align')
      if align:
        jc = XmlTag("w:jc")
        jc.set_attr(WT.ATTR_VAL, align)
        p_pr.add_tag(jc)
      p_style = XmlTag(WT.TAG_P_STYLE)
      if cell_type == "th":
        p_style.add_attr(WT.ATTR_VAL, styles.style_map.get(Styles.CFG_STYLE_TABLE_HEADER))
      else:
        p_style.add_attr(WT.ATTR_VAL, styles.style_map.get(Styles.CFG_STYLE_TABLE_CELL))
      p_pr.add_tag(p_style)
      p.add_tag(p_pr)
      r = XmlTag("w:r")
      t = XmlTag("w:t")
      text = cell.get('#text')
      if text:
        t.set_text(text)
      r.add_tag(t)
      p.add_tag(r)
      tc.add_tag(p)
      tr.add_tag(tc)
    tbl.add_tag(tr)
  out.append(tbl)

  caption = table_props.get('caption')
  if caption:
    p = XmlTag(WT.TAG_P)
    ppr = XmlTag(WT.TAG_PPR)
    pstyle = XmlTag(WT.TAG_P_STYLE)
    pstyle.add_attr(WT.ATTR_VAL, styles.style_map.get(Styles.CFG_STYLE_TABLE_CAPTION))
    ppr.add_tag(pstyle)
    p.add_tag(ppr)
    run = XmlTag(WT.TAG_R)
    p.add_tag(run)
    run.set_tag_text(WT.TAG_T, "Tabla " + str(num_table) + ". " + XmlParser.escape_entities(caption), False)
    out.append(p)

  return out

def get_color(css_color: str|None) -> str|None:
  """
  Converts a CSS color to hex without '#'.

  Args:
    css_color: CSS color.

  Returns:
    Hex without '#' or None.
  """
  if css_color:
    css_color = html.get_rgb_color(css_color)
    if css_color.startswith('#'):
      css_color = css_color[1:]
  return css_color
