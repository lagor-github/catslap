# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026


import re
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
SIZE_EMU_PER_TWIP = 635
SIZE_TWIPS_PER_CM = 567
SIZE_WIDTH_CM = 17
SIZE_HEIGHT_CM = 24
SIZE_EMU_PER_PX = 9525          # 914400 EMU/inch ÷ 96 DPI
SIZE_MAX_WIDTH_EMU = int(SIZE_WIDTH_CM * 360000)  # 17 cm en EMU
SIZE_DEFAULT_TABLE_WIDTH_TWIPS = SIZE_WIDTH_CM * SIZE_TWIPS_PER_CM


class _TableRow(list):
  def __init__(self, is_header: bool = False):
    super().__init__()
    self.is_header = is_header


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
    max_width_twips = runprops.get('max_width_twips')
    return create_image(mediatype, data, wd, hg, relationships, types, max_width_twips)
  if image and image.startswith('blob:'):
    try:
      mediatype, data = html.fetch_blob_image(image)
    except html.HtmlException as e:
      raise XmlParserException(str(e))
    wd = runprops.get('width')
    hg = runprops.get('height')
    max_width_twips = runprops.get('max_width_twips')
    return create_image(mediatype, data, wd, hg, relationships, types, max_width_twips)

  bold = runprops.get('bold') is True
  italic = runprops.get('italic') is True
  strike = runprops.get('strike') is True
  underline = runprops.get('underline') is True
  color = runprops.get('color')
  bgcolor = runprops.get('bgcolor')
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
    bgcolor = bgcolor if bgcolor is not None else rpr_tag.get_tag_attr(WT.TAG_SHADOW, WT.ATTR_FILL, False)
  else:
    out_rpr_tag = XmlTag(WT.TAG_RPR)

  out_r_tag = XmlTag(WT.TAG_R)
  out_r_tag.add_tag(out_rpr_tag)
  create_rpr_style(out_rpr_tag, bold, italic, underline, strike, style, color, bgcolor)

  if text is None:
    text = ''
  out_t_tag = out_r_tag.add_tag(XmlTag('w:t', {'xml:space': 'preserve'}))
  #-- escapa sólo si no tiene caracteres de escape (ya está escapado)
  if text.find('&lt;') < 0 and text.find('&gt;') < 0 and text.find('&amp;') < 0:
    text = XmlParser.escape_entities(text)
  out_t_tag.add_text(text)
  if link:
    relationship = relationships.add_relationship_hyperlink(link)
    hyper_tag = XmlTag(WT.TAG_HYPERLINK, {WT.ATTR_ID: relationship.rid, 'w:history': '1'})
    hyper_tag.add_tag(out_r_tag)
    return hyper_tag
  return out_r_tag

def create_rpr_style(out_rpr_tag: XmlTag, bold: bool, italic: bool, underline: bool, strike: bool, style: str|None, color: str|None, bgcolor: str|None):
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
    out_rpr_tag.add_tag(XmlTag(WT.TAG_COLOR, {WT.ATTR_VAL: get_color(color)}))
  if bgcolor:
    out_rpr_tag.add_tag(XmlTag(WT.TAG_SHADOW, {WT.ATTR_VAL: WT.ATTR_VAL_CLEAR, WT.ATTR_COLOR: WT.ATTR_VAL_AUTO, WT.ATTR_FILL: get_color(bgcolor)}))

def _parse_px(value) -> float | None:
  """Converts a CSS pixel value ('300px', '300', or a number) to float pixels."""
  if value is None:
    return None
  if isinstance(value, (int, float)):
    return float(value)
  v = str(value).strip()
  if v.endswith('px'):
    try:
      return float(v[:-2])
    except ValueError:
      return None
  try:
    return float(v)
  except ValueError:
    return None


def create_image(mediatype: str, data: bytes, pxwd: int|None, pxhg: int|None, relationships: Relationships, types: ContentTypes, max_width_twips: int | None = None) -> XmlTag:
  """
  Creates a run with an embedded image.

  Args:
    mediatype: Image MIME type.
    data: Image bytes.
    pxwd: Width in pixels (optional, accepts int or CSS string like '300px').
    pxhg: Height in pixels (optional, accepts int or CSS string like '200px').
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
  px_wd = _parse_px(pxwd)
  px_hg = _parse_px(pxhg)
  if not px_wd or not px_hg:
    stream = BytesIO(data)
    img = Image.open(stream).convert("RGBA")
    stream.close()
    if px_wd and not px_hg:
      px_hg = px_wd * img.height / img.width
    elif px_hg and not px_wd:
      px_wd = px_hg * img.width / img.height
    if not px_wd:
      px_wd = img.width
    if not px_hg:
      px_hg = img.height
  dpi_wd = int(round(px_wd * SIZE_EMU_PER_PX))
  dpi_hg = int(round(px_hg * SIZE_EMU_PER_PX))
  max_width_emu = SIZE_MAX_WIDTH_EMU
  if isinstance(max_width_twips, int) and max_width_twips > 0:
    max_width_emu = max(1, int(round(max_width_twips * SIZE_EMU_PER_TWIP)))
  if dpi_wd > max_width_emu:
    dpi_hg = int(round(dpi_hg * max_width_emu / dpi_wd))
    dpi_wd = max_width_emu

  run_tag = XmlTag('w:r')
  rpr_tag = run_tag.add_tag('w:rPr')
  rpr_tag.add_tag(XmlTag('w:rFonts', {'w:cstheme': 'minorHAnsi'}))
  rpr_tag.add_tag(XmlTag('w:noProof'))
  drawing_tag = run_tag.add_tag('w:drawing')
  anchor_tag = drawing_tag.add_tag(XmlTag('wp:anchor', {
    'distT': '0',
    'distB': '0',
    'distL': '114300',
    'distR': '114300',
    'simplePos': '0',
    'relativeHeight': '251658240',
    'behindDoc': '0',
    'locked': '0',
    'layoutInCell': '1',
    'allowOverlap': '0',
    'wp14:anchorId': '27F8AE68',
    'wp14:editId': '2DB58A57',
  }))
  anchor_tag.add_tag(XmlTag('wp:simplePos', {'x': '0', 'y': '0'}))
  anchor_tag.add_tag(XmlTag('wp:positionH', {'relativeFrom': 'column'})).add_tag(XmlTag('wp:align')).add_text('left')
  anchor_tag.add_tag(XmlTag('wp:positionV', {'relativeFrom': 'paragraph'})).add_tag(XmlTag('wp:posOffset')).add_text('0')
  anchor_tag.add_tag(XmlTag('wp:extent', {'cx': dpi_wd, 'cy': dpi_hg}))
  anchor_tag.add_tag(XmlTag('wp:effectExtent', {'l': '0', 't': '0', 'r': '0', 'b': '0'}))
  anchor_tag.add_tag(XmlTag('wp:wrapSquare', {'wrapText': 'bothSides'}))
  anchor_tag.add_tag(XmlTag('wp:docPr', {'id': num+1, 'name': 'Imagen 19', 'descr': 'Icono&#xA;&#xA;Descripción generada automáticamente'}))
  cnv_gr_tag = anchor_tag.add_tag(XmlTag('wp:cNvGraphicFramePr'))
  cnv_gr_tag.add_tag(XmlTag('a:graphicFrameLocks', {'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'noChangeAspect': '1'}))
  graph_tag = anchor_tag.add_tag(XmlTag('a:graphic', {'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}))
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

def get_css_properties(istyle, props = None):
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
    props['bold'] = bold in ['bold', '600', '700', '800', '900']
  #-- height
  pxhg = stylemap.get('height')
  if pxhg:
    props['height'] = pxhg
  #-- width
  pxwd = stylemap.get('width')
  if pxwd:
    props['width'] = pxwd
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
      process_html_table_properties(tag, table_props, [], False)
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
  if width:
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

def process_html_table_properties(html_tag: XmlTag, table_props: dict, row: list, in_header: bool = False):
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
  if tag_name == 'thead':
    in_header = True
  if tag_name == 'tr':
    rows = table_props.get('rows')
    if not rows:
      rows = []
      table_props['rows'] = rows    
    row = _TableRow(in_header)
    rows.append(row)
  if tag_name == 'th' or tag_name == 'td':
    cell = get_html_table_item_properties(html_tag, True)
    cell['cell'] = tag_name
    row.append(cell)
    return
  for tag in html_tag.elements:
    if isinstance(tag, XmlTag):
      process_html_table_properties(tag, table_props, row, in_header)


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
        return int(float(value[0:-2]) * SIZE_TWIPS_PER_PX)
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


def _fit_widths_to_limit(widths: list[int | None], max_width: int | None) -> list[int | None]:
  if not max_width or max_width <= 0:
    return list(widths)
  numeric = [width for width in widths if isinstance(width, int) and width > 0]
  total = sum(numeric)
  if total <= 0 or total <= max_width:
    return list(widths)
  scale = max_width / total
  fitted: list[int | None] = []
  for width in widths:
    if not isinstance(width, int) or width <= 0:
      fitted.append(width)
      continue
    fitted.append(max(1, int(round(width * scale))))
  overflow = sum(width for width in fitted if isinstance(width, int) and width > 0) - max_width
  idx = len(fitted) - 1
  while overflow > 0 and idx >= 0:
    width = fitted[idx]
    if isinstance(width, int) and width > 1:
      fitted[idx] = width - 1
      overflow -= 1
    else:
      idx -= 1
  return fitted


def _extract_plain_cell_text(value: str | None) -> str:
  text = str(value or '')
  text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
  text = re.sub(r'</p\s*>', '\n', text, flags=re.IGNORECASE)
  text = re.sub(r'<[^>]+>', ' ', text)
  text = XmlParser.resolve_entities(text)
  text = re.sub(r'[ \t\r\f\v]+', ' ', text)
  text = re.sub(r' *\n *', '\n', text)
  return text.strip()


def _estimate_cell_weight(cell: dict) -> float:
  text = _extract_plain_cell_text(cell.get('#text'))
  if not text:
    return 2.0
  lines = [line.strip() for line in text.split('\n') if line.strip()]
  words = re.findall(r'\S+', text)
  longest_word = max((len(word) for word in words), default=1)
  max_line = max((len(line) for line in lines), default=len(text))
  total_len = len(text)
  weight = max(
    longest_word * 2.2,
    max_line * 1.25,
    total_len * 0.35,
    2.0,
  )
  if cell.get('cell') == 'th':
    weight = max(weight * 1.2, longest_word * 2.8, max_line * 1.5)
  return weight


def _count_table_columns(rows: list) -> int:
  max_cols = 0
  for row in rows:
    total = 0
    for cell in row:
      colspan = cell.get("colspan", 1) or 1
      total += colspan if colspan > 0 else 1
    if total > max_cols:
      max_cols = total
  return max_cols


def _build_column_widths(rows: list, table_wd: int | None, max_width: int) -> list[int | None]:
  num_cols = _count_table_columns(rows)
  if num_cols <= 0:
    return []

  explicit_widths: list[int | None] = [None] * num_cols
  weights: list[float] = [2.0] * num_cols

  for row in rows:
    col_idx = 0
    for cell in row:
      colspan = cell.get("colspan", 1) or 1
      if colspan < 1:
        colspan = 1
      cell_width = get_px_width(cell.get("width"), table_wd if table_wd else max_width)
      if cell_width:
        piece = max(1, int(round(cell_width / colspan)))
        for offset in range(colspan):
          idx = col_idx + offset
          if idx < num_cols:
            explicit_widths[idx] = max(explicit_widths[idx] or 0, piece)
      weight_piece = _estimate_cell_weight(cell) / colspan
      for offset in range(colspan):
        idx = col_idx + offset
        if idx < num_cols:
          weights[idx] = max(weights[idx], weight_piece)
      col_idx += colspan

  explicit_total = sum(width for width in explicit_widths if isinstance(width, int) and width > 0)
  remaining_cols = [idx for idx, width in enumerate(explicit_widths) if not isinstance(width, int) or width <= 0]
  if table_wd and table_wd > 0:
    target_width = min(table_wd, max_width)
  elif explicit_total > 0 and not remaining_cols:
    target_width = min(explicit_total, max_width)
  else:
    target_width = max_width

  fitted = list(explicit_widths)
  if explicit_total > 0:
    # If some columns have no explicit width, reserve at least 1 twip for each
    # before proportionally scaling the explicit ones into the remaining budget.
    min_remaining = len(remaining_cols)
    explicit_budget = max(0, target_width - min_remaining)
    fitted = _fit_widths_to_limit(fitted, explicit_budget)

  if not remaining_cols:
    return fitted

  used_width = sum(width for width in fitted if isinstance(width, int) and width > 0)
  remaining_width = max(0, target_width - used_width)
  if remaining_width <= 0:
    for idx in remaining_cols:
      fitted[idx] = 1
    return _fit_widths_to_limit(fitted, target_width)

  total_weight = sum(weights[idx] for idx in remaining_cols)
  if total_weight <= 0:
    total_weight = len(remaining_cols)
    for idx in remaining_cols:
      weights[idx] = 1.0

  assigned = 0
  for pos, idx in enumerate(remaining_cols):
    if pos == len(remaining_cols) - 1:
      width = remaining_width - assigned
    else:
      width = int(round(remaining_width * weights[idx] / total_weight))
      assigned += width
    fitted[idx] = max(1, width)
  return _fit_widths_to_limit(fitted, target_width)


def _is_header_row(row: list) -> bool:
  if getattr(row, "is_header", False):
    return True
  cells = [cell for cell in row if isinstance(cell, dict)]
  return len(cells) > 0 and all(cell.get("cell") == "th" for cell in cells)


def _configure_table_row_pagination(tr: XmlTag, is_header: bool):
  tr_pr = tr.get_tag("w:trPr", False)
  if tr_pr is None:
    tr_pr = XmlTag("w:trPr")
    tr.elements.insert(0, tr_pr)
    tr_pr.parent = tr
  tr_pr.remove_tag("w:cantSplit")
  if is_header and tr_pr.get_tag("w:tblHeader", False) is None:
    tr_pr.add_tag(XmlTag("w:tblHeader"))


def create_table(num_table, table_props, styles, max_table_width: int | None = None) -> list:
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
  tbl_layout = XmlTag("w:tblLayout")
  tbl_layout.set_attr(WT.ATTR_TYPE, "fixed")
  tbl_pr.add_tag(tbl_layout)
  max_width = max_table_width if isinstance(max_table_width, int) and max_table_width > 0 else SIZE_DEFAULT_TABLE_WIDTH_TWIPS
  table_wd = get_px_width(table_props.get("width"), max_width)
  col_widths = _build_column_widths(rows, table_wd, max_width)
  final_total = sum(width for width in col_widths if isinstance(width, int) and width > 0)
  if table_wd:
    table_wd = min(table_wd, max_width, final_total if final_total > 0 else table_wd)
  elif final_total > 0:
    table_wd = min(final_total, max_width)
  else:
    table_wd = max_width
  if table_wd:
    tbl_w = XmlTag("w:tblW")
    tbl_w.set_attr(WT.ATTR_WIDTH, table_wd)
    tbl_w.set_attr(WT.ATTR_TYPE, WT.VAL_TYPE_DXA)
    tbl_pr.add_tag(tbl_w)
  tbl.add_tag(tbl_pr)

  tbl_grid = XmlTag("w:tblGrid")
  for cell_wd in col_widths:
    grid_col = XmlTag("w:gridCol")
    if cell_wd:
      grid_col.set_attr(WT.ATTR_WIDTH, cell_wd)
      grid_col.set_attr(WT.ATTR_TYPE, WT.VAL_TYPE_DXA)
    tbl_grid.add_tag(grid_col)
  tbl.add_tag(tbl_grid)

  numrow = 0
  header_rows_open = True
  for row in rows:
    numrow += 1
    tr = XmlTag(WT.TAG_TABLE_ROW)
    is_header_row = header_rows_open and _is_header_row(row)
    if not is_header_row:
      header_rows_open = False
    _configure_table_row_pagination(tr, is_header_row)
    col_idx = 0
    for cell in row:
      tc = XmlTag(WT.TAG_TABLE_CELL)
      tc_pr = XmlTag("w:tcPr")
      cell_type = cell.get("cell")

      bgcolor = get_color(cell.get('bgcolor'))
      if not bgcolor:
        if cell_type == 'th':
          bgcolor = styles.style_map.get(Styles.CFG_STYLE_TABLE_HEADER_BGCOLOR)
        else:
          bgcolor = styles.style_map.get(Styles.CFG_STYLE_TABLE_CELL_BGCOLOR) if (numrow % 2) == 0 else styles.style_map.get(Styles.CFG_STYLE_TABLE_CELL_BGCOLOR2)
      if bgcolor:
        shd = XmlTag("w:shd")
        shd.set_attr(WT.ATTR_VAL, "clear")
        shd.set_attr("w:fill", bgcolor)
        shd.set_attr("w:color", "auto")
        tc_pr.add_tag(shd)

      cs = cell.get("colspan", 1)
      if not cs or cs < 1:
        cs = 1
      cell_wd = 0
      found_cell_width = False
      for offset in range(cs):
        width = col_widths[col_idx + offset] if (col_idx + offset) < len(col_widths) else None
        if isinstance(width, int) and width > 0:
          cell_wd += width
          found_cell_width = True
      if not found_cell_width:
        cell_wd = get_px_width(cell.get("width"), table_wd if table_wd else max_width)
      if cell_wd:
        tc_w = XmlTag("w:tcW")
        tc_w.set_attr("w:w", cell_wd)
        tc_w.set_attr("w:type", "dxa")
        tc_pr.add_tag(tc_w)
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
      p_style_id = styles.style_map.get(Styles.CFG_STYLE_TABLE_HEADER) if cell_type == "th" else styles.style_map.get(Styles.CFG_STYLE_TABLE_CELL)
      if align:
        jc = XmlTag("w:jc")
        jc.set_attr(WT.ATTR_VAL, align)
        p_pr.add_tag(jc)
      p_style = XmlTag(WT.TAG_P_STYLE)
      p_style.add_attr(WT.ATTR_VAL, p_style_id)
      p_pr.add_tag(p_style)
      p.add_tag(p_pr)
      r = XmlTag("w:r")
      r_pr = styles.get_style_run_properties(p_style_id)
      if r_pr is None:
        r_pr = XmlTag(WT.TAG_RPR)
      r.add_tag(r_pr)
  
      bold = cell.get('bold')
      italic = cell.get('italic')
      underline = cell.get('underline')
      strike = cell.get('strike')
      color = cell.get('color')
      if not color:
        if cell_type == 'th':
          color = styles.style_map.get(Styles.CFG_STYLE_TABLE_HEADER_COLOR)
        else:
          color = styles.style_map.get(Styles.CFG_STYLE_TABLE_CELL_COLOR) if (numrow % 2) == 0 else styles.style_map.get(Styles.CFG_STYLE_TABLE_CELL_COLOR2)
      create_rpr_style(r_pr, bold, italic, underline, strike, None, color, None)
      t = XmlTag("w:t")
      text = cell.get('#text')
      if text:
        t.set_text(text)
      r.add_tag(t)
      p.add_tag(r)
      tc.add_tag(p)
      tr.add_tag(tc)
      col_idx += cs
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
  color = html.get_rgb_color(css_color)
  return color
