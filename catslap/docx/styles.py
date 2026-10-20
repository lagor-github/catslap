# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from catslap.utils.xml import XmlParser
from catslap.utils import text as text_util


class Styles(XmlParser):

  CFG_STYLE_QUOTE = 'quote' 
  CFG_STYLE_PARAGRAPH = 'paragraph' 
  CFG_STYLE_TABLE_CAPTION = 'table_caption' 
  CFG_STYLE_TABLE_CELL = 'table_cell' 
  CFG_STYLE_TABLE_CELL_BGCOLOR = 'table_cell_bgcolor' 
  CFG_STYLE_TABLE_CELL_BGCOLOR2 = 'table_cell_bgcolor2' 
  CFG_STYLE_TABLE_CELL_COLOR = 'table_cell_color' 
  CFG_STYLE_TABLE_CELL_COLOR2 = 'table_cell_color2' 
  CFG_STYLE_TABLE_HEADER = 'table_header' 
  CFG_STYLE_TABLE_HEADER_BGCOLOR = 'table_header_bgcolor' 
  CFG_STYLE_TABLE_HEADER_COLOR = 'table_header_color' 
  CFG_STYLE_HEADING = 'heading'
  CFG_STYLE_LIST_BULLET = 'list_bullet'
  CFG_STYLE_LIST_NUMBER = 'list_number'
  CFG_STYLE_LINK_TITLE = 'link_title'
  CFG_STYLE_LINK_URL = 'link_url'
  CFG_STYLE_TOKEN = 'token'
  CFG_STYLE_CODEBLOCK = 'codeblock'
  CFG_STYLE_CODE = 'code'
  CFG_STYLE_IMAGE_TITLE = 'image_title'
  CFG_IMAGE_MARGIN = 'image_margin'
  CFG_STYLE_SEC_LEVEL = 'security_level'

  DEFAULT_STYLES = {
    CFG_STYLE_QUOTE: 'Cita',
    CFG_STYLE_PARAGRAPH: 'Normal',
    CFG_STYLE_TABLE_CAPTION: ' Normal',
    CFG_STYLE_TABLE_CELL: 'Normal',
    CFG_STYLE_TABLE_HEADER: 'Normal',
    CFG_STYLE_TABLE_CELL_BGCOLOR: '#FFFFFF',
    CFG_STYLE_TABLE_CELL_BGCOLOR2: "#E9E9E9",
    CFG_STYLE_TABLE_CELL_COLOR: '#000080',
    CFG_STYLE_TABLE_CELL_COLOR2: "#000080",
    CFG_STYLE_TABLE_HEADER_BGCOLOR: '#000080',
    CFG_STYLE_TABLE_HEADER_COLOR: '#FFFFFF',
    CFG_STYLE_HEADING: 'Ttulo',
    CFG_STYLE_LIST_BULLET: 'Listaconvietas, Listaconvieta, List-Bullet, OOXML-BulletList',
    CFG_STYLE_LIST_NUMBER: 'Listaconnmeros, Listaconnmero, List-Number, OOML-NumberList',
    CFG_STYLE_LINK_TITLE: 'Normal',
    CFG_STYLE_LINK_URL: 'Hipervnculo',
    CFG_STYLE_TOKEN: 'Normal',
    CFG_STYLE_CODEBLOCK: 'Normal',
    CFG_STYLE_CODE: 'Normal',
    CFG_STYLE_IMAGE_TITLE: 'Normal',
    CFG_STYLE_SEC_LEVEL: 'Normal',
  }

  CFG_LIST_STYLES = [
    CFG_STYLE_HEADING,
    CFG_STYLE_LIST_BULLET,
    CFG_STYLE_LIST_NUMBER,
  ]
  CFG_SIMPLE_STYLES = [
    CFG_STYLE_QUOTE,
    CFG_STYLE_PARAGRAPH,
    CFG_STYLE_TABLE_CAPTION,
    CFG_STYLE_TABLE_CELL,
    CFG_STYLE_TABLE_CELL_BGCOLOR,
    CFG_STYLE_TABLE_CELL_BGCOLOR2,
    CFG_STYLE_TABLE_CELL_COLOR,
    CFG_STYLE_TABLE_CELL_COLOR2,
    CFG_STYLE_TABLE_HEADER,
    CFG_STYLE_TABLE_HEADER_BGCOLOR,
    CFG_STYLE_TABLE_HEADER_COLOR,
    CFG_STYLE_TOKEN,
    CFG_STYLE_CODEBLOCK,
    CFG_STYLE_CODE,
    CFG_STYLE_IMAGE_TITLE,
    CFG_STYLE_LINK_TITLE,
    CFG_STYLE_LINK_URL,
    CFG_STYLE_SEC_LEVEL,
  ]

  """
  Word styles manager based on styles.xml.

  Attributes:
    style_map: Map of styles
  """
  def __init__(self, pathfile):
    super().__init__()
    self.parse_file(pathfile, 'w:styles')
    self.style_map = {}
    self.image_margin = None
    self._style_tags_by_id = {}

  def find_styles(self):
    """
    Finds relevant styles based on template parameters.

    Args:
      template_params: Style parameter map.
    """
    style_tags = self.root_tag.get_tags('w:style')
    for style_name in Styles.CFG_LIST_STYLES:
      param_value = Styles.DEFAULT_STYLES.get(style_name)
      Styles.find_positional_style(style_tags, param_value, self.style_map, style_name)
    for style_name in Styles.CFG_SIMPLE_STYLES:
      param_value = Styles.DEFAULT_STYLES.get(style_name)
      Styles.find_style(style_tags, param_value, self.style_map, style_name)
    self._style_tags_by_id = {}
    for style_tag in style_tags:
      style_id = style_tag.get_attr('w:styleId')
      if style_id:
        self._style_tags_by_id[style_id] = style_tag

  def get_style_run_properties(self, style_id: str | None):
    if not style_id:
      return None
    style_tag = self._style_tags_by_id.get(style_id)
    if style_tag is None:
      return None
    return self.__resolve_style_rpr(style_tag)

  def get_style_numbering_properties(self, style_id: str | None) -> tuple[int | None, int | None]:
    if not style_id:
      return None, None
    style_tag = self._style_tags_by_id.get(style_id)
    if style_tag is None:
      return None, None
    return self.__resolve_style_numpr(style_tag)

  def __resolve_style_numpr(self, style_tag) -> tuple[int | None, int | None]:
    based_on = style_tag.get_tag('w:basedOn', False)
    base_num_id = None
    base_ilvl = None
    if based_on:
      base_style_id = based_on.get_attr('w:val')
      base_style_tag = self._style_tags_by_id.get(base_style_id) if base_style_id else None
      if base_style_tag is not None:
        base_num_id, base_ilvl = self.__resolve_style_numpr(base_style_tag)
    ppr_tag = style_tag.get_tag('w:pPr', False)
    numpr_tag = ppr_tag.get_tag('w:numPr', False) if ppr_tag else None
    if numpr_tag is None:
      return base_num_id, base_ilvl
    num_id_tag = numpr_tag.get_tag('w:numId', False)
    ilvl_tag = numpr_tag.get_tag('w:ilvl', False)
    num_id = num_id_tag.get_attr_int('w:val') if num_id_tag else None
    ilvl = ilvl_tag.get_attr_int('w:val') if ilvl_tag else None
    return num_id if num_id is not None else base_num_id, ilvl if ilvl is not None else base_ilvl

  def get_positional_style_left_indent(self, style_name: str, level: int) -> int | None:
    style_id = self.get_positional_style_id(style_name, level)
    if not style_id:
      return None
    style_tag = self._style_tags_by_id.get(style_id)
    if style_tag is None:
      return None
    left, _ = Styles.__get_style_indent(style_tag)
    return left

  def get_positional_style_id(self, style_name: str, level: int) -> str | None:
    style_list = self.style_map.get(style_name)
    if not isinstance(style_list, list) or len(style_list) == 0:
      return None
    if level < 0:
      level = 0
    elif level >= len(style_list):
      level = len(style_list) - 1
    return style_list[level] or None

  def __resolve_style_rpr(self, style_tag):
    based_on = style_tag.get_tag('w:basedOn', False)
    base_rpr = None
    if based_on:
      base_style_id = based_on.get_attr('w:val')
      base_style_tag = self._style_tags_by_id.get(base_style_id) if base_style_id else None
      if base_style_tag is not None:
        base_rpr = self.__resolve_style_rpr(base_style_tag)
    current_rpr = style_tag.get_tag('w:rPr', False)
    if base_rpr is None and current_rpr is None:
      return None
    if base_rpr is None:
      return current_rpr.clone(True)
    merged = base_rpr.clone(True)
    if current_rpr:
      for child in current_rpr.elements:
        if not hasattr(child, 'name'):
          merged.add_element(child.clone(True))
          continue
        merged.remove_tag(child.name)
        merged.add_tag(child.clone(True))
    return merged

  @staticmethod
  def __only_ascii(value) -> str:
    out = ''
    for ch in value:
      if (ch >= 'a' and ch <= 'z') or (ch >= 'A' and ch <= 'Z') or (ch >= '0' and ch <= '9') or ch in ['-', '_', ',']:
        out += ch
    return out

  @staticmethod
  def __get_style_list_autonum(stylevalue) -> list:
    curlist = text_util.split_no_empty(stylevalue, ',')
    idx = len(curlist)
    if idx <= 0:
      return []
    (value, num) = Styles._get_name_and_num(curlist[idx - 1])
    if not num:
      num = idx
    while idx < 6:
      curlist.append(value + str(num + 1))
      num += 1
      idx += 1
    return curlist

  @staticmethod
  def _get_name_and_num(name: str) -> tuple[str, int]:
    match = re.match(r"(.+?)(\d+)$", name)
    num = None
    if match and match.group(1) in Styles.CFG_LIST_STYLES:
      name = match.group(1)
      num = int(match.group(2))
    return name, num


  def setStyle(self, name, value) -> str | list[str] | tuple[int, int, int, int]:
    """
    Updates a specific style and returns its normalized value.

    Args:
      name: Style parameter name.
      value: Style value.
    """
    if name.startswith('style_'):
      name = name[6:]
    if name == Styles.CFG_IMAGE_MARGIN:
      self.image_margin = Styles.__parse_image_margin(value)
      return self.image_margin
    value = Styles.__only_ascii(value)

    (name, num) = Styles._get_name_and_num(name)
    if num is None and name in Styles.CFG_LIST_STYLES:
      value = Styles.__get_style_list_autonum(value)
      self.style_map[name] = value
      return value
    if num is not None:
      if num <= 0:
        num = 1
      elif num > 6:
        num = 6
      style_list = self.style_map.get(name)
      if style_list is not None and (num - 1) < len(style_list):
        style_list[num - 1] = value
    else:
      self.style_map[name] = value
    return value

  @staticmethod
  def __parse_image_margin(value: str) -> tuple[int, int, int, int]:
    """Convert top/right/bottom/left margins from centimeters to Word EMUs."""
    error = 'image_margin requires four non-negative finite numbers in cm (top right bottom left)'
    parts = value.split()
    if len(parts) != 4:
      raise ValueError(error)
    margins = []
    for part in parts:
      try:
        cm = Decimal(part)
      except InvalidOperation as exc:
        raise ValueError(error) from exc
      if not cm.is_finite() or cm < 0:
        raise ValueError(error)
      if cm > Decimal(4294967295) / Decimal(360000):
        raise ValueError('image_margin exceeds the Word distance limit')
      margins.append(int((cm * 360000).to_integral_value(rounding=ROUND_HALF_UP)))
    return tuple(margins)


  @staticmethod
  def find_positional_style(style_tags: list, stylenames: str, style_map, style_name):
    """
    Assigns positional styles (1..6) to a list.

    Args:
      style_tags: List of style tags.
      stylenames: Comma-separated candidate names.
      style_map: Style map where set style_name
      style_name: Style name
    """
    stylename_list = stylenames.split(',')
    listbox = style_map.get(style_name)
    if listbox is None:
      listbox = ['', '', '', '', '', '', '']
      style_map[style_name] = listbox
    for stylename in stylename_list:
      Styles.__set_style_id(stylename, style_tags, listbox)
  @staticmethod
  def __get_style_indent(style_tag) -> tuple[int, int]:
    ppr_tag = style_tag.get_tag('w:pPr', False)
    ind_tag = ppr_tag.get_tag('w:ind', False) if ppr_tag else None
    left = ind_tag.get_attr_int('w:left') if ind_tag else None
    hanging = ind_tag.get_attr_int('w:hanging') if ind_tag else None
    return left if left is not None else 720, hanging if hanging is not None else 360

  @staticmethod
  def __set_style_id(stylename: str, style_tags: list, listbox: list):
    for style_tag in style_tags:
      style_id = style_tag.get_attr('w:styleId')
      stylename = text_util.trim(stylename).lower()
      style_id_low = style_id.lower()
      if not style_id_low.startswith(stylename):
        continue
      num = 1
      if style_id_low != stylename:
        snum = style_id_low[len(stylename):]
        if snum[0] == ' ':
          snum = snum[1:]
        try:
          num = int(snum)
        except ValueError:
          continue
      if num < 1 or num > 6 or (num - 1) >= len(listbox) or (listbox[num-1] is not None and listbox[num-1] != ''):
        continue
      listbox[num-1] = style_id

  @staticmethod
  def find_style(style_tags: list, stylenames: str, style_map, style_name):
    """
    Finds a style by name in the style list.

    Args:
      style_tags: List of style tags.
      stylenames: Comma-separated candidate names.
      style_map: Style map where set style_name
      style_name: Style name

    Returns:
      Style identifier or None.
    """
    stylename_list = stylenames.split(',')
    for stylename in stylename_list:
      for style_tag in style_tags:
        style_id = style_tag.get_attr('w:styleId')
        stylename = text_util.trim(stylename).lower()
        style_id_low = style_id.lower()
        if style_id_low == stylename:
          style_map[style_name] = style_id
          return
