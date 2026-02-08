# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)
from catslap.utils.xml import XmlParser
from catslap.utils import text as text_util


class Styles(XmlParser):

  CFG_STYLE_QUOTE = 'quote' 
  CFG_STYLE_PARAGRAPH = 'paragraph' 
  CFG_STYLE_TABLE_CAPTION = 'table_caption' 
  CFG_STYLE_TABLE_CELL = 'table_cell' 
  CFG_STYLE_TABLE_CELL_BGCOLOR = 'table_cell_bgcolor' 
  CFG_STYLE_TABLE_CELL_BGCOLOR2 = 'table_cell_bgcolor2' 
  CFG_STYLE_TABLE_HEADER = 'table_header' 
  CFG_STYLE_TABLE_HEADER_BGCOLOR = 'table_header_bgcolor' 
  CFG_STYLE_HEADING = 'heading'
  CFG_STYLE_LIST_BULLET = 'list_bullet'
  CFG_STYLE_LIST_NUMBER = 'list_number'
  CFG_STYLE_LINK_TITLE = 'link_title'
  CFG_STYLE_LINK_URL = 'link_url'
  CFG_STYLE_TOKEN = 'token'
  CFG_STYLE_CODEBLOCK = 'codeblock'
  CFG_STYLE_CODE = 'code'
  CFG_STYLE_SEC_LEVEL = 'security_level'

  DEFAULT_STYLES = {
    CFG_STYLE_QUOTE: 'Cita',
    CFG_STYLE_PARAGRAPH: 'Normal',
    CFG_STYLE_TABLE_CAPTION: ' Normal',
    CFG_STYLE_TABLE_CELL: 'Normal',
    CFG_STYLE_TABLE_HEADER: 'Normal',
    CFG_STYLE_TABLE_CELL_BGCOLOR: '#FFFFFF',
    CFG_STYLE_TABLE_CELL_BGCOLOR2: "#E9E9E9",
    CFG_STYLE_TABLE_HEADER_BGCOLOR: '#000080',
    CFG_STYLE_HEADING: 'Ttulo',
    CFG_STYLE_LIST_BULLET: 'Listaconvietas, Listaconvieta',
    CFG_STYLE_LIST_NUMBER: 'Listaconnmeros, Listaconnmero',
    CFG_STYLE_LINK_TITLE: 'Normal',
    CFG_STYLE_LINK_URL: 'Hipervnculo',
    CFG_STYLE_TOKEN: 'Normal',
    CFG_STYLE_CODEBLOCK: 'Normal',
    CFG_STYLE_CODE: 'Normal',
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
    CFG_STYLE_TABLE_HEADER,
    CFG_STYLE_TABLE_HEADER_BGCOLOR,
    CFG_STYLE_TOKEN,
    CFG_STYLE_CODEBLOCK,
    CFG_STYLE_CODE,
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
    while idx < 6:
      value = curlist[idx - 1]
      if value.endswith(str(idx)):
        value = value[0:-1]
      value += str(idx + 1)
      curlist.append(value)
      idx += 1
    return curlist

  def setStyle(self, name, value) -> str:
    """
    Updates a specific style and returns its normalized value.

    Args:
      name: Style parameter name.
      value: Style value.
    """
    if name.startswith('style_'):
      name = name[6:]
    value = Styles.__only_ascii(value)

    if name in Styles.CFG_LIST_STYLES:
      value = Styles.__get_style_list_autonum(value)
      self.style_map[name] = value
      return value
    self.style_map[name] = value
    return value


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
      if num < 1 or num > 6 or (num - 1) >= len(listbox) or listbox[num-1] is not None:
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
