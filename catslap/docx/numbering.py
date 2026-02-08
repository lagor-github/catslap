# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)
from catslap.utils.xml import XmlParser
from catslap.utils import types as types_util
from catslap.utils import text as text_util

class Numbering(XmlParser):
  """
  Manages list numbering in Word.
  """
  def __init__(self, pathfile):
    super().__init__()
    self.parse_file(pathfile, 'w:numbering')
    new_num_id = 1
    num_tags = self.root_tag.get_tags('w:num')
    for num_tag in num_tags:
      num_id = types_util.to_int(num_tag.get_attr('w:numId'))
      if num_id > new_num_id:
        new_num_id = num_id + 1
    self.new_num_id = new_num_id

  def add_numbering_start(self, stylename: str) -> int|None:
    """
    Creates a new numbering sequence from a style.

    Args:
      stylename: List style name.

    Returns:
      New numId or None if the style is not found.
    """
    abstract_id = self.find_abstract_id_by_style(stylename)
    if abstract_id is not None:
      num_tag = self.root_tag.add_tag('w:num')
      new_num_id = self.new_num_id
      self.new_num_id += 1
      num_tag.add_attr('w:numId', str(new_num_id))
      abs_tag = num_tag.add_tag('w:abstractNumId')
      abs_tag.add_attr('w:val', str(abstract_id))
      lvl_tag = num_tag.add_tag('w:lvlOverride')
      lvl_tag.add_attr('w:ilvl', '0')
      start_tag = lvl_tag.add_tag('w:startOverride')
      start_tag.add_attr('w:val', '1')
      return new_num_id
    return None

  def find_abstract_id_by_style(self, stylename: str):
    """
    Finds the abstractNumId associated with a style.

    Args:
      stylename: Style name.

    Returns:
      abstractNumId or None if it does not exist.
    """
    abstract_tags = self.root_tag.get_tags('w:abstractNum')
    stylename = text_util.trim(stylename).lower()
    for abstract_tag in abstract_tags:
      lvl_tag = abstract_tag.get_tag('w:lvl', False)
      if not lvl_tag:
        continue
        # -- Remove extra numbering styles not provided by styles.xml
      ppr = lvl_tag.get_tag('w:pPr', False)
      if ppr:
        lvl_tag.remove_tag('w:pPr')
      pstyle_tag = lvl_tag.get_tag('w:pStyle', False)
      if not pstyle_tag:
        continue
      style_id = pstyle_tag.get_attr('w:val')
      if not style_id:
        continue
      style_id = text_util.trim(style_id).lower()
      style_id_low = style_id.lower()
      if style_id_low == stylename:
        return types_util.to_int(abstract_tag.get_attr('w:abstractNumId'))
    return None
