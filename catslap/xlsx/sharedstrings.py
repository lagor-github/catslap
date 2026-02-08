# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)
from catslap.base import utils as common
from catslap.utils import text as text_util
from catslap.utils.xml import XmlParser, XmlTag


EXCEL_SHARED_STRINGS = "xl/sharedStrings.xml"
SHAREDSTRINGS_XMLNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


class SharedStrings(XmlParser):
  """
  Manages the Excel sharedStrings.xml file.

  Attributes:
    strings: List of stored strings.
    root_tag: Root XML tag.
  """
  def __init__(self, pathfile: str|None = None):
    super().__init__()
    self.strings = []
    if pathfile:
      pathfile = common.complete_pathfile(pathfile, EXCEL_SHARED_STRINGS)
      try:
        sst = self.parse_file(pathfile, 'sst')
        si_list = sst.get_tags('si')
        for si in si_list:
          text = si.get_tag_text('t', False);
          if text:
            self.strings.append(XmlParser.resolve_entities(text))
      except FileNotFoundError:
        pass
    self.root_tag = XmlTag('sst')
    self.root_tag.attrs['xmlns'] = SHAREDSTRINGS_XMLNS

  def get_string(self, idx: int) -> str:
    """
    Gets a string by index.

    Args:
      idx: String index.

    Returns:
      String or '' if it does not exist.
    """
    return self.strings[idx] if idx >= 0 and idx < len(self.strings) else ''

  def set_string(self, idx: int, string: str):
    """
    Assigns a string at a position.

    Args:
      idx: String index.
      string: Value to assign.
    """
    if idx >= 0 and idx < len(self.strings):
      self.strings[idx] = string

  def del_string(self, idx: int):
    """
    Deletes a string by index.

    Args:
      idx: Index to remove.

    Returns:
      New list size.
    """
    if idx >= 0 and idx < len(self.strings):
      del self.strings[idx]
    return len(self.strings)

  def index_of(self, value: str) -> int:
    """
    Gets the index of a string.

    Args:
      value: String to search.

    Returns:
      Index or -1 if not found.
    """
    try:
      return self.strings.index(value)
    except ValueError:
      return -1

  def add_string(self, value: str) -> int:
    """
    Adds a string if it does not exist.

    Args:
      value: String to add.

    Returns:
      String index.
    """
    idx = self.index_of(value)
    if idx < 0:
      self.strings.append(value)
      idx = len(self.strings) - 1
    return idx

  def count(self) -> int:
    """
    Returns the number of stored strings.

    Returns:
      Number of strings.
    """
    return len(self.strings)

  def write(self):
    """
    Writes the sharedStrings.xml file.

    Raises:
      OSError: If the file cannot be written.
    """
    sst = self.root_tag
    for string in self.strings:
      si_tag = sst.add_tag('si')
      if text_util.is_empty(string):
        t_tag = si_tag.add_tag_text('t', '\n')
        t_tag.add_attr('xml:space', 'preserve')
      else:
        si_tag.add_tag_text('t', string)
    self.write_file()
