# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)

from catslap.utils.xml import XmlParser, XmlParserException

WORD_CONTENT_TYPES_XMLNS = "http://schemas.openxmlformats.org/package/2006/content-types"

def add_dict(dmap: dict, key: str, value):
  """
  Adds a key to a dict if the value is truthy.

  Args:
    dmap: Destination dictionary.
    key: Key to add.
    value: Value to evaluate.
  """
  if value:
    dmap[key] = value

class Default:
    """
    Default entry for Content Types.
    """
    def __init__(self, attrs: dict):
        """
        Represents a Default entry in [Content_Types].xml.

        Args:
          attrs: Attributes of the Default tag.
        """
        self.extension = attrs.get('Extension')
        self.content_type = attrs.get('ContentType')

class Override:
    """
    Override entry for Content Types.
    """
    def __init__(self, attrs: dict):
        """
        Represents an Override entry in [Content_Types].xml.

        Args:
          attrs: Attributes of the Override tag.
        """
        self.part_name = attrs.get('PartName')
        self.content_type = attrs.get('ContentType')

class ContentTypes(XmlParser):
  """
  Manages the [Content_Types].xml file.

  Attributes:
    defaults: List of Default entries.
    overrides: List of Override entries.
  """
  def __init__(self, pathfile):
    super().__init__()
    self.defaults = []
    self.overrides = []
    block = self.parse_file(pathfile, 'Types')
    blocks = block.get_tags()
    for block in blocks:
      tag = block.name
      attrs = block.attrs
      if tag == 'Default':
        entry = Default(attrs)
        self.defaults.append(entry)
      elif tag == 'Override':
        entry = Override(attrs)
        self.overrides.append(entry)
      else:
        raise XmlParserException("XML tag <" + tag + " ...> not valid for this file")

  def add_default(self, extension: str, content_type: str) -> Default:
    """
    Adds or updates a Default entry.

    Args:
      extension: File extension.
      content_type: Associated Content-Type.

    Returns:
      Created or updated Default entry.
    """
    for entry in self.defaults:
      if entry.extension == extension:
        entry.content_type = content_type
        return entry
    attrs = {}
    add_dict(attrs, 'Extension', extension)
    add_dict(attrs, 'ContentType', content_type)
    self.defaults.append(Default(attrs))
