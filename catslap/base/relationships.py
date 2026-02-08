# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)

from catslap.utils import file as file_util
from catslap.utils.xml import XmlParser, XmlParserException, XmlTag

# -- Common
RELATIONSHIPS_XMLNS = "http://schemas.openxmlformats.org/package/2006/relationships"
RELATIONSHIP_BASE_URL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
RELATIONSHIP_TYPE_THEME = 'theme'
RELATIONSHIP_TYPE_STYLES = 'styles'

# -- Word
RELATIONSHIP_TYPE_SETTINGS = 'settings'
RELATIONSHIP_TYPE_FONT_TABLE = 'fontTable'
RELATIONSHIP_TYPE_NUMBERING = 'numbering'
RELATIONSHIP_TYPE_HYPERLINK = 'hyperlink'
RELATIONSHIP_TYPE_IMAGE = 'image'
RELATIONSHIP_TYPE_WEB_SETTING = 'webSettings'

# -- Excel
RELATIONSHIP_TYPE_TABLE = 'table'
RELATIONSHIP_TYPE_WORKSHEET = 'worksheet'
RELATIONSHIP_TYPE_SHAREDSTRINGS = 'sharedStrings'


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


class Relationship:
    """
    Represents an OOXML relationship.
    """
    def __init__(self, attrs: dict):
        """
        Creates a relationship from XML attributes.

        Args:
          attrs: Attributes of the Relationship tag.
        """
        self.rid = attrs.get('Id')
        self.type = attrs.get('Type')
        self.target = attrs.get('Target')
        self.target_mode = attrs.get('TargetMode')


class Relationships(XmlParser):
  """
  Manages OOXML relationships for a package.

  Attributes:
    relations: List of loaded relationships.
    max_id: Last assigned rId.
    path_file: Base path for the .rels file.
    base_path: Base path of the package.
  """
  def __init__(self, basefile: str, filename: str, load: bool = True):
    super().__init__()
    self.relations = []
    self.max_id = 0
    self.path_file = file_util.get_pathname(filename)
    if self.path_file.endswith('_rels/'):
      self.path_file = self.path_file[0:-6]
    self.base_path = basefile
    if not self.base_path.endswith('/'):
      self.base_path = self.base_path + '/'
    if load:
      try:
        rel_tags = self.parse_file(filename, 'Relationships')
        blocks = rel_tags.get_tags('Relationship')
        num = 1
        for block in blocks:
          attrs = block.attrs
          if not attrs:
            continue
          rs = Relationship(attrs)
          target = rs.target
          if target:
            if target.startswith('../') or not target.startswith('/'):
              rs.target = file_util.complete_path(self.path_file, target)
            elif target.startswith('/'):
              rs.target = basefile + target
            rid = rs.rid
            if not rid.startswith("rId"):
              raise XmlParserException("Invalid format id for rId: " + str(rid))
            try:
              num = max(num, int(rid[3:]))
            except ValueError:
              raise XmlParserException("Invalid format id for rId: " + str(rid))
            self.relations.append(rs)
            self.max_id = num
      except FileNotFoundError:
        pass
    else:
      self.root_tag = XmlTag('Relationships', {'xmlns': RELATIONSHIPS_XMLNS})

  def get_relationships(self, rtype: str or None, rtarget: str or None) -> list:
    """
    Returns relationships filtered by type and/or target.

    Args:
      rtype: Relationship type suffix (or None for all).
      rtarget: Target suffix (or None for all).

    Returns:
      List of Relationship.
    """
    found = []
    for relationship in self.relations:
      if (not rtype or (relationship.type and relationship.type.endswith(rtype)))\
        and (not rtarget or (relationship.target and relationship.target.endswith(rtarget))):
        found.append(relationship)
    return found

  def get_relationship_by_id(self, rid: str) -> Relationship or None:
    """
    Gets a relationship by its rId.

    Args:
      rid: rId identifier.

    Returns:
      Relationship or None if not found.
    """
    for relationship in self.relations:
      if relationship.rid == rid:
        return relationship
    return None

  def add_relationship_image(self, image_ref: str) -> Relationship:
    """
    Adds an image relationship.

    Args:
      image_ref: Image name/file.

    Returns:
      Created or existing Relationship.
    """
    return self.add_relationship(RELATIONSHIP_TYPE_IMAGE, 'media/' + image_ref, None)

  def add_relationship_hyperlink(self, url: str) -> Relationship:
    """
    Adds an external hyperlink relationship.

    Args:
      url: Destination URL.

    Returns:
      Created or existing Relationship.
    """
    return self.add_relationship(RELATIONSHIP_TYPE_HYPERLINK, url, 'External')

  def add_relationship(self, rtype: str, rtarget: str|None, rtargetmode: str|None) -> Relationship:
    """
    Adds a generic relationship.

    Args:
      rtype: Relationship type (suffix or full URL).
      rtarget: Relationship target.
      rtargetmode: TargetMode (e.g., External).

    Returns:
      Created or existing Relationship.
    """
    relationships = self.get_relationships(rtype, rtarget)
    if len(relationships) > 0:
      return relationships[0]
    self.max_id += 1
    attrs = {}
    if not rtype.startswith('http:'):
      rtype = RELATIONSHIP_BASE_URL + rtype

    idx = rtarget.rfind('/xl/')
    if idx >= 0:
      rtarget = rtarget[idx+4:]
    """
    idx = rtarget.find('/')
    if idx > 0:
      next_path = rtarget[:idx]
      if next_path in ['drawings', 'media', 'tables', 'theme']:
        # rtarget = '../' + rtarget
        print("rtarget=" + rtarget)
    """
    add_dict(attrs, 'Id', 'rId' + str(self.max_id))
    add_dict(attrs, 'Type', rtype)
    add_dict(attrs, 'Target', rtarget)
    add_dict(attrs, 'TargetMode', rtargetmode)
    relationship = Relationship(attrs)
    self.relations.append(relationship)
    rtag = self.root_tag.add_tag('Relationship')
    rtag.add_attrs(attrs)
    return relationship

  def add_image(self, name: str, data: bytes,):
    """
    Adds a physical image file to the package.

    Args:
      name: Image file name.
      data: Image bytes.

    Raises:
      OSError: If the file cannot be written.
    """
    path_file = file_util.complete_path(self.path_file, 'media/') + name
    infile_handler = open(path_file, "wb")
    with infile_handler:
      infile_handler.write(data)
