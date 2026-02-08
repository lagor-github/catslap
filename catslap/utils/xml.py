# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

from abc import abstractmethod

from catslap.utils import encoding as enc_util
from catslap.utils import file as file_util
from catslap.utils import text as text_util
from catslap.utils import types
from catslap.utils import utils



class XmlException(Exception):
  """
  Base XML processing exception.
  """
  pass


class XmlParserException(XmlException):
  """
  XML parser specific exception.
  """
  pass


CONFIG_PARAM_STRICT = "STRICT"  # -- opciones estrictas de XML
CONFIG_PARAM_HTML = "HTML"  # -- ignora los tags cerrados sin que hayan abierto previamente (True) o error (False)
CONFIG_PARAM_AUTO_END_TAGS = "AUTO_END_TAGS"  # -- tags que deben o pueden ser autocerrados (string list)
CONFIG_PARAM_PRETTY_OUTPUT = "PRETTY_OUTPUT"
CONFIG_PARAM_INCLUDE_DECL = "INCLUDE_DECL"
HTML_AUTO_END_TAGS = ['img', 'br', 'meta', 'input']


xmlWriterDefaultConfig = {
    CONFIG_PARAM_AUTO_END_TAGS: HTML_AUTO_END_TAGS,
    CONFIG_PARAM_PRETTY_OUTPUT: True,
    CONFIG_PARAM_INCLUDE_DECL: False
}

xmlParserDefaultConfig = {
    CONFIG_PARAM_AUTO_END_TAGS: HTML_AUTO_END_TAGS,
    CONFIG_PARAM_HTML: True,
    CONFIG_PARAM_STRICT: False,
}


class XmlWriter:
  """
  XML writer with pretty print support.
  """
  def __init__(self, config: dict | None = None):
    self.out = []
    self.config = types.merge_dicts(config, xmlWriterDefaultConfig) if isinstance(config,
                                                                                      dict) else xmlWriterDefaultConfig
    if self.config[CONFIG_PARAM_INCLUDE_DECL]:
      self.out.append('<?xml version="1.0" encoding="UTF-8"?>')
      if self.config[CONFIG_PARAM_PRETTY_OUTPUT]:
        self.out.append('\n')

  def write(self, content: str):
    """
    Writes XML/HTML content to the buffer.

    Args:
      content: Text, XmlElement, or list of elements to serialize.
    """
    if isinstance(content, list):
      for item in content:
        if isinstance(item, XmlElement):
          item.write(0, self)
      return
    if isinstance(content, XmlElement):
      content.write(0, self)
      return
    if not isinstance(content, str):
      return
    self.out.append(content)

  def __str__(self) -> str:
    return ''.join(self.out)


class XmlElement:

  def __init__(self):
    self.parent = None

  """
  XML/HTML element.
  """
  @abstractmethod
  def write(self, indent: int, writer: XmlWriter):
    """
    Serializes the element.

    Args:
      indent: Indentation level.
      writer: Output writer.
    """
    pass

  @abstractmethod
  def clone(self, deep: bool = True) -> "XmlElement":
    """
    Clones the element.

    Args:
      deep: If True, recursively clones children.

    Returns:
      Element copy.
    """
    pass

  def to_xml(self) -> str:
    """
    Returns the serialized XML of the element.

    Returns:
      XML string.
    """
    writer = XmlWriter()
    self.write(0, writer)
    return str(writer)

  @abstractmethod
  def to_json(self) -> any:
    """
    Converts the element to a JSON representation.

    Returns:
      Equivalent JSON structure.
    """
    return None


class XmlText(XmlElement):
  """
  XML/HTML text node.
  """
  def __init__(self, content: str):
    self.content = content

  def append(self, text: str):
    """
    Appends text to the current node.

    Args:
      text: Text to concatenate.
    """
    self.content += text

  def write(self, indent: int, writer: XmlWriter):
    """
    Writes the text escaping entities.

    Args:
      indent: Indentation level (unused for text).
      writer: Output writer.
    """
    writer.write(XmlParser.escape_entities(self.content))

  def to_json(self) -> any:
    """
    Converts text to JSON.

    Returns:
      Dictionary with the 'text' key.
    """
    return {'text': self.content}

  def clone(self, deep: bool = True) -> XmlElement:
    """
    Clones the text node.

    Args:
      deep: Ignored (no children).

    Returns:
      New XmlText instance.
    """
    xml = XmlText(self.content)
    xml.parent = self.parent
    return xml


class XmlTag(XmlElement):
  """
  XML/HTML tag.
  """
  def __init__(self, tagname: str, attrs: dict | None = None):
    self.name = tagname
    self.attrs = attrs.copy() if isinstance(attrs, dict) else {}
    self.elements = []

  def clone(self, deep: bool = True) -> "XmlTag":
    """
    Clones the tag.

    Args:
      deep: If True, recursively clones children.

    Returns:
      Tag copy.
    """
    xml = XmlTag(self.name, self.attrs.copy())
    xml.parent = self.parent
    if deep:
      for item in self.elements:
        xml.add_element(item.clone())
    return xml

  def get_attr(self, attrname: str) -> str | None:
    """
    Gets an attribute by name.

    Args:
      attrname: Attribute name.

    Returns:
      Attribute value or None.
    """
    return self.attrs.get(attrname)

  def get_attr_int(self, attrname: str) -> int | None:
    """
    Gets an attribute and converts it to int if possible.

    Args:
      attrname: Attribute name.

    Returns:
      Integer or None.
    """
    value = self.get_attr(attrname)
    if value:
      value = value.strip()
      try:
        return int(value)
      except ValueError:
        pass
    return None

  def set_attr(self, attrname: str, attrvalue):
    """
    Sets an attribute.

    Args:
      attrname: Attribute name.
      attrvalue: Attribute value.
    """
    self.attrs[attrname] = attrvalue

  def remove_attr(self, attrname: str):
    """
    Removes an attribute if it exists.

    Args:
      attrname: Attribute name.
    """
    if self.attrs.get(attrname):
      del self.attrs[attrname]

  def add_attr(self, attrname: str, attrvalue: str | None):
    """
    Adds or updates an attribute.

    Args:
      attrname: Attribute name.
      attrvalue: Attribute value.
    """
    self.attrs[attrname] = attrvalue

  def add_attrs(self, attrs: dict):
    """
    Adds multiple attributes.

    Args:
      attrs: Attribute dictionary.
    """
    for attrname, attrvalue in attrs.items():
      self.attrs[attrname] = attrvalue

  def add_element(self, item: XmlElement) -> XmlElement:
    """
    Adds a child element.

    Args:
      item: XML element.

    Returns:
      Added element.

    Raises:
      XmlParserException: If the type is not XmlElement.
    """
    if not isinstance(item, XmlElement):
      raise XmlParserException("Invalid XML element type: " + str(type(item)))
    self.elements.append(item)
    item.parent = self
    return item

  def add_tag(self, tag: str or "XmlTag", attrs: dict | None = None) -> "XmlTag":
    """
    Creates or adds a child tag.

    Args:
      tag: Tag name or XmlTag.
      attrs: Optional attributes if tag is str.

    Returns:
      Added tag.

    Raises:
      XmlParserException: If the type is not XmlTag.
    """
    if isinstance(tag, str):
      tag = XmlTag(tag, attrs)
    if not isinstance(tag, XmlTag):
      raise XmlParserException("Invalid XML element type: " + str(type(tag)))
    self.add_element(tag)
    tag.parent = self
    return tag

  def add_tag_text(self, tag: str or "XmlTag", text: str|XmlText) -> "XmlTag":
    """
    Adds a child tag with text.

    Args:
      tag: Tag name or XmlTag.
      text: Text or XmlText.

    Returns:
      Created tag.
    """
    tag = self.add_tag(tag)
    tag.add_text(text)
    return tag

  def set_text(self, text: str|XmlText) -> XmlText | None:
    """
    Replaces content with a single text node.

    Args:
      text: Text or XmlText.

    Returns:
      Created text node.
    """
    self.elements = []
    return self.add_text(text)

  def add_text(self, text: str|XmlText) -> XmlText | None:
    """
    Adds text as a child of the tag.

    Args:
      text: Text or XmlText.

    Returns:
      Added text node.
    """
    if isinstance(text, XmlText):
      return self.add_element(text)
    if text is None:
      text = ''
    text = str(text)
    tags = self.elements
    text = XmlParser.resolve_entities(text)
    last_tag = tags[len(tags) - 1] if len(tags) > 0 else None
    if last_tag and isinstance(last_tag, XmlText):
      last_tag.append(text)
      return last_tag
    return self.add_element(XmlText(text))

  def write(self, indent: int, writer: "XmlWriter"):
    """
    Writes the tag and its content to the writer.

    Args:
      indent: Indentation level.
      writer: Output writer.
    """
    has_indent = writer.config[CONFIG_PARAM_PRETTY_OUTPUT]
    if has_indent:
      writer.write(utils.repeat(' ', indent * 2))
    writer.write('<' + self.name)
    if len(self.attrs) > 0:
      for attrname, attrvalue in self.attrs.items():
        writer.write(' ' + attrname)
        if attrvalue is not None:
          writer.write('="' + XmlParser.escape_attr_value(attrvalue) + '"')
    if len(self.elements) == 0:
      auto_end_tags = writer.config[CONFIG_PARAM_AUTO_END_TAGS]
      if auto_end_tags and self.name in auto_end_tags:
        writer.write('>')
      else:
        writer.write('/>')
      if has_indent:
        writer.write('\n')
      return
    writer.write('>')
    if len(self.elements) == 1 and isinstance(self.elements[0], XmlText):
      self.elements[0].write(indent + 1, writer)
      writer.write('</' + self.name + '>')
      if has_indent:
        writer.write('\n')
      return
    if has_indent:
      writer.write('\n')
    for tag in self.elements:
      tag.write(indent + 1, writer)
    if has_indent:
      writer.write(utils.repeat(' ', indent * 2))
    writer.write('</' + self.name + '>')
    if has_indent:
      writer.write('\n')

  def to_json(self) -> dict:
    """
    Converts the tag to JSON.

    Returns:
      Dictionary with name, attributes, and elements.
    """
    jmap = {'tag': self.name}
    if len(self.attrs) > 0:
      jmap['attrs'] = self.attrs
    if len(self.elements) > 0:
      tags = []
      for elem in self.elements:
        tags.append(elem.to_json())
      jmap['tags'] = tags
    return jmap

  def get_tag_path(self, paths: list, mandatory: bool = True) -> any:
    """
    Gets a tag by walking a name path.

    Args:
      paths: List of tag names (supports wildcards with '*').
      mandatory: If True, raises if missing.

    Returns:
      Found tag or None.

    Raises:
      XmlParserException: If a required tag is not found.
    """
    tag = self
    for item in paths:
      idx = item.find('*')
      if idx < 0:
        tag = tag.get_tag(item, mandatory)
        if tag is None:
          return None
        continue
      selected = None
      starts = item[0:idx]
      ends = item[idx + 1:]
      for elem in tag.elements:
        if not isinstance(elem, XmlTag):
          continue
        if (starts != '' and elem.name.startswith(starts)) or \
        (ends != '' and elem.name.endswith(ends)):
          selected = elem
          break
      if not selected:
        raise XmlParserException(f"Tag '{item}' nor found")
      tag = selected
    return tag

  def clear_tags(self):
    """
    Removes all children from the tag.
    """
    self.elements = []

  def get_tag(self, tag_name: str | None = None, mandatory: bool = True) -> "XmlTag":
    """
    Gets the first child tag with the given name.

    Args:
      tag_name: Tag name or None for the first.
      mandatory: If True, raises if missing.

    Returns:
      Found tag or None.

    Raises:
      XmlParserException: If missing and mandatory is True.
    """
    def __raise_exception():
      ex_name = '<' + tag_name + '>' if tag_name is not None else 'XML'
      raise XmlParserException(ex_name + ' tag expected below <' + self.name + '> tag')

    elements = self.elements
    for element in elements:
      if not isinstance(element, XmlTag):
        continue
      if tag_name is None or tag_name == element.name:
        return element
    if mandatory:
      __raise_exception()
    return None

  def remove_tag(self, tag_name: str):
    """
    Removes the first child tag with the given name.

    Args:
      tag_name: Tag name to remove.
    """
    idx = 0
    while idx < len(self.elements):
      elem = self.elements[idx]
      if not isinstance(elem, XmlTag):
        idx += 1
        continue
      if elem.name == tag_name:
        del self.elements[idx]
        return
      idx += 1

  def remove(self, elem0: XmlElement):
    """
    Removes a child element by reference.

    Args:
      elem0: Element to remove.
    """
    idx = 0
    while idx < len(self.elements):
      elem = self.elements[idx]
      if elem == elem0:
        del self.elements[idx]
        return
      idx += 1

  def remove_tags(self, tag_name: str):
    """
    Removes all child tags with the given name.

    Args:
      tag_name: Tag name to remove.
    """
    idx = 0
    while idx < len(self.elements):
      elem = self.elements[idx]
      if not isinstance(elem, XmlTag):
        idx += 1
        continue
      if elem.name == tag_name:
        del self.elements[idx]
        continue
      idx += 1

  def get_tag_text(self, tag_name: str, mandatory: bool = True) -> str | None:
    """
    Gets the text of the first child tag with the given name.

    Args:
      tag_name: Tag name.
      mandatory: If True, raises if missing.

    Returns:
      Tag text or None.
    """
    tag = self.get_tag(tag_name, mandatory)
    if tag is None:
      return None
    return tag.get_text()

  def get_tag_attr(self, tag_name: str, tag_attr: str, mandatory: bool = True) -> str | None:
    """
    Gets an attribute from the first child tag with the given name.

    Args:
      tag_name: Child tag name.
      tag_attr: Attribute name.
      mandatory: If True, raises if missing.

    Returns:
      Attribute value or None.
    """
    tag = self.get_tag(tag_name, mandatory)
    if tag is None:
      return None
    return tag.get_attr(tag_attr)

  def set_tag_text(self, tag_name: str, tag_value: str, mandatory: bool = True):
    """
    Sets text on the first child tag with the given name.

    Args:
      tag_name: Child tag name.
      tag_value: Text to set.
      mandatory: If True and missing, creates the tag.
    """
    tag = self.get_tag(tag_name, mandatory)
    if tag is None:
      tag = XmlTag(tag_name)
      self.add_tag(tag)
    tag.set_text(tag_value)

  def get_tags(self, tag_name: str | None = None) -> list:
    """
    Gets all child tags with the given name.

    Args:
      tag_name: Tag name or None for all.

    Returns:
      List of XmlTag.
    """
    tags = []
    elements = self.elements
    for block in elements:
      if not isinstance(block, XmlTag):
        continue
      if tag_name is not None and block.name != tag_name:
        continue
      tags.append(block)
    return tags

  def get_text(self) -> str:
    """
    Gets the text of the first text node.

    Returns:
      Node text.
    """
    text_node = self.get_text_node()
    return text_node.content

  def get_text_node(self, mandatory: bool = True) -> XmlText:
    """
    Gets the first text node.

    Args:
      mandatory: If True, raises if there is no text.

    Returns:
      XmlText (empty if missing and mandatory is False).
    """
    elements = self.elements
    if len(elements) == 0:
      return XmlText('')
    if not isinstance(elements[0], XmlText):
      if mandatory:
        raise XmlParserException('Text expected into <' + self.name + '> tag: ' + self.to_xml())
      return XmlText('')
    return elements[0]

  def get_inner_html(self) -> str:
    """
    Gets the inner HTML/XML of the tag.

    Returns:
      Serialized string of child elements.
    """
    writer = XmlWriter()
    for item in self.elements:
      item.write(0, writer)
    return str(writer)

class XmlParser:
  """
  XML/HTML parser with flexible configuration.
  """

  def __init__(self, config: dict = None):
    self.config = types.merge_dicts(config, xmlParserDefaultConfig) if isinstance(config, dict) else xmlParserDefaultConfig
    self.max_id = 0
    self.pathfile = None
    self.root_tag = None

  def parse_file(self, pathfile: str, roottag: str | None = None) -> XmlTag:
    """
    Parses an XML file and returns the root tag.

    Args:
      pathfile: XML file path.
      roottag: Expected root tag name.

    Returns:
      Parsed root tag.

    Raises:
      FileNotFoundError: If the file does not exist.
      XmlParserException: If XML is invalid or root mismatch.
    """
    self.pathfile = pathfile
    xml_content = text_util.trim(str(file_util.read_bytes(pathfile), enc_util.UTF_8))
    if xml_content.startswith('<?'):
      idx = xml_content.find('?>')
      if idx < 0:
        raise XmlParserException(": XML mark '...?>' not found")
      xml_content = text_util.trim(xml_content[idx + 2:])
    blocks = self.parse_text(xml_content)
    if len(blocks) < 1 or not isinstance(blocks[0], XmlTag):
      raise XmlParserException("None XML tag defined")
    block = blocks[0]
    if roottag:
      tag = block.name
      if tag != roottag:
        raise XmlParserException("XML root tag <" + roottag + " ...> was expected at start. Found <" + block.name + "> tag")
    self.root_tag = block
    return block

  def parse_text(self, text: str) -> list:
    """
    Parses XML/HTML from text.

    Args:
      text: XML/HTML text.

    Returns:
      List of parsed elements.

    Raises:
      XmlParserException: If the content is invalid.
    """
    self.root_tag = XmlTag('#root')
    self.__parse_tags(self.root_tag, text, 0, 0)
    return self.root_tag.elements

  def __parse_tags(self, parent_tag: XmlTag, row: str, idx: int, recurrent: int = 0) -> int:
    config_strict = self.config[CONFIG_PARAM_STRICT]
    config_html = self.config[CONFIG_PARAM_HTML]
    config_autoendtags = self.config[CONFIG_PARAM_AUTO_END_TAGS]
    opentag = parent_tag.name
    preserve = True
    while idx < len(row):
      idx0 = idx
      idx = row.find('<', idx0)
      if idx < 0:
      # -- texto como resto del doc
        text = row[idx0:]
        if not preserve:
          text = text_util.rtrim(text)
        _add_preserving_text(parent_tag, text, preserve)
        break
      if idx + 1 == len(row):
        if config_strict:
          XmlParser.__raise_xml_parse_exception(row, idx, "Invalid tag start mark at end of content")
          # -- texto como resto del doc (finaliza en <)
        text = row[idx0:]
        _add_preserving_text(parent_tag, text, preserve)
        break
        # -- texto hasta siguiente ...<tag> o ...</tag>
      text = row[idx0:idx]
      endtag = row[idx + 1] == '/'
      if endtag:
        idx += 1
        # -- caso especial '< ' o '</ '
      if not text_util.is_alpha(row[idx + 1]):
        idx += 1
        continue
        # -- toma el nombre del tag
      idx += 1
      idx0 = idx
      while idx < len(row) and row[idx] != '>' and row[idx] != ' ' and row[idx] != '/':
        idx += 1
      tag_name = row[idx0:idx]
      if tag_name == '':
        XmlParser.__raise_xml_parse_exception(row, idx0, "Empty tag name: <?>")
      idx2 = row.find(">", idx)
      if idx2 < 0:
        XmlParser.__raise_xml_parse_exception(row, idx0, f"End of tag mark expected: <{tag_name}...?")
      tag = XmlTag(tag_name)
      # -- parse atributos
      self.__parse_attrs(tag, row, idx, idx2) if not endtag else None
      autoendtag = row[idx2 - 1] == '/'
      if endtag and autoendtag:
        XmlParser.__raise_xml_parse_exception(row, idx0, f"Invalid autoend tag for tag end:  </{tag_name}/?")
      idx2 += 1
      idx0 = idx2
      # -- final de tag del tag de entrada, retorna
      if endtag and tag_name == opentag:
        _add_preserving_text(parent_tag, text, preserve)
        return idx0
      if not preserve:
        text = text_util.ltrim(text)
      if tag_name in config_autoendtags or autoendtag:
        _add_preserving_text(parent_tag, text, preserve)
        idx = idx0
        parent_tag.add_tag(tag)
        continue
      _add_preserving_text(parent_tag, text, preserve)
      parent_tag.add_tag(tag)
      if endtag:
        if config_strict and not config_html:
          XmlParser.__raise_xml_parse_exception(row, idx0, f"Invalid tag end </{tag_name}> without tag start. Expected </{opentag}>")
        idx = idx0
        # -- ignora tag
        continue
      idx = self.__parse_tags(tag, row, idx0, recurrent + 1)
    return idx

  def __parse_attrs(self, tag: XmlTag, row: str, idx: int, idx2: int):
    config_strict = self.config[CONFIG_PARAM_STRICT]
    tag_name = tag.name
    while idx < idx2:
    # -- ignora espacios
      while idx < idx2 and row[idx] in [' ', '\t']:
        idx += 1
        # -- coge nombre de attributo
      idx0 = idx
      while idx < idx2 and row[idx] not in [' ', '\t', '=', '/', '>']:
        idx += 1
      attr_name = row[idx0:idx]
      if attr_name == '':
        break
      if attr_name.find('"') >= 0 or attr_name.find("'") >= 0:
        XmlParser.__raise_xml_parse_exception(row, idx, f"Invalid attr name: <{tag_name} {attr_name}...")
      while idx < idx2 and row[idx] in [' ', '\t']:
        idx += 1
        # -- ignora el igual y coge el valor
      if idx >= idx2 or row[idx] != '=':
        if config_strict:
          XmlParser.__raise_xml_parse_exception(row, idx, f"Equals sign expected for attr: <{tag_name} {attr_name}=?...")
        tag.add_attr(attr_name, None)
        continue
      idx += 1
      while idx < idx2 and row[idx] in [' ', '\t']:
        idx += 1
      quote = row[idx]
      if quote != '"' and quote != "'":
        XmlParser.__raise_xml_parse_exception(row, idx, f"Quote expected around attribute from tag. Found '{quote}' char: <{tag_name} {attr_name}=?>")
      idx += 1
      idx0 = idx
      while idx < idx2 and row[idx] != quote:
        idx += 1
      quote2 = row[idx]
      attr_value = row[idx0:idx]
      if quote2 != quote:
        XmlParser.__raise_xml_parse_exception(row, idx, f"Quote expected around attribute from tag. Found '{quote2}' char: <{tag_name} {attr_name}={quote}{attr_value}{quote2}? ...>")
        continue
      idx += 1
      tag.add_attr(attr_name, XmlParser.resolve_entities(attr_value))
      if attr_name == 'id':
        try:
          idvalue = int(attr_value)
          if idvalue > self.max_id:
            self.max_id = idvalue
        except ValueError:
          pass

  def write_file(self):
    """
    Writes the root XML to the associated file.

    Raises:
      OSError: If the file cannot be written.
    """
    if self.pathfile and self.root_tag:
      writer = XmlWriter({CONFIG_PARAM_INCLUDE_DECL: True})
      self.root_tag.write(0, writer)
      content = str(writer)
      file_util.write_bytes(self.pathfile, content.encode(enc_util.UTF_8))

  @staticmethod
  def __calc_number_of_line(row: str, idx: int):
    line = 1
    if idx >= len(row):
      idx = len(row) - 1
    while idx >= 0:
      if row[idx] == '\n':
        line += 1
      idx -= 1
    return line

  @staticmethod
  def __raise_xml_parse_exception(row: str, idx: int, message: str):
    line = XmlParser.__calc_number_of_line(row, idx)
    raise XmlParserException(f"[line: {line}] {message}")

  @staticmethod
  def get_xml(blocks: list | None, config: dict | None = None) -> str:
    """
    Serializes elements to XML without pretty print.

    Args:
      blocks: List of elements.
      config: Additional configuration.

    Returns:
      Generated XML.
    """
    return XmlParser.__get_xml(blocks, types.merge_dicts({
      CONFIG_PARAM_PRETTY_OUTPUT: False
    }, config))

  @staticmethod
  def get_outer_xml(tag: XmlTag, config: dict | None = None) -> str:
    """
    Serializes a full tag to XML.

    Args:
      tag: Tag to serialize.
      config: Additional configuration.

    Returns:
      Generated XML.
    """
    return XmlParser.__get_xml(tag, types.merge_dicts({
      CONFIG_PARAM_PRETTY_OUTPUT: False
    }, config))

  @staticmethod
  def get_pretty_xml(blocks: list | None, config: dict | None = None) -> str:
    """
    Serializes elements to XML with pretty print.

    Args:
      blocks: List of elements.
      config: Additional configuration.

    Returns:
      Generated XML.
    """
    return XmlParser.__get_xml(blocks, types.merge_dicts({
      CONFIG_PARAM_PRETTY_OUTPUT: True
    }, config))

  @staticmethod
  def __get_xml(blocks: XmlElement | list | None, xml_writer_config: dict | None) -> str:
    if blocks is None:
      return ''
    out = XmlWriter(xml_writer_config)
    if isinstance(blocks, XmlElement):
      blocks.write(0, out)
      return str(out)
    if isinstance(blocks, list):
      for block in blocks:
        block.write(0, out)
      return str(out)
    raise XmlException("Invalid Xml argument: " + str(type(blocks)))

  @staticmethod
  def escape_entities(row) -> str:
    """
    Escapes basic XML entities.

    Args:
      row: Input text.

    Returns:
      Escaped text.
    """
    if row is None:
      return ''
    if isinstance(row, int) or isinstance(row, float):
      row = str(row)
    if isinstance(row, str):
      row = row.replace('&', '&amp;')
      row = row.replace('<', '&lt;')
      row = row.replace('>', '&gt;')
      return row
    return ''

  @staticmethod
  def escape_attr_value(row) -> str:
    """
    Escapes entities in attribute values.

    Args:
      row: Attribute value.

    Returns:
      Escaped value.
    """
    if isinstance(row, int) or isinstance(row, float):
      row = str(row)
    if isinstance(row, str):
      row = row.replace('&', '&amp;')
      row = row.replace('"', '&quot;')
      return row
    return ''

  @staticmethod
  def resolve_entities(row) -> str:
    """
    Resolves basic XML entities.

    Args:
      row: Input text.

    Returns:
      Text with resolved entities.
    """
    if isinstance(row, int) or isinstance(row, float):
      row = str(row)
    if isinstance(row, str):
      row = row.replace('&#xD;', '\r')
      row = row.replace('&#xA;', '\n')
      row = row.replace('&lt;', '<')
      row = row.replace('&gt;', '>')
      row = row.replace('&quot;', '"')
      row = row.replace('&apos;', "'")
      row = row.replace('&nbsp;', ' ')
      row = row.replace('&amp;', '&')  # -- tiene que ser al final para que no haya auto-resoluciones
      return row
    return ''

  @staticmethod
  def compose_autoclosed_tag(tagname: str, attrs: dict):
    """
    Composes a self-closing tag.

    Args:
      tagname: Tag name.
      attrs: Tag attributes.

    Returns:
      Self-closing XML string.
    """
    out = '<' + tagname
    if not attrs or len(attrs) == 0:
      out += '/>'
      return out
    for key, value in attrs.items():
      out += ' ' + key + '="' + value + '"'
    out += '/>\r\n'
    return out


def _add_preserving_text(parent_tag: XmlTag, text: str | None, preserve: bool):
  if text is None:
    return
  if not preserve:
    text = text.replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
    text = text_util.trim(text.replace('  ', ' '))
  if text == '':
    return
  parent_tag.add_text(text)
