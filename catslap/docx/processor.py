import json
import os
import re
import shutil
import sys
import time
from datetime import timedelta, datetime

from catslap.base.document import Document
from catslap.base.relationships import Relationships
from catslap.base.types import ContentTypes
from catslap.base.utils import dict_repeat_resolver, dict_value_resolver, resolve_param_value, resolve_param_repeating
from catslap.docx import elements as doc_elements
from catslap.docx import word_tags as WT
from catslap.docx.numbering import Numbering
from catslap.docx.styles import Styles
from catslap.utils import encoding as enc_util
from catslap.utils import file as file_util
from catslap.utils import soffice
from catslap.utils import text
from catslap.utils import text as text_util
from catslap.utils import types
from catslap.utils.dotdict import DotDict
from catslap.utils.xml import XmlParser, XmlTag, XmlText, XmlElement, CONFIG_PARAM_INCLUDE_DECL
from catslap.xlsx.document import ExcelDocument, parse_data_ref, get_cell_format_position, get_cell_num, get_row_num



class Processor():
  def __init__(self):
    self.nifs = 0
    self.nfors = 0

  def __process_tag_element(self, element: XmlTag):
    """
    Processes paragraphs resolving directives and placeholders.

    Args:
      element: XML element to process.
    """
    for element in element.elements:
      if not isinstance(element, XmlTag):
        continue
      tagname = element.name
      if tagname == WT.TAG_P:
        self.__process_paragraph(element)
        continue
      if tagname == WT.TAG_TBL:
        self.__process_table(element)
        continue
      self.__process_paragraph(element)

  def __process_paragraph(self, paragraph: XmlTag):
    # -- obtiene los runs
    runs = paragraph.get_tags(WT.TAG_R)
    for run in runs:
      self.__process_run(run)

  def __process_run(self, run: XmlTag):
    # -- obtiene los textos
    texts = run.get_tags(WT.TAG_T)
    for text in texts:
      self.__process_text(text)

  def __process_text(self, text_tag: XmlTag):
    text = text_tag.get_text()
    if not text:
      return    
    idx = 0
    while idx < len(text):
      ch = text[idx]
      if ch == '{' and idx + 1 < len(text):
        if text[idx + 1] == '%':
          idx += 2
          idx0 = idx
          while idx < len(text) and text[idx:idx+2] != '%}':
            idx += 1
          nifs = self.nifs
          text = self.__process_directive(text_tag, text[idx0:idx])
          idx += 2
          text_tag.set_text(text[:idx0] + text[idx:])
          if nifs 
          continue
        if text[idx + 1] == '{':
          idx += 2
          idx0 = idx
          while idx < len(text) and text[idx:idx+2] != '}}':
            idx += 1
          resolved = self.__process_dynamic_value(text_tag, text[idx0:idx])
          idx += 2
          text_tag.set_text(text[:idx0] + resolved + text[idx:])
          continue
        idx += 1

  def __process_directive(self, text_tag: XmlTag, text: str):
    idx = text.find(' ')
    keyword = text[0:idx].strip() if idx >= 0 else text
    arg = text[idx + 1:].strip() if idx >= 0 else None
    if keyword.equals(KEYWORD_IF):
      self.nifs += 1
      if self.value_resolver(None, arg) is not True:
        #-- elimina los bloques hasta el end-if
        self.__delete_blocks_until_endif(text_tag.parent, text_tag)
        #-- elimina el texto que queda
        return False;
      #-- procesa los bloques hasta el end-if
      self.__process_blocks_until_endif(text_tag)
      #-- mantiene el texto que queda
      return True
    if keyword.equals(KEYWORD_ENDIF):
      self.nifs -= 1
    if keyword.equals(KEYWORD_FOR):
      self.fors += 1
    if keyword.equals(KEYWORD_FOR):
      self.nfors -= 1
    return True

  def __process_dynamic_value(self, text_tag: XmlTag, text: str) -> str:
    return text
  
  def __find_tag_child(self, tag: XmlTag, child: XmlElement|None) -> int:
    elements = tag.elements
    idx = 0
    if child:
      while idx < len(elements):
        if elements[idx] == child:
          return idx + 1
        idx += 1
    return idx

  def __delete_blocks_until_endif(self, from_tag: XmlTag, child_tag: XmlTag|None):
    idx = self.__find_tag_child(from_tag, child_tag) if child_tag else 0
    elements = from_tag.elements
    while idx < len(elements):
      tag = elements[idx]
      if not isinstance(tag, XmlTag):
        del elements[idx]
        continue
      tagname = tag.name
      if tagname != WT.TAG_T:
        self.__delete_blocks_until_endif(tag, None)
        if len(tag.elements) == 0:
          del elements[idx]
        continue
      self.__process_text(tag)
      if self.nifs == 0:
        break
      idx += 1

