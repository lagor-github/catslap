# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)

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


CFG_PARAM_OUTPUT_FORMAT_TYPE = 'output_format_type'
CFG_PARAM_REINDEX = 'reindex'

# -- puntos de powerpoint por cm
PT_PER_CM = 360000

WORD_DOCUMENT = "word/document.xml"
WORD_DOCUMENT_RELS = "word/_rels/document.xml.rels"
WORD_STYLES = "word/styles.xml"
WORD_NUMBERING = "word/numbering.xml"
WORD_DOCUMENT_TYPES = "[Content_Types].xml"
WORD_CHARTS = "word/charts/"
WORD_CHARTS_RELS = WORD_CHARTS + "_rels/"


IGNORABLE_TAGS = ['w:lastRenderedPageBreak', 'w:proofErr', 'w:noProof', 'w:lang', 'w:bookmarkStart', 'w:bookmarkEnd']
IGNORABLE_EMPTY_TAGS = ['w:rPr', 'w:pPr']

KEYWORD_IF = 'if'
KEYWORD_ELSE = 'else'
KEYWORD_ENDIF = 'endif'
KEYWORD_FOR = 'for'
KEYWORD_ENDFOR = 'endfor'
KEYWORD_STYLE = 'style'


class ProcessStatus:
  def __init__(self):
    self.nifs = 0
    self.nfors = 0
    self.else_mode = False
    self.search = None

  def process_directive_keyword(self, keyword: str):
    if keyword == KEYWORD_IF:
      if self.search is None and self.nifs == 0:
        self.search = keyword
      self.nifs += 1
    elif keyword == KEYWORD_ENDIF:
      self.nifs -= 1
    elif keyword == KEYWORD_ELSE:
      if self.nifs == 1:
        self.else_mode = True
    elif keyword == KEYWORD_FOR:
      if self.search is None and self.nfors == 0:
        self.search = keyword
      self.nfors += 1
    elif keyword == KEYWORD_ENDFOR:
      self.nfors -= 1
    if self.search == KEYWORD_IF and self.nifs == 0:
      return True
    if self.search == KEYWORD_FOR and self.nfors == 0:
      return True
    return False


class WordDocument(Document):
  """
  Processes and generates Word documents (.docx/.docm) using a template file.

  Args:
    template: Template word file

      Attributes:
    max_id: Maximum ID used in the document.
    types: Package ContentTypes.
    relationships: Document relationships.
    styles: Loaded styles.
    numbering: Numbering manager.
    num_tables: Inserted table counter.
  """
  def __init__(self, template_file: str):
    super().__init__(template_file)
    self.max_id = 0
    self.types = None
    self.relationships = None
    self.styles = None
    self.num_tables = 0;
    self.process_keyword = None
    self.process_deep = 1

  def process_template(self, tempdir: str):
    """
    Processes the DOCX template in the temporary directory.

    Args:
      tempdir: Temporary directory with ZIP contents.

    Raises:
      XmlParserException: If XML is invalid.
      OSError: If file read/write fails.
    """
    word_rel = tempdir + "/" + WORD_DOCUMENT_RELS
    self.relationships = Relationships(tempdir, word_rel)

    word_styles = tempdir + "/" + WORD_STYLES
    self.styles = Styles(word_styles)
    self.styles.find_styles()

    numbering = tempdir + "/" + WORD_NUMBERING
    self.numbering = Numbering(numbering)

    word_types = tempdir + "/" + WORD_DOCUMENT_TYPES
    self.types = ContentTypes(word_types)

    self.__process_word_ext_files(tempdir, 'header', WT.TAG_HDR)
    self.__process_word_ext_files(tempdir, 'footer', WT.TAG_FTR)

    self.__process_word_file(tempdir)
    if not self.test_mode:
      self.types.write_file()
      self.relationships.write_file()
      self.numbering.write_file()

  def __process_word_file(self, tempdir: str):
    word_file = tempdir + "/" + WORD_DOCUMENT
    xml = XmlParser()
    document = xml.parse_file(word_file, WT.TAG_DOCUMENT)
    body = document.get_tag(WT.TAG_BODY)
    blocks = body.elements
    self.collapse_paragraphs(blocks)
    self.max_id = xml.max_id
    self.search_graphic_frames(tempdir, blocks)
    self.process_descr_attrs(blocks)
    self.process_paragraphs(body)
    self.expand_content(body)
    xml_content = xml.get_pretty_xml(document, {CONFIG_PARAM_INCLUDE_DECL: True})
    file_util.write_bytes(word_file, bytes(xml_content, enc_util.UTF_8))
    file_util.write_bytes("/tmp/word.txt", bytes(xml_content, enc_util.UTF_8))

    # -- procesa los archivos de excel
    files = file_util.list_files(tempdir + '/' + WORD_CHARTS_RELS)
    if files:
      for file in files:
        chart_file = tempdir + '/' + WORD_CHARTS_RELS + file
        if file.endswith('.xml.rels'):
          chart_name = file[0:len(file) - 5]
          relationships = Relationships(tempdir, chart_file)
          relations = relationships.get_relationships(None, ".xlsx")
          if len(relations) == 0:
            continue
          relation = relations[0]
          target = relation.target
          self.__process_chart_excel_file(chart_name, tempdir, target)

  def __process_word_ext_files(self, tempdir: str, file_name: str, tag_name: str):
    xml = XmlParser()
    # -- parsea los archivos adjuntos para actualizar el atributo id
    idx = 1
    found = True
    while found:
      file0 = tempdir + "/word/" + file_name + str(idx) + ".xml"
      found = file_util.exist(file0)
      if found:
        document = xml.parse_file(file0, tag_name)
        blocks = document.elements
        self.collapse_paragraphs(blocks)
        self.max_id = max(self.max_id, xml.max_id)
        self.process_paragraphs(document)
        self.expand_content(document)
        xml_content = xml.get_pretty_xml(document, {CONFIG_PARAM_INCLUDE_DECL: True})
        file_util.write_bytes(file0, bytes(xml_content, enc_util.UTF_8))
        file_util.write_bytes("/tmp/" + file_name + str(idx) + ".txt", bytes(xml_content, enc_util.UTF_8))
        idx += 1

  def collapse_paragraphs(self, elements: list, rep: int = 0):
    """
    Collapses runs and removes ignorable tags in an XML tree.

    Args:
      elements: List of XML elements.
      rep: Recursion level (internal use).
    """
    idx = 0
    last_rpr = None
    last_t = None
    while idx < len(elements):
      element = elements[idx]
      if not isinstance(element, XmlTag):
        idx += 1
        continue

      self.collapse_paragraphs(element.elements, rep + 1)

      tag = element
      tag_name = tag.name
      tags = tag.elements
      # -- ignora el tag
      if tag_name in IGNORABLE_TAGS or (len(tags) == 0 and tag_name in IGNORABLE_EMPTY_TAGS):
        del elements[idx]
        continue
        # -- funde todos los w:r que sean iguales en un solo w:r
      if tag_name == WT.TAG_R:
        rpr = tag.get_tag(WT.TAG_RPR, False)
        t = tag.get_tag(WT.TAG_T, False)
        if len(tags) == 0 or t is None:
          idx += 1
          last_t = t
          continue
        if last_t is not None and WordDocument.__is_the_same_rpr(last_rpr, rpr):
          ct = t.elements[0].content if len(t.elements) > 0 and isinstance(t.elements[0], XmlText) else None
          if ct is not None:
            if len(last_t.elements) == 0:
              last_t.add_text(ct)
            else:
              last_t.elements[0].append(ct)
            last_t.attrs['xml:space'] = 'preserve'
            # -- elimina el párrafo actual
            del elements[idx]
            continue
          last_rpr = None
          last_t = None
          idx += 1
          continue
        last_rpr = rpr
        last_t = t
      idx += 1

  @staticmethod
  def __is_the_same_rpr(rpr1: dict | None, rpr2: dict | None) -> bool:
    if rpr1 is None and rpr2 is None:
      return True
    xml = XmlParser()
    dump1 = xml.get_xml(rpr1)
    dump2 = xml.get_xml(rpr2)
    return dump1 == dump2

  def process_paragraphs(self, tag: XmlTag):
    """
    Processes paragraphs resolving directives and placeholders.

    Args:
      element: XML element to process.
    """
    elements = tag.elements
    idx = 0
    while idx < len(elements):
      idx = self.__process_paragraph(elements, idx)

  def __process_paragraph(self, elements: list, idx: int):
    element = elements[idx]
    if not isinstance(element, XmlTag):
      del elements[idx]
      return idx
    tagname = element.name
    if tagname == WT.TAG_P:
      keyword, condition = self.__get_directive(element)
      if keyword is not None:
        if keyword == KEYWORD_STYLE:
          self.__process_style_directive(condition)
          del elements[idx]
          return idx
        if keyword == KEYWORD_IF:
          if_value = self.resolve_value(None, condition)
          if isinstance(if_value, bool):
            is_true = if_value is True
          elif isinstance(if_value, int):
            is_true = if_value != 0
          elif isinstance(if_value, list):
            is_true = len(if_value) > 0
          else:
            is_true = str(if_value) != ''
          self.__get_blocks_until_endif(is_true, elements, idx)
          return idx
        if keyword == KEYWORD_FOR:
          for_varname, for_list_name = self.__get_for_values(condition)
          for_list = self.resolve_value(None, for_list_name)
          for_blocks = self.__get_blocks_until_endfor(elements, idx)
          out = []
          row = 1
          for varvalue in for_list:
            self.default_params[for_varname] = varvalue
            if isinstance(varvalue, DotDict) or isinstance(varvalue, dict):
              varvalue['row'] = row
            idx2 = 0
            for_blocks_copy = []
            for block in for_blocks:
              for_blocks_copy.append(block.clone(True))
            while idx2 < len(for_blocks_copy):
              idx2 = self.__process_paragraph(for_blocks_copy, idx2)
            out = out + for_blocks_copy
            row += 1
          elements[idx:idx] = out
          idx += len(out)
          self.default_params[for_varname] = {}
          return idx
        # -- elimina la directiva por defecto
        del elements[idx]
        return idx
      # -- resuelve el párrafo
      _, somedollar, _ = self.__resolve_text_value(None, element)
      # -- si está vacío (sin runs ni textos), pero hubo evaluaciones dinámicas, lo elimina
      if somedollar and self.__is_paragraph_empty(element):
        del elements[idx]
        return idx
      # -- si no está vacío pasa al siguiente elemento
      return idx + 1
    if tagname == WT.TAG_TBL:
      self.__process_table(element)
      return idx + 1
    self.__resolve_text_value(None, element)
    return idx + 1

  def __get_for_values(self, param: str) -> tuple[str, str]:
    idx = param.find(' in ')
    if idx < 0:
      raise RuntimeError("List value expected in for ... in expression")
    value_name = param[0:idx].strip()
    list_name = param[idx + 4:].strip()
    return value_name, list_name

  def __process_table(self, tag: XmlTag):
    tagname = tag.name
    if tagname != WT.TAG_TABLE_CELL:
      for tag2 in tag.elements:
        self.__process_table(tag2)
      return
    self.process_paragraphs(tag)

  def __is_paragraph_empty(self, para_tag:XmlTag):
    has_text = False
    tag_runs = para_tag.get_tags(WT.TAG_R)
    if not tag_runs or len(tag_runs) == 0:
      return False
    for tag_run in tag_runs:
      tag_texts = tag_run.get_tags(WT.TAG_T)
      if not tag_texts or len(tag_texts) == 0:
        return False
      has_text = True
      for tag_text in tag_texts:
        text = tag_text.get_text()
        if text.strip() != '':
          return False
    return has_text

  def __get_blocks_until_endif(self, is_true: bool, elements: list, idx: int) -> list:
    if_blocks = []
    status = ProcessStatus()
    status.search = KEYWORD_IF
    # -- elimina la directiva condicional
    tag = elements[idx]
    self.__process_directive_with_status(tag, status)
    del elements[idx]
    while idx < len(elements):
      tag = elements[idx]
      if not isinstance(tag, XmlTag):
        del elements[idx]
        continue
      if self.__process_directive_with_status(tag, status):
        # -- elimina la directiva final
        del elements[idx]
        break
      if not status.else_mode:
        if is_true:
          if_blocks.append(tag)
          idx += 1
        else:
          del elements[idx]
      else:
        if not is_true:
          if_blocks.append(tag)
          idx += 1
        else:
          del elements[idx]

    return if_blocks
  
  def __get_blocks_until_endfor(self, elements: list, idx: int) -> list:
    for_blocks = []
    status = ProcessStatus()
    status.search = KEYWORD_FOR
    # -- elimina la directiva condicional
    tag = elements[idx]
    self.__process_directive_with_status(tag, status)
    del elements[idx]
    while idx < len(elements):
      tag = elements[idx]
      if not isinstance(tag, XmlTag):
        del elements[idx]
        continue
      if self.__process_directive_with_status(tag, status):
        # -- elimina la directiva final
        del elements[idx]
        break
      tag = tag.clone(True)
      for_blocks.append(tag)
      # -- elimina el tag del bloque for
      del elements[idx]
    return for_blocks

  def __process_directive_with_status(self, tag: XmlTag, status: ProcessStatus) -> bool:
    keyword, _ = self.__get_directive(tag)
    if keyword is not None:
      return status.process_directive_keyword(keyword)
    return False

  def __get_directive(self, tag: XmlTag) -> tuple[str, str]:
    if not isinstance(tag, XmlTag):
      return None, None
    tagname = tag.name
    if tagname != WT.TAG_T:
      for tag2 in tag.elements:
        keyword, arg = self.__get_directive(tag2)
        if keyword is not None:
          return keyword, arg
      return None, None
    text = tag.get_text()
    idx = text.find('{%')
    if idx < 0:
      return None, None
    idx += 2
    idx2 = text.find('%}', idx)
    if idx2 < 0:
      return None, None
    directive = text[idx:idx2].strip()
    idx1 = directive.find(' ')
    keyword = directive[0:idx1].strip() if idx1 >= 0 else directive
    arg = directive[idx1 + 1:].strip() if idx1 >= 0 else ''
    return keyword, arg

  @staticmethod
  def __reassign_ids(block: XmlTag, maxid: int) -> int:
    if not isinstance(block, XmlTag):
      return maxid
    elements = block.elements
    for item in elements:
      maxid = WordDocument.__reassign_ids(item, maxid)
    attrs = block.attrs
    if not attrs or len(attrs) == 0:
      return maxid
    attrname = 'o:spid'
    idvalue = attrs.get(attrname)
    if idvalue:
      del attrs[attrname]
    attrname = 'id'
    idvalue = attrs.get(attrname)
    if not idvalue:
      return maxid
    try:
      maxid += 1
      int(idvalue)
      attrs[attrname] = maxid
    except ValueError:
      attrs[attrname] = '_id_n' + str(maxid)
    return maxid

  def __resolve_text_value(self, row: int | None, block: XmlTag) -> tuple[bool, bool, bool]:
    """
    Processes a tag and resolves its text when present.
    """
    if not isinstance(block, XmlTag):
      return False, False, False
    tag = block.name
    elements = block.elements
    if len(elements) == 0:
      return False, False, False;
    if tag != WT.TAG_T:
      sometext = False
      somedollar = False
      sometextdollar = False
      idx = 0
      while idx < len(elements):
        item = elements[idx]
        if not isinstance(item, XmlTag):
          idx += 1
          continue
        tag0 = item.name
        sometext0, somedollar0, sometextdollar0 = self.__resolve_text_value(row, item)
        # -- si un TR no tienen entre sus TC algo de texto resuelto con $, se borra
        if tag0 == WT.TAG_TR and not sometextdollar0 and somedollar0:
          del elements[idx]
          continue
          # -- si un TBL no tienen entre sus TC algo de texto resuelto con $, se borra
        if tag0 == WT.TAG_TBL and not sometextdollar0 and somedollar0:
          del elements[idx]
          continue
        sometext = sometext0 or sometext
        somedollar = somedollar0 or somedollar
        sometextdollar = sometextdollar0 or sometextdollar
        idx += 1
      return sometext, somedollar, sometextdollar
    text_node = elements[0]
    if not isinstance(text_node, XmlText):
      return False, False, False
    value = text_node.content
    if not isinstance(value, str):
      return False, False, False
    idx1 = value.find('{{')
    if idx1 < 0:
      return text_util.trim(value) != '', False, False
    text1 = ''
    text2 = ''
    idx0 = 0
    while idx1 >= 0:
      idx2 = value.find('}}', idx1)
      if idx2 <= idx1:
        break
      text1 += value[idx0:idx1]
      text2 += value[idx0:idx1]
      param = value[idx1+2:idx2]
      resolved = self.resolve_value(row, param)
      if resolved is not None:
        resolved = str(resolved)
        text1 += resolved
      idx0 = idx2 + 2
      idx1 = value.find('{{', idx0)
    text1 = text1 + value[idx0:]
    text2 = text2 + value[idx0:]
    sometext = text1 != text2
    text_node.content = text1
    return sometext, True, sometext

  def __process_style_directive(self, param: str) -> bool:
    idx = param.find('=')
    if idx < 0:
      return False
    name = param[0:idx].strip()
    value = text.remove_quotes(param[idx + 1:].strip())
    self.styles.setStyle(name, value)
    return True

  @staticmethod
  def __get_schema_level(block: XmlTag) -> tuple[int, str]:
    if not isinstance(block, XmlTag):
      return -1, None
    tag = block.name
    elements = block.elements
    level = -1
    style = None
    if tag != WT.TAG_P_STYLE:
      for item in elements:
        level0, style0 = WordDocument.__get_schema_level(item)
        level = max(level0, level)
        if not style and style0:
          style = style0
      return level, style
    attrs = block.attrs
    if attrs:
      style = attrs.get(WT.ATTR_VAL)
      if style:
        try:
          level = int(style[len(style)-1])
        except ValueError:
          level = 0
    return level, style

  @staticmethod
  def __find_block(tagname: str, block: XmlTag):
    if not isinstance(block, XmlTag):
      return None
    tag = block.name
    if tag == tagname:
      return block
    content = block.elements
    for item in content:
      found = WordDocument.__find_block(tagname, item)
      if found:
        return found
    return None

  def expand_content(self, word_tag: XmlTag):
    """
    Expands embedded HTML content inside paragraphs.

    Args:
      word_tag: Root tag (document/body or similar).
    """
    idx = 0
    word_tag_children = word_tag.elements
    while idx < len(word_tag_children):
      word_tag_child = word_tag_children[idx]
      if not isinstance(word_tag_child, XmlTag):
        idx += 1
        continue
      word_tag_name = word_tag_child.name
      # Recurre el procesado para asegura que se procesa algún paragraph dentro
      if word_tag_name != WT.TAG_P:
        self.expand_content(word_tag_child)
        idx += 1
        continue
      ptag = word_tag_child
      # Verifica que tiene algún run dentro de paragraphs
      r_tags = ptag.get_tags(WT.TAG_R)
      if r_tags is None:
        idx += 1
        continue
        # Verifica que tiene algún text dentro del run
      was_text = False
      for r_tag in r_tags:
        tcontent = r_tag.get_tag_text(WT.TAG_T, False)
        if tcontent:
          was_text = True
          break
      if not was_text:
        idx += 1
        continue
        # -- Si es un párrafo, con runs y textos, expande su contenido
      out = self.expand_paragraph(ptag)
      # borra el párrafo procesado
      del word_tag_children[idx]
      # incluye la salida en el documento final      
      for tag in out:
        word_tag_children.insert(idx, tag)
        idx += 1      

  def expand_paragraph(self, p_tag: XmlTag):
    """
    Expands a paragraph with embedded HTML content.

    Args:
      p_tag: Paragraph tag.

    Returns:
      List of resulting tags.
    """
    p_tag_children = p_tag.elements
    # recorrre los elementos del párrafo
    out = []
    out_p_tag = None
    idx2 = 0
    while idx2 < len(p_tag_children):
      p_tag_child = p_tag_children[idx2];
      p_tag_child_name = p_tag_child.name
      if p_tag_child_name != WT.TAG_R:
        idx2 += 1
        continue
      r_tag = p_tag_child
      # -- Obtiene el texto del Run
      tcontent = r_tag.get_tag_text(WT.TAG_T, False)
      # -- Si no tiene texto, se ignora
      if not tcontent:
        if out_p_tag is None:
          out_p_tag = self.__create_empty_paragraph(p_tag, None)
          out.append(out_p_tag)
        out_p_tag.add_tag(r_tag.clone(True))
        idx2 += 1
        continue
        # -- Si el texto no tiene XML, lo ignora, pero lo escapa
      if not bool(re.search(r'<\s*[a-zA-Z][^>]*>', tcontent)):
        if out_p_tag is None:
          out_p_tag = self.__create_empty_paragraph(p_tag, None)
          out.append(out_p_tag)
        out_r_tag = doc_elements.create_run(r_tag, tcontent, {}, self.relationships, self.types, self.styles)
        out_p_tag.add_tag(out_r_tag)
        idx2 += 1
        continue
      parser = XmlParser()
      tag_list = parser.parse_text(tcontent)
      out_p_tag = self.__expand_html_tags(out, out_p_tag, p_tag, r_tag, None, tag_list, {})
      idx2 += 1
    return out
  
  def __expand_html_tags(self, out: list, out_p_tag: XmlTag|None, p_tag: XmlTag, r_tag: XmlTag, root_tag: XmlTag|None, html_tags: list, runprops0: dict):
    runprops = self.__process_html_tag_styles(root_tag, runprops0) if root_tag else dict(runprops0)
    root_name = root_tag.name.lower() if root_tag else None
    if root_name == 'img':
      if out_p_tag is None:
        out_p_tag = self.__create_empty_paragraph(p_tag, runprops)
        out.append(out_p_tag)
      out_r_tag = doc_elements.create_run(r_tag, '', runprops, self.relationships, self.types, self.styles)
      out_p_tag.add_tag(out_r_tag)
      return out_p_tag
    first_node = True
    for html_tag in html_tags:
      if isinstance(html_tag, XmlText):
        tcontent = html_tag.content
        if tcontent.strip() == '' and first_node:
          first_node = False;
          continue
        if out_p_tag is None:
          out_p_tag = self.__create_empty_paragraph(p_tag, runprops)
          out.append(out_p_tag)
        out_r_tag = doc_elements.create_run(r_tag, tcontent, runprops, self.relationships, self.types, self.styles)
        out_p_tag.add_tag(out_r_tag)
        first_node = False;
        continue
      first_node = False;
      # asegura que sea un XmlTag
      if not isinstance(html_tag, XmlTag):
        continue
      tag_name = html_tag.name.lower()
      # procesa estilos y resto de propiedades html para runprops: {
      #   "style" -> Estilo de Word
      #   "align" -> Alineación de texto de Word
      #   "code" -> Es código fuente
      #   "underline" -> Está subrayado
      #   "bold" -> Está en negrita
      #   "strike" -> Está tachado
      #   "italic" -> Está en itálica
      #   "color" -> Color
      #   "link" -> URL de enlace del texto
      #   "image" -> Datos de la imagen
      #   "height" -> Altura del elemento
      #   "width" -> Anchura del elemento
      # }
      if tag_name == 'br':
        out_p_tag = None
        continue  
      if tag_name == 'p':
      # crea un nuevo párrafo
        out_p_tag = self.__expand_html_tags(out, None, p_tag, r_tag, html_tag, html_tag.elements, runprops)
        continue  
      if tag_name == 'pre':
        out_p_tag = self.__expand_html_tags(out, None, p_tag, r_tag, html_tag, html_tag.elements, runprops)
        continue  
      if tag_name == 'blockquote':
        out_p_tag = self.__expand_html_tags(out, None, p_tag, r_tag, html_tag, html_tag.elements, runprops)
        continue  
      if tag_name == 'ul' or tag_name == 'ol':
        indent = runprops.get('indent')
        if indent is None or indent < 0:
          indent = 0
        else:
          indent += 1
        runprops['indent'] = indent
        runprops2 = self.__process_html_tag_styles(html_tag, runprops)
        li_list = html_tag.elements
        num_id = None
        if tag_name == 'ol':
          stylename = runprops2.get('style');
          num_id = self.numbering.add_numbering_start(stylename)
          runprops2['num_id'] = num_id
        else:
          runprops2['num_id'] = None
        for li in li_list:
          if not isinstance(li, XmlTag):
            continue
          out_p_tag = self.__expand_html_tags(out, None, p_tag, r_tag, li, li.elements, runprops2)
          runprops2['num_id'] = None
        indent = runprops.get('indent')
        if indent is not None and indent >= 0:
          indent -= 1
          runprops['indent'] = indent if indent >= 0 else None
        runprops['num_id'] = None
        continue
      headers = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']
      try:
        hidx = headers.index(tag_name)
      except ValueError:
        hidx = -1
      if hidx >= 0:
        out_p_tag = self.__expand_html_tags(out, None, p_tag, r_tag, html_tag, html_tag.elements, runprops)
        continue
      if tag_name == 'table':
        table_props = doc_elements.get_html_table_properties_to_json(html_tag)
        self.num_tables += 1
        table_out = doc_elements.create_table(self.num_tables, table_props, self.styles)
        for tag in table_out:
          out_p_tag = None
          self.expand_content(tag)
          out.append(tag)
        out_p_tag = None
        continue  
      out_p_tag = self.__expand_html_tags(out, out_p_tag, p_tag, r_tag, html_tag, html_tag.elements, runprops)
    return out_p_tag


  def __create_empty_paragraph(self, p_tag: XmlTag, runprops: dict|None):
    out_p_tag = p_tag.clone(False)
    out_ppr = None
    tags = p_tag.elements
    for tag in tags:
      if tag.name == WT.TAG_PPR:
        out_ppr = tag.clone(True)                  
        out_p_tag.add_tag(out_ppr)
        break

    if runprops:
      style = runprops.get('style')
      align = runprops.get('align')
      num_id = runprops.get('num_id')
      if style or align:
        if not out_ppr:     
          out_ppr = out_p_tag.add_tag(XmlTag(WT.TAG_PPR))
        if style:
          WordDocument.__set_ppr_style(out_ppr, WT.TAG_P_STYLE, style)
          if num_id:
            num_tag = out_ppr.add_tag('w:numPr')
            lvl_tag = num_tag.add_tag('w:ilvl')
            lvl_tag.add_attr('w:val', '0')
            num_id_tag = num_tag.add_tag('w:numId')
            num_id_tag.add_attr('w:val', str(num_id))
        if align:
          WordDocument.__set_ppr_style(out_ppr, WT.TAG_ALIGN, align)

    return out_p_tag

  @staticmethod
  def __set_ppr_style(ppr_tag: XmlTag, tag_name: str, value: str):
    if value:
      tag = ppr_tag.get_tag(tag_name, False)
      if not tag:
        tag = ppr_tag.add_tag(XmlTag(tag_name))
      tag.set_attr(WT.ATTR_VAL, value)

  def __process_html_tag_styles(self, html_tag: XmlTag, runprops: dict) -> dict:
    style = runprops.get('style')
    align = runprops.get('align')
    indent = runprops.get('indent')
    num_id = runprops.get('num_id')
    attrs = html_tag.attrs
    tag_name = html_tag.name.lower()
    classname = attrs.get('class')
    if classname:
      # -- mapea las classes HTML a estilos WORD
      classes = classname.split(' ')
      for classname in classes:
        if classname in ['left', 'right', 'center']:
          align = classname
        elif classname == 'justify':
          align = 'both'
        elif classname in Styles.CFG_SIMPLE_STYLES:
          style = self.styles.style_map.get(classname)
        elif classname.startswith('style_'):
          style = self.styles.style_map.get(classname[6:])
        else:
          for sname in Styles.CFG_LIST_STYLES:
            if classname == sname:
              stylelist = self.styles.style_map.get(classname)
              if stylelist:
                if indent is None or indent <= 0:
                  indent = 0
                elif indent >= len(style_list):
                  indent = len(style_list) - 1
                style = style_list[indent]
              break
            if classname.startswith(sname):
              off = len(sname)
              num = types.to_int(classname[off:])
              stylename = classname[:off]
              if num <= 0:
                num = 1
              elif num > 6:
                num = 6
              stylelist = self.styles.style_map.get(stylename)
              if isinstance(stylelist, list) and num >= 1 and num <= 6:
                style = stylelist[num - 1]
              break

    if style is None:
      if tag_name == 'p':
        style = self.styles.style_map.get(Styles.CFG_STYLE_PARAGRAPH)
      if tag_name == 'pre':
        style = self.styles.style_map.get(Styles.CFG_STYLE_CODEBLOCK)
      if tag_name == 'blockquote':
        style = self.styles.style_map.get(Styles.CFG_STYLE_QUOTE)
      if tag_name == 'code':
        style = self.styles.style_map.get(Styles.CFG_STYLE_CODE)
      if tag_name == 'ul':
        style_list = self.styles.style_map.get(Styles.CFG_STYLE_LIST_BULLET)
        if indent is None or indent <= 0:
          indent = 0
        elif indent >= len(style_list):
          indent = len(style_list) - 1
        style = style_list[indent]
      if tag_name == 'ol':
        style_list = self.styles.style_map.get(Styles.CFG_STYLE_LIST_NUMBER)
        if indent is None or indent <= 0:
          indent = 0
        elif indent >= len(style_list):
          indent = len(style_list) - 1
        style = style_list[indent]

      headers = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']
      try:
        hidx = headers.index(tag_name)
      except ValueError:
        hidx = -1
      if hidx >= 0:
        heading_list = self.styles.style_map.get(Styles.CFG_STYLE_HEADING)
        if heading_list and hidx >= len(heading_list):
          hidx = len(heading_list) - 1
        style = heading_list[hidx]

    if align not in ['left', 'right', 'center', 'both']:
      align = 'left'

    out_runprops = {
      'style': style,
      'align': align,
      'indent': indent,
      'num_id': num_id
    }
    if tag_name == 'b' or tag_name == 'strong':
      out_runprops['bold'] = True
    elif tag_name == 'i' or tag_name == 'em':
      out_runprops['italic'] = True
    elif tag_name == 'strike' or tag_name == 'stroke' or tag_name == 's':
      out_runprops['strike'] = True
    elif tag_name == 'u':
      out_runprops['underline'] = True
    elif tag_name == 'code':
      out_runprops['code'] = True
    elif tag_name == 'font':
      color = attrs.get('color')
      if color:
        out_runprops['color'] = color
    elif tag_name == 'a':
      url = attrs.get('href')
      if url:
        out_runprops['link'] = url
    elif tag_name == 'img':
      src = attrs.get('src')
      if src:
        out_runprops['image'] = src

    istyle = attrs.get('style')
    if istyle:
      doc_elements.get_css_properties(istyle, out_runprops)
    return out_runprops

  def __process_chart_excel_file(self, chart_name, tempdir, target):
    excel = ExcelDocument(target)
    try:
      if self.test_mode:
        excel.test_with_json({})
      else:
        excel.create_doc_with_resolvers(target, self.value_resolver, self.repeating_resolver)
        self.clear_chart_values(chart_name, tempdir, excel)
    finally:
      excel.close()
    self.access_ok_param_list = types.merge_list_unique(self.access_ok_param_list, excel.access_ok_param_list)
    self.access_err_param_list = types.merge_list_unique(self.access_err_param_list, excel.access_err_param_list)

  def search_graphic_frames(self, tempdir, elements0):
    """
    Finds and processes graphic frames in the document.

    Args:
      tempdir: Temporary directory.
      elements0: List of XML elements.
    """
    idx0 = 0
    while idx0 < len(elements0):
      element0 = elements0[idx0]
      if isinstance(element0, XmlTag):
        if element0.name == 'w:drawing':
          self.__process_graphic_frame(tempdir, element0)
        else:
          self.search_graphic_frames(tempdir, element0.elements)
      idx0 += 1

  def __process_graphic_frame(self, tempdir, element0):
    cnvpr = element0.get_tag_path(['wp:*', 'wp:docPr'])
    chart = element0.get_tag_path(['wp:*', 'a:graphic', 'a:graphicData', 'c:chart'], False)
    if cnvpr is None or chart is None:
      return
    descr = cnvpr.get_attr('descr')
    if descr is None:
      return
      # -- obtiene los datos de la inyección
    value = self.__resolve_descr(descr)
    if not value.startswith('{') or self.test_mode:
      return
    cnvpr.set_attr('descr', '')

    rid = chart.get_attr('r:id')
    if rid is None:
      return
    rel = self.relationships.get_relationship_by_id(rid)
    if rel is None:
      return
    target = rel.target
    tfile = file_util.get_filename(target)
    chart_rel_file = tempdir + '/' + WORD_CHARTS_RELS + tfile + '.rels'
    relationships = Relationships(tempdir, chart_rel_file)
    relations = relationships.get_relationships(None, ".xlsx")
    if len(relations) == 0:
      return
    relation = relations[0]
    # -- obtiene nombre del archivo excel
    xml_file = relation.target
    # -- carga el json
    diagram = json.loads(value)
    # Crear datos en las celdas A1:C5
    datos = [
      diagram.get('legends')
    ]
    cats = diagram.get('categories')
    series = diagram.get('series')
    for idx, cat in enumerate(cats):
      row = [cat]
      for serie in series:
        data = serie.get('data')
        value = data[idx] if idx < len(data) else 0
        row.append(value)
      datos.append(row)

    if file_util.exist(xml_file):
      file_util.remove_file(xml_file)
    path = file_util.get_pathname(sys.modules['catslap.xlsx'].__file__)
    path = path + 'empty.xlsx'
    shutil.copy2(path, xml_file)
    # -- abre archivo excel para escritura
    excel = ExcelDocument(xml_file, False)
    try:
      excel.write_cells(datos)
      dbytes = excel.get_document_bytes()
      file_util.write_bytes(xml_file, dbytes)
      self.clear_chart_values(tfile, tempdir, excel)
    finally:
      excel.close()
      # -- modifica charts
    chart_file = tempdir + '/' + WORD_CHARTS + tfile
    parser = XmlParser()
    root_tag = parser.parse_file(chart_file)
    title_tag = root_tag.get_tag_path(['c:chart', 'c:title', 'c:tx', 'c:rich', 'a:p', 'a:r', 'a:t'], False)
    if title_tag:
      title_tag.set_text(diagram.get('title'))
    chart_tag = root_tag.get_tag_path(['c:chart', 'c:plotArea', '*Chart'], False)
    if chart_tag is not None:
      ser_tags = chart_tag.get_tags('c:ser')
      for idx, ser_tag in enumerate(ser_tags):
        if idx >= len(series):
          chart_tag.remove(ser_tag)
          continue
        v_tag = ser_tag.get_tag_path(['c:tx', 'c:strRef', 'c:strCache', 'c:pt', 'c:v'], False)
        if v_tag:
          name = series[idx].get('name')
          if name:
            v_tag.set_text(name)
    parser.write_file()

  def __resolve_descr(self, descr):
    descr = XmlParser.resolve_entities(descr)
    descr = text_util.trim(descr)
    if not descr.startswith("{{") or not descr.endswith("}}"):
      return ''
    param = descr[2:len(descr) - 2]
    value = self.resolve_value([], param)
    if text_util.is_empty(value):
      return ''
    if isinstance(value, dict):
      value = json.dumps(value)
    elif isinstance(value, str) and value.startswith('{'):
      value = eval(value)
    return value

  def process_descr_attrs(self, elements0):
    """
    Processes shape 'descr' attributes and replaces content.

    Args:
      elements0: List of XML elements.
    """
    idx0 = 0
    while idx0 < len(elements0):
      element0 = elements0[idx0]
      if not isinstance(element0, XmlTag):
        idx0 += 1
        continue
      elements = element0.elements
      self.process_descr_attrs(elements)
      for element in elements:
        if not isinstance(element, XmlTag):
          continue
        if element.name != 'p:cNvPr':
          continue
        descr = element.get_attr('descr')
        if descr is None:
          continue
        value = self.__resolve_descr(descr)
        if self.test_mode:
          continue
        element.set_attr('descr', '')
        if isinstance(value, dict):
          shape = value
          if shape is None:
            continue
          shape = dict(shape)
          x = shape.get('x')
          y = shape.get('y')
          wd = shape.get('wd')
          hg = shape.get('hg')
          bg = shape.get('bg')
          fg = shape.get('fg')
          txt = shape.get('text')
        else:
          x = None
          y = None
          wd = None
          hg = None
          bg = None
          fg = None
          txt = str(value)
          # -- busca posición de objeto
        idx1 = 0
        sppr = None
        txbody = None
        while idx1 < len(elements0):
          element1 = elements0[idx1]
          if element1.name == 'p:spPr':
            sppr = element1
          if element1.name == 'p:txBody':
            txbody = element1
          idx1 += 1
          # -- posición, dimensión y color de fondo
        if sppr is not None:
        # -- busca posición de la forma
          xfrm = sppr.get_tag('a:xfrm', False)
          if xfrm is not None:
            off = xfrm.get_tag('a:off')
            ext = xfrm.get_tag('a:ext')
            if x is not None and text.is_decimal(x):
              x_value = int(round(PT_PER_CM * float(x)))
              off.set_attr('x', x_value)
            if y is not None and text.is_decimal(y):
              y_value = int(round(PT_PER_CM * float(y)))
              off.set_attr('y', y_value)
            if wd is not None and text.is_decimal(wd):
              wd_value = int(round(PT_PER_CM * float(wd)))
              ext.set_attr('cx', wd_value)
            if hg is not None and text.is_decimal(hg):
              hg_value = int(round(PT_PER_CM * float(hg)))
              ext.set_attr('cy', hg_value)
              # -- busca color de la forma
          solid_fill = sppr.get_tag(WT.TAG_SOLID_FILL, False)
          if solid_fill is None:
            solid_fill = sppr.get_tag(WT.TAG_NO_FILL, False)
            if solid_fill is not None:
              solid_fill.name = WT.TAG_SOLID_FILL
          if solid_fill is not None and bg is not None and len(bg) == 6 and text.is_hex(bg):
            solid_fill.clear_tags()
            solid_fill.add_tag(WT.TAG_SRGB_CLR, {'val': bg})
            # -- color de texto
        if txbody is not None and fg is not None and len(fg) == 6 and text.is_hex(fg):
          def __change_solid_fill_tag_color(tag, rgb) -> bool:
            if tag.name == WT.TAG_NO_FILL:
              tag.name = WT.TAG_SOLID_FILL
            if tag.name == WT.TAG_SOLID_FILL:
              tag.clear_tags()
              tag.add_tag(WT.TAG_SRGB_CLR, {'val': rgb})
              return True
            for elem in tag.elements:
              if __change_solid_fill_tag_color(elem, rgb):
                return True
            return False
          __change_solid_fill_tag_color(txbody, fg)
          # -- sustitucion de texto
        if txbody is not None and txt is not None:
          tag_p = txbody.get_tag('a:p', False)
          tag_r = tag_p.get_tag('a:r', False) if tag_p is not None else None
          if tag_r is None:
            tag_endrpr = tag_p.get_tag('a:endParaRPr', False)
            if tag_endrpr is not None:
              tag_p.remove_tag('a:endParaRPr')
              tag_r = tag_p.add_tag('a:r')
              tag_endrpr.name = 'a:rPr'
              tag_r.add_element(tag_endrpr)
            else:
              tag_r = tag_p.add_tag('a:r')
          tag_t = tag_r.get_tag('a:t', False)
          if tag_t is None:
            tag_t = tag_r.add_tag('a:t')
          txt = text_util.trim(txt)
          tag_t.set_text(XmlParser.escape_entities(txt))
      idx0 += 1

  def clear_chart_values(self, chart_name, tempdir, excel):
    """
    Updates and clears chart series caches.

    Args:
      chart_name: Chart name.
      tempdir: Temporary directory.
      excel: Auxiliary Excel document.
    """
    if self.test_mode:
      return
    parser = XmlParser()
    chart_file = tempdir + '/' + WORD_CHARTS + chart_name
    tag = parser.parse_file(chart_file, 'c:chartSpace')
    self.collapse_paragraphs(tag.elements)
    # -- coge el chart desde el path XML
    tag = tag.get_tag_path(['c:chart', 'c:plotArea', '*Chart'])
    # -- coge las series de datos del chart
    sers = tag.get_tags('c:ser')
    for ser in sers:
      cat = ser.get_tag('c:cat')
      ref_tag = cat.get_tag('c:strRef', False)
      tag_name = 'c:strCache'
      if ref_tag is None:
        ref_tag = cat.get_tag('c:numRef', True)
        tag_name = 'c:numCache'
      f_tag = ref_tag.get_tag('c:f')
      sheet_ref = f_tag.get_text()
      str_cache = ref_tag.get_tag(tag_name)
      # -- Elimina la cache de texto de campos para rehazerla (podría no venir)
      str_cache.clear_tags()
      # -- extrae los datos desde la referencia de excel (Ej: Hoja1!$A$2:$A$3)
      sheet_name, sdata, edata = parse_data_ref(sheet_ref)
      scell_num = get_cell_num(sdata)
      srow_num = get_row_num(sdata)
      ecell_num = get_cell_num(edata)
      sdata = get_cell_format_position(scell_num, srow_num, True)
      sheet_ref = sheet_name + "!" + sdata
      values = excel.extract_data(sheet_ref)
      edata = get_cell_format_position(ecell_num, srow_num + len(values) - 1, True)
      sdata = get_cell_format_position(scell_num, srow_num, True)
      sheet_ref = sheet_name + "!" + sdata + ":" + edata
      f_tag.set_text(sheet_ref)
      str_cache.add_tag('c:ptCount', {'val': len(values)})
      for idx in range(0, len(values)):
        value_idx = values[idx]
        if value_idx and len(value_idx) > 0:
          str_cache.add_tag('c:pt', {'idx': idx}).add_tag_text('c:v', value_idx[0])
      # -- Introduce los valores
      vals = ser.get_tags('c:val')
      for val in vals:
        ref_tag = val.get_tag('c:numRef', False)
        tag_name = 'c:numCache'
        if ref_tag is None:
          ref_tag = val.get_tag('c:strRef', True)
          tag_name = 'c:strCache'
        f_tag = ref_tag.get_tag('c:f')
        sheet_ref = f_tag.get_text()
        sheet_name, sdata, edata = parse_data_ref(sheet_ref)
        scell_num = get_cell_num(sdata)
        srow_num = get_row_num(sdata)
        ecell_num = get_cell_num(edata)
        sdata = get_cell_format_position(scell_num, srow_num, True)
        sheet_ref = sheet_name + "!" + sdata
        values = excel.extract_data(sheet_ref)
        edata = get_cell_format_position(ecell_num, srow_num + len(values) - 1, True)
        sdata = get_cell_format_position(scell_num, srow_num, True)
        sheet_ref = sheet_name + "!" + sdata + ":" + edata
        f_tag.set_text(sheet_ref)

        num_cache = ref_tag.get_tag(tag_name)
        num_cache.clear_tags()
        # -- extrae los datos desde la referencia de excel
        values = excel.extract_data(sheet_ref)
        num_cache.add_tag_text('c:formatCode', 'General')
        num_cache.add_tag('c:ptCount', {'val': len(values)})
        for idx in range(0, len(values)):
          value_idx = values[idx]
          if value_idx and len(value_idx) > 0:
            num_cache.add_tag('c:pt', {'idx': idx}).add_tag_text('c:v', value_idx[0])
    parser.write_file()

  def get_bytes_with_json(self, json: dict) -> bytes:
    """
    Generates document bytes using JSON.

    Args:
      json: Value map.

    Returns:
      Bytes of the generated document.
    """
    value_resolver = dict_value_resolver(json)
    repeating_resolver = dict_repeat_resolver(json)
    reindex = self.config_params.get(CFG_PARAM_REINDEX)

    if types.to_bool(reindex) == False:
      return self.get_bytes_with_resolvers(value_resolver, repeating_resolver)

    format_type = self.config_params.get(CFG_PARAM_OUTPUT_FORMAT_TYPE)
    extension = 'pdf' if format_type == 'PDF' else 'docx'
    bytes = self.get_bytes_with_resolvers(value_resolver, repeating_resolver)
    try:
      input_file = file_util.get_temp_file(None, '_in.docx')
      file_util.write_bytes(input_file, bytes)
      output_file = file_util.get_temp_file(None, '_out.' + extension)
      # -- actualiza la tabla de contenidos
      print(str(['--headless', '--norestore', '--invisible', 'Macro.odt', f'macro://Macro/Standard.Module1.UpdateTOCAndExportToPDF({input_file},{output_file})']))
      process = soffice.execute(['--headless', '--norestore', '--invisible', 'Macro.odt', f'macro://Macro/Standard.Module1.UpdateTOCAndExportToPDF({input_file},{output_file})'])
      print("Office process launched")
      init_time = datetime.now()
      completed_file = output_file + '.completed'
      while (init_time + timedelta(minutes=2)) > datetime.now() and not file_util.exist(completed_file):
        time.sleep(5)
      if not file_util.exist(output_file):
        raise Exception('No es posible generar el documento')
      bytes3 = file_util.read_bytes(output_file)
      if len(bytes3) == 0:
        raise Exception('No es posible generar el documento (documento vacío)')
      process.terminate()
      return bytes3
    except Exception as e:
      print(e)      
    finally:
      try: 
        os.remove(completed_file)
      except:
        pass
      os.remove(output_file)
      os.remove(input_file)
