# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)
import sys
import json
import shutil

from catslap.base.document import Document
from catslap.base.relationships import Relationships
from catslap.base.types import ContentTypes
from catslap.pptx import elements as doc_elements
from catslap.xlsx.document import ExcelDocument, parse_data_ref, get_cell_format_position, get_cell_num, get_row_num
from catslap.utils import encoding as enc_util
from catslap.utils import file as file_util
from catslap.utils import html
from catslap.utils import text as text_util
from catslap.utils import types
from catslap.utils.xml import XmlParser, XmlTag, XmlText, CONFIG_PARAM_INCLUDE_DECL, CONFIG_PARAM_HTML

# -- puntos de powerpoint por cm
PT_PER_CM = 360000

PPT_SLIDES = "ppt/slides/"
PPT_DIAGRAMS = "ppt/diagrams/"
PPT_EMBEDINGS = "ppt/embeddings/"
PPT_CHARTS = "ppt/charts/"
PPT_CHARTS_RELS = PPT_CHARTS + "_rels/"
DOT_RELS = '.rels'

WORD_DOCUMENT_RELS = "word/_rels/document.xml.rels"
PPT_DOCUMENT_TYPES = "[Content_Types].xml"

TAG_SLD = 'p:sld'
TAG_DATAMODEL = 'dgm:dataModel'
TAG_TBL = 'a:tbl'
TAG_TR = 'a:tr'
TAG_P = 'a:p'
TAG_R = 'a:r'
TAG_RPR = 'a:rPr'
TAG_T = 'a:t'
TAG_SOLID_FILL = 'a:solidFill'
TAG_NO_FILL = 'a:noFill'
TAG_SRGB_CLR = 'a:srgbClr'


IGNORABLE_TAGS = ['c:lang']
IGNORABLE_EMPTY_TAGS = ['a:effectLst']
IGNORABLE_EMPTY_STYLE_TAGS = ['a:ea', 'a:cs', TAG_NO_FILL, 'a:ln', 'a:uLnTx', 'a:uFillTx', 'a:latin']
IGNORABLE_ATTRS = ['dirty', 'err']


class PowerPointDocument(Document):
  """
  Processes and generates PowerPoint presentations (.pptx).

  Attributes:
    max_id: Maximum ID used in the document.
    types: Package ContentTypes.
    relationships: Relationships for the current slide.
  """
  def __init__(self, file: str):
    super().__init__(file)
    self.max_id = 0
    self.types = None
    self.relationships = None

  def process_template(self, tempdir: str):
    """
    Processes the PPTX template in the temporary directory.

    Args:
      tempdir: Temporary directory with ZIP contents.
    """
    ppt_types = tempdir + "/" + PPT_DOCUMENT_TYPES
    self.types = ContentTypes(ppt_types)
    self.__process_ppt_slides(tempdir, TAG_SLD)
    if not self.test_mode:
      self.types.write_file()

  def __process_ppt_slides(self, tempdir: str, tag_name: str):
    xml = XmlParser()
    # -- parsea los slides de la PPT
    idx = 1
    found = True
    while found:
      slide_file = tempdir + '/' + PPT_SLIDES + "slide" + str(idx) + ".xml"
      found = file_util.exist(slide_file)
      if found:
        file0_rel = tempdir + '/' + PPT_SLIDES + "_rels/slide" + str(idx) + ".xml.rels"
        self.relationships = Relationships(tempdir, file0_rel)
        document = xml.parse_file(slide_file, tag_name)
        blocks = document.elements
        self.collapse_paragraphs(blocks)
        self.max_id = max(self.max_id, xml.max_id)
        self.search_graphic_frames(tempdir, blocks)
        self.process_descr_attrs(blocks)
        self.process_paragraphs(blocks)
        self.process_html_content(blocks)
        xml_content = xml.get_pretty_xml(document, {CONFIG_PARAM_INCLUDE_DECL: True})
        file_util.write_bytes(slide_file, bytes(xml_content, enc_util.UTF_8))
        if not self.test_mode:
          self.relationships.write_file()
        idx += 1
    idx = 1
    found = True
    while found:
      diagram_file = tempdir + '/' + PPT_DIAGRAMS + "data" + str(idx) + ".xml"
      found = file_util.exist(diagram_file)
      if found:
        self.relationships = None
        document = xml.parse_file(diagram_file, TAG_DATAMODEL)
        blocks = document.elements
        self.collapse_paragraphs(blocks)
        self.max_id = max(self.max_id, xml.max_id)
        self.process_paragraphs(blocks)
        self.process_html_content(blocks)
        xml_content = xml.get_pretty_xml(document, {CONFIG_PARAM_INCLUDE_DECL: True})
        file_util.write_bytes(diagram_file, bytes(xml_content, enc_util.UTF_8))
        idx += 1

    files = file_util.list_files(tempdir + '/' + PPT_CHARTS_RELS)
    if files:
      for file in files:
        chart_file = tempdir + '/' + PPT_CHARTS_RELS + file
        if file.endswith('.xml.rels'):
          chart_name = file[0:len(file) - 5]
          relationships = Relationships(tempdir, chart_file)
          relations = relationships.get_relationships(None, ".xlsx")
          if len(relations) == 0:
            continue
          relation = relations[0]
          target = relation.target
          self.__process_chart_excel_file(chart_name, tempdir, target)

  def __process_chart_excel_file(self, chart_name, tempdir, target):
    excel = ExcelDocument(target, False)
    try:
      if self.test_mode:
        excel.test_with_json({})
      else:
        excel.create_doc_with_resolvers(target, self.value_resolver, self.repeating_resolver)
        chart_file = tempdir + '/' + PPT_CHARTS + chart_name
        self.clear_chart_values(chart_file, excel)
    finally:
      excel.close()
    self.access_ok_param_list = types.merge_list_unique(self.access_ok_param_list, excel.access_ok_param_list)
    self.access_err_param_list = types.merge_list_unique(self.access_err_param_list, excel.access_err_param_list)

  def clear_chart_values(self, chart_file, excel):
    """
    Updates and clears chart series caches.

    Args:
      chart_file: Chart file.
      excel: Auxiliary Excel document.
    """
    if self.test_mode:
      return
    parser = XmlParser()
    tag = parser.parse_file(chart_file, 'c:chartSpace')
    self.collapse_paragraphs(tag.elements)
    tag = tag.get_tag_path(['c:chart', 'c:plotArea', '*Chart'])
    sers = tag.get_tags('c:ser')
    for ser in sers:
      cat = ser.get_tag('c:cat')
      ref_tag = cat.get_tag('c:strRef', False)
      tag_name = 'c:strCache'
      if ref_tag is None:
        ref_tag = cat.get_tag('c:numRef', False)
        tag_name = 'c:numCache'
      if ref_tag is not None:
        f_tag = ref_tag.get_tag('c:f')
        sheet_ref = f_tag.get_text()
        str_cache = ref_tag.get_tag(tag_name)
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
      vals = ser.get_tags('c:val')
      for val in vals:
        ref_tag = val.get_tag('c:numRef', False)
        tag_name = 'c:numCache'
        if ref_tag is None:
          ref_tag = val.get_tag('c:strRef', False)
          tag_name = 'c:strCache'
        if ref_tag is not None:
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
      for attr_name in IGNORABLE_ATTRS:
        tag.remove_attr(attr_name)
      tag_name = tag.name
      tags = tag.elements
      # -- ignora el tag
      if tag_name in IGNORABLE_TAGS or (len(tags) == 0 and tag_name in IGNORABLE_EMPTY_TAGS):
        del elements[idx]
        continue
        # -- funde todos los a:r que sean iguales en un solo a:r
      if tag_name == TAG_R:
        rpr = tag.get_tag(TAG_RPR, False)
        t = tag.get_tag(TAG_T, False)
        if len(tags) == 0 or t is None:
          idx += 1
          continue
        tag_name = t.name
        if tag_name == TAG_T:
          if last_t is not None and PowerPointDocument.__is_the_same_rpr(last_rpr, rpr):
            ct = t.elements[0].content if len(t.elements) > 0 and isinstance(t.elements[0], XmlText) else None
            if ct is not None:
              if len(last_t.elements) == 0:
                last_t.add_text(ct)
              else:
                last_t.elements[0].append(ct)
            last_t.attrs['xml:space'] = 'preserve'
            del elements[idx]
            continue
          last_rpr = rpr
          last_t = t
        else:
          last_rpr = None
          last_t = None
      idx += 1

  @staticmethod
  def __is_the_same_rpr(rpr1: XmlTag or None, rpr2: XmlTag or None) -> bool:
    if rpr1 is None and rpr2 is None:
      return True
    xml = XmlParser()
    rpr1 = rpr1.clone(True)
    rpr2 = rpr2.clone(True)
    PowerPointDocument.__remove_ignorable_style_tags(rpr1)
    PowerPointDocument.__remove_ignorable_style_tags(rpr2)
    dump1 = xml.get_outer_xml(rpr1)
    dump2 = xml.get_outer_xml(rpr2)
    return dump1 == dump2

  @staticmethod
  def __remove_ignorable_style_tags(tag: XmlTag):
    elements = tag.elements
    idx = 0
    while idx < len(elements):
      element = elements[idx]
      if not isinstance(element, XmlTag):
        idx += 1
        continue
      tag_name = element.name
      if tag_name in IGNORABLE_EMPTY_STYLE_TAGS:
        del elements[idx]
        continue
      PowerPointDocument.__remove_ignorable_style_tags(element)
      idx += 1

  def __resolve_descr(self, descr):
    descr = XmlParser.resolve_entities(descr)
    descr = text_util.trim(descr)
    if not descr.startswith("{{") or not descr.endswith("}}"):
      return ''
    param = descr[2:len(descr) - 2]
    value = self.resolve_value(None, param)
    if text_util.is_empty(value):
      return ''
    if isinstance(value, dict):
      value = json.dumps(value)
    elif isinstance(value, str) and value.startswith('{'):
      value = eval(value)
    return value

  def search_graphic_frames(self, tempdir, elements0):
    """
    Finds and processes graphic frames in slides.

    Args:
      tempdir: Temporary directory.
      elements0: List of XML elements.
    """
    idx0 = 0
    while idx0 < len(elements0):
      element0 = elements0[idx0]
      if isinstance(element0, XmlTag):
        if element0.name == 'p:graphicFrame':
          self.__process_graphic_frame(tempdir, element0)
        else:
          self.search_graphic_frames(tempdir, element0.elements)
      idx0 += 1

  def __process_graphic_frame(self, tempdir, element0):
    cnvpr = element0.get_tag_path(['p:nvGraphicFramePr', 'p:cNvPr'])
    chart = element0.get_tag_path(['a:graphic', 'a:graphicData', 'c:chart'], False)
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
    chart_rel_file = tempdir + '/' + PPT_CHARTS_RELS + tfile + DOT_RELS
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

    path = file_util.get_pathname(sys.modules['ooxml2.xlsx'].__file__)
    path = path + 'empty.xlsx'
    shutil.copy2(path, xml_file)
    # -- abre archivo excel para escritura
    excel = ExcelDocument(xml_file, False)
    try:
      excel.write_cells(datos)
      dbytes = excel.get_document_bytes()
      file_util.write_bytes(xml_file, dbytes)
      chart_file = tempdir + '/' + PPT_CHARTS + tfile
      self.clear_chart_values(chart_file, excel)
    finally:
      excel.close()
      # -- modifica charts
    chart_file = tempdir + '/' + PPT_CHARTS + tfile
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
            if x is not None and text_util.is_decimal(x):
              x_value = int(round(PT_PER_CM * float(x)))
              off.set_attr('x', x_value)
            if y is not None and text_util.is_decimal(y):
              y_value = int(round(PT_PER_CM * float(y)))
              off.set_attr('y', y_value)
            if wd is not None and text_util.is_decimal(wd):
              wd_value = int(round(PT_PER_CM * float(wd)))
              ext.set_attr('cx', wd_value)
            if hg is not None and text_util.is_decimal(hg):
              hg_value = int(round(PT_PER_CM * float(hg)))
              ext.set_attr('cy', hg_value)
              # -- busca color de la forma
          solid_fill = sppr.get_tag(TAG_SOLID_FILL, False)
          if solid_fill is None:
            solid_fill = sppr.get_tag(TAG_NO_FILL, False)
            if solid_fill is not None:
              solid_fill.name = TAG_SOLID_FILL
          if solid_fill is not None and bg is not None and len(bg) == 6 and text_util.is_hex(bg):
            solid_fill.clear_tags()
            solid_fill.add_tag(TAG_SRGB_CLR, {'val': bg})
            # -- color de texto
        if txbody is not None and fg is not None and len(fg) == 6 and text_util.is_hex(fg):
          def __change_solid_fill_tag_color(tag, rgb) -> bool:
            if tag.name == TAG_NO_FILL:
              tag.name = TAG_SOLID_FILL
            if tag.name == TAG_SOLID_FILL:
              tag.clear_tags()
              tag.add_tag(TAG_SRGB_CLR, {'val': rgb})
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

  def process_paragraphs(self, elements: list):
    """
    Processes paragraphs resolving placeholders and repetitions.

    Args:
      elements: List of XML elements.
    """
    idx = 0
    while idx < len(elements):
      element = elements[idx]
      if not isinstance(element, XmlTag):
        idx += 1
        continue
      tag_name = element.name
      # -- bloques de primer nivel (párrafos, tablas, etc.)
      # -- Caso 4: hay repeticiones dentro de una tabla (repeticiones de fila)
      if tag_name == TAG_TBL:
        self.__process_table(element)
        if len(element.elements) == 0:
          del elements[idx]
          continue
        idx += 1
        continue
      sometext0, somedollar0, _ = self.__resolve_text_value(None, elements[idx])
      if not sometext0 and somedollar0:
        del elements[idx]
        continue
      idx += 1

  def __process_table(self, tbl_tag: XmlTag):
    tr_tags = tbl_tag.elements
    pos = 0
    while pos < len(tr_tags):
      tr_tag = tr_tags[pos]
      # -- asegura que se procesan las filas de la tabla
      if not isinstance(tr_tag, XmlTag) or tr_tag.name != TAG_TR:
        pos += 1
        continue
        # -- comprueba si hay repeticiones de fila dentro de las celdas
      repeating = self.__resolve_text_repeating(tr_tag)
      if repeating > 1:
        self. __repeat_block_from(tr_tags, pos, repeating)
        continue
        # -- borra la fila si no hay texto o si hay pero no es de dollar habiendo dollar
      sometext, somedollar, sometextdollar = self.__resolve_text_value(None, tr_tags[pos])
      if not sometext or (somedollar and not sometextdollar):
        del tr_tags[pos]
        continue
      pos += 1

  def __repeat_block_from(self, tags: list, idx: int, repeating: int):
    """
    Repeats a tag from a tag list at a given position.
    """
    newblocks = []
    elem = tags[idx]
    for row in range(0, repeating):
      tr0 = elem.clone()
      sometext, somedollar, sometextdollar = self.__resolve_text_value(row, tr0)
      # si es un tr con $ en alguna celda pero sin texto de $, ignora toda la fila
      if not sometextdollar and somedollar:
        continue
      self.max_id = PowerPointDocument.__reassign_ids(tr0, self.max_id)
      newblocks.append(tr0)
      # -- elimina el bloque que sirvió de patrón
    del tags[idx]
    # -- inserta todos los nuevos bloques resueltos
    pos = idx
    for newblock in newblocks:
      tags.insert(pos, newblock)
      pos += 1

  @staticmethod
  def __reassign_ids(block: XmlTag, maxid: int) -> int:
    if not isinstance(block, XmlTag):
      return maxid
    elements = block.elements
    for item in elements:
      maxid = PowerPointDocument.__reassign_ids(item, maxid)
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

  def __resolve_text_value(self, row: int or None, block: XmlTag) -> (bool, bool, bool):
    """
    Processes a tag and resolves its text when present.
    """
    if not isinstance(block, XmlTag):
      return False, False, False
    tag = block.name
    elements = block.elements
    if tag != TAG_T:
      sometext = False
      somedollar = False
      sometextdollar = False
      idx = 0
      while idx < len(elements):
        item = elements[idx]
        if not isinstance(item, XmlTag):
          idx += 1
          continue
        sometext0, somedollar0, sometextdollar0 = self.__resolve_text_value(row, item)
        tag0 = item.name
        # -- si un TR no tienen entre sus TC algo de texto resuelto con $, se borra
        if tag0 == TAG_TR and not sometextdollar0 and somedollar0:
          del elements[idx]
          continue
          # -- si un TBL no tienen entre sus TC algo de texto resuelto con $, se borra
        if tag0 == TAG_TBL and not sometextdollar0 and somedollar0:
          del elements[idx]
          continue
        sometext = sometext0 or sometext
        somedollar = somedollar0 or somedollar
        sometextdollar = sometextdollar0 or sometextdollar
        idx += 1
      return sometext, somedollar, sometextdollar
    text_node = elements[0] if len(elements) > 0 else None
    if not isinstance(text_node, XmlText):
      return False, False, False
    value = text_node.content
    if not isinstance(value, str):
      return False, False, False
    idx1 = value.find('{{')
    if idx1 < 0:
      return text_util.trim(value) != '', False, False
    htext = ''
    text2 = ''
    idx0 = 0
    while idx1 >= 0:
      idx2 = value.find('}}', idx1)
      if idx2 <= idx1:
        break
      htext += value[idx0:idx1]
      text2 += value[idx0:idx1]
      param = value[idx1+2:idx2]
      resolved = self.resolve_value(row, param)
      if resolved is not None:
        htext = htext + str(resolved)
      idx0 = idx2 + 2
      idx1 = value.find('{{', idx0)
    htext = text_util.trim(htext + value[idx0:])
    text2 = text_util.trim(text2 + value[idx0:])
    sometext = htext != text2
    if sometext:
      text1 = PowerPointDocument.normalize_html_text(htext)
      text_node.content = text1
    else:
      text_node.content = ''
    return sometext, True, sometext

  def __resolve_text_repeating(self, block: XmlTag) -> int:
    if not isinstance(block, XmlTag):
      return 0
    tag = block.name
    elements = block.elements
    if tag != TAG_T:
      repeating = 0
      for item in elements:
        repeating = max(self.__resolve_text_repeating(item), repeating)
      return repeating
    text_node = elements[0]
    if not isinstance(text_node, XmlText):
      return 0
    value = text_node.content
    idx2 = 0
    repeating = 0
    while idx2 >= 0:
      idx1 = value.find('{{', idx2)
      idx2 = value.find('}}', idx1) if idx1 >= 0 else -1
      if idx2 > idx1:
        param = value[idx1 + 2:idx2]
        rep0 = self.resolve_repeating(param)
        if rep0 is None:
          rep0 = 0
        repeating = max(repeating, rep0)
        idx2 += 2
    return repeating

  @staticmethod
  def __find_block(tagname: str, block: XmlTag):
    if not isinstance(block, XmlTag):
      return None
    tag = block.name
    if tag == tagname:
      return block
    content = block.elements
    for item in content:
      found = PowerPointDocument.__find_block(tagname, item)
      if found:
        return found
    return None

  def process_html_content(self, blocks: list):
    """
    Expands embedded HTML content in text.

    Args:
      blocks: List of XML elements.
    """
    idx = 0
    while idx < len(blocks):
      block = blocks[idx]
      if not isinstance(block, XmlTag):
        idx += 1
        continue
      tag = block.name
      if tag == TAG_P:
        rblock = PowerPointDocument.__find_block(TAG_R, block)
        if rblock:
          r_elements = rblock.elements
          if len(r_elements) == 0:
            idx += 1
            continue
          t_tag = None
          for rtag in r_elements:
            if not isinstance(rtag, XmlTag):
              continue
            tag = rtag.name
            if tag == TAG_T:
              t_tag = rtag
              break
          if not t_tag or len(t_tag.elements) == 0 or not isinstance(t_tag.elements[0], XmlText):
            self.process_html_content(r_elements)
            idx += 1
            continue
          tcontent = t_tag.elements[0].content
          if tcontent.find('<') < 0:
            idx += 1
            continue
          html_blocks = self.__expand_html_content(tcontent)
          if html_blocks:
            del blocks[idx]
            for html_block in html_blocks:
              blocks.insert(idx, html_block)
              idx += 1
          else:
            idx += 1
          continue
      content = block.elements
      if isinstance(content, list):
        self.process_html_content(content)
      idx += 1

  def __expand_html_content(self, text_html: str) -> list or None:
    out = []
    xml = XmlParser({CONFIG_PARAM_HTML: True})
    blocks = xml.parse_text(text_html)
    for block in blocks:
      if not isinstance(block, XmlTag):
        continue
      tag = block.name
      attrs = block.attrs
      paraprops = {
                'align': None,
                'size': 800,
                'codeblock': None
            }
      classname = attrs.get('class')
      if classname:
      # -- asegura que la clase no tiene el caret
        classname = text_util.trim(classname.replace('caret', ''))
        # -- mapea las classes HTML a estilos WORD
        classes = classname.split(' ')
        for classname in classes:
          if classname == 'left':
            paraprops['align'] = 'l'
          elif classname == 'right':
            paraprops['align'] = 'r'
          elif classname == 'center':
            paraprops['align'] = 'ctr'
          elif classname == 'justify':
            paraprops['align'] = 'just'
          elif classname.startswith('list-'):
            paraprops['list'] = classname[5:]
          elif classname == 'codeblock' or classname == 'token':
            paraprops['codeblock'] = True
          elif classname == 'security-level':
            paraprops['align'] = 'l'
            paraprops['bold'] = True
            paraprops['size'] = 1000
          elif classname == 'link-title':
            paraprops['align'] = 'l'
            paraprops['bold'] = True
            paraprops['size'] = 800
          elif classname == 'link-url':
            paraprops['align'] = 'l'
            paraprops['code'] = True
            paraprops['size'] = 800

      elements = block.elements
      astyle = attrs.get('style')
      if astyle:
        css = html.parse_css(astyle)
        PowerPointDocument.parse_css_properties(css, paraprops)
      if tag == 'pre':
        paraprops['codeblock'] = True
      if tag == 'ul' or tag == 'ol':
        for li in elements:
          if not isinstance(li, XmlTag):
            continue
          self.__create_paragraph(out, paraprops, li.elements)
        continue
      headers = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']
      try:
        idx = headers.index(tag)
      except ValueError:
        idx = -1
      if idx >= 0:
        paraprops['size'] = 800 + ((5 - idx) * 200)
        self.__create_paragraph(out, paraprops, elements)
        continue
      self.__create_paragraph(out, paraprops, elements)
    return out

  def __create_paragraph(self, out: list, paraprops: dict, content: str or list):
    if not isinstance(content, list):
      out.append(doc_elements.create_paragraph(paraprops, content, self.relationships, self.types))
      return
    properties = paraprops.copy()
    if properties.get('codeblock'):
      properties['code'] = True
    runs = []
    self.__process_tag_content(content, runs, properties)
    out.append(doc_elements.create_paragraph(paraprops, runs, self.relationships, self.types))

  def __process_tag_content(self, blocks: list, runs: list, properties: dict):
    for item in blocks:
      if isinstance(item, XmlText):
        runs.append(doc_elements.create_run(item.content, properties, self.relationships, self.types))
        continue
      tag = item.name
      props = properties.copy()
      attrs = item.attrs
      if tag == 'font':
        color = attrs.get('color')
        if color:
          props['color'] = color
      elif tag == 'b':
        props['bold'] = True
      elif tag == 'i':
        props['italic'] = True
      elif tag == 'strike' or tag == 'stroke':
        props['strike'] = True
      elif tag == 'u':
        props['underline'] = True
      elif tag == 'code':
        props['code'] = True
      elif tag == 'a':
        url = attrs.get('href')
        if url:
          props['link'] = url
          props['color'] = '#0000ff'
      elif tag == 'img':
        pass  # -- No se soportan imágenes en línea de párrafo
      astyle = attrs.get('style')
      if astyle:
        css = html.parse_css(astyle)
        PowerPointDocument.parse_css_properties(css, props)

      content = item.elements
      self.__process_tag_content(content, runs, props)

  @staticmethod
  def parse_css_properties(css: dict or None, props):
    """
    Applies CSS properties to a props dictionary.

    Args:
      css: CSS dictionary.
      props: Destination properties dictionary.
    """
    if not css:
      return
    text_align = css.get('text-align')
    if text_align:
      align = None
      if text_align == 'left':
        align = 'l'
      elif text_align == 'right':
        align = 'r'
      elif text_align == 'center':
        align = 'ctr'
      elif text_align == 'justify':
        align = 'just'
      if align:
        props['align'] = align
    font_weight = css.get('font-weight')
    if font_weight:
      props['bold'] = True if font_weight != 'normal' and font_weight != '200' else False
    font_style = css.get('font-style')
    if font_style:
      props['italic'] = True if font_style == 'italic' else False
    color = css.get('color')
    if color:
      color = html.get_rgb_color(color)
      props['color'] = color
    text_decoration = css.get('text-decoration')
    if text_decoration:
      props['underline'] = False
      props['strike'] = False
      if text_decoration == 'underline':
        props['underline'] = True
      elif text_decoration == 'line-through':
        props['strike'] = True
    font_size = css.get('font-size')
    if font_size and font_size.endswith("px"):
      try:
        font_size = int(text_util.trim(font_size[0:len(font_size) - 2])) * 100
        props['size'] = font_size
      except ValueError:
        pass
    pxhg = css.get('height')
    if pxhg and pxhg.endswith("px"):
      try:
        pxhg = int(text_util.trim(pxhg[0:len(pxhg) - 2]))
        props['height'] = pxhg
      except ValueError:
        pass
    pxwd = css.get('width')
    if pxwd and pxwd.endswith("px"):
      try:
        pxwd = int(text_util.trim(pxwd[0:len(pxwd) - 2]))
        props['width'] = pxwd
      except ValueError:
        pass
