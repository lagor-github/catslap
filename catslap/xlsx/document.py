# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)
import json

from catslap.base.relationships import RELATIONSHIP_TYPE_TABLE
from catslap.xlsx.sharedstrings import SharedStrings
from catslap.base.types import ContentTypes
from catslap.base.document import Document
from catslap.base.relationships import Relationships
from catslap.pptx import elements as doc_elements
from catslap.utils import types as types
from catslap.utils import text as text_util
from catslap.utils import encoding as enc_util
from catslap.utils import file as file_util
from catslap.utils import html
from catslap.utils.xml import XmlParser, XmlTag, XmlText, CONFIG_PARAM_INCLUDE_DECL, CONFIG_PARAM_HTML


EXCEL_WORKBOOK = "/xl/workbook.xml"
EXCEL_WORKBOOK_RELS = "/xl/_rels/workbook.xml.rels"
EXCEL_WORKSHEETS = "/xl/worksheets"
EXCEL_WORKSHEETS_RELS = "/xl/worksheets/_rels"
EXCEL_WORKSHEETS_TABLES = "/xl/worksheets/tables"
EXCEL_CALC_CHAIN = "/xl/calcChain.xml"
EXCEL_SHARED_STRINGS = "/xl/sharedStrings.xml"
EXCEL_DRAWINGS = "/xl/drawings/"
EXCEL_CHARTS = "/xl/charts/"
EXCEL_CHARTS_RELS = EXCEL_CHARTS + "_rels/"
DOT_RELS = '.rels'
EXCEL_DOCUMENT_TYPES = "[Content_Types].xml"

# -- puntos de powerpoint por cm
PT_PER_CM = 360000

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


class ExcelException(Exception):
  """
  Excel processing specific exception.
  """
  pass


class ExcelDocument(Document):
  """
  Processes and generates Excel documents (.xlsx).

  Attributes:
    max_id: Maximum ID used in the document.
    sharedstrings: Input SharedStrings.
    output_sharedstrings: Output SharedStrings.
    relationships: Relationships for the current sheet.
    types: Package ContentTypes.
  """
  def __init__(self, file: str):
    super().__init__(file)
    self.max_id = 0
    self.sharedstrings = None
    self.output_sharedstrings = None
    self.relationships = None
    self.types = None

  def process_template(self, tempdir: str):
    """
    Processes the XLSX template in the temporary directory.

    Args:
      tempdir: Temporary directory with ZIP contents.

    Raises:
      XmlParserException: If XML is invalid.
      OSError: If file read/write fails.
    """
    # -- carga sharedStrings
    self.sharedstrings = SharedStrings(tempdir)
    # --sharedStrings de salida
    self.output_sharedstrings = SharedStrings()
    self.output_sharedstrings.pathfile = self.sharedstrings.pathfile

    ppt_types = tempdir + "/" + EXCEL_DOCUMENT_TYPES
    self.types = ContentTypes(ppt_types)

    # -- procesa sheet-names del workbook
    workbook_file_rels = self.tempdir + EXCEL_WORKBOOK_RELS
    workbook_rels = Relationships(self.tempdir, workbook_file_rels)
    workbook_file = self.tempdir + EXCEL_WORKBOOK

    parser = XmlParser()
    workbook = parser.parse_file(workbook_file, 'workbook')
    sheets_tag = workbook.get_tag('sheets')

    sheet_tags = sheets_tag.get_tags('sheet')
    max_sheet_id = 0
    for sheet in sheet_tags:
      max_sheet_id = max(max_sheet_id, int(sheet.get_attr('sheetId')))

    rewrite = False
    sheet_files = []
    sheet_tags = sheets_tag.elements
    for idx in range(0, len(sheet_tags)):
      sheet = sheet_tags[idx]
      if not isinstance(sheet, XmlTag):
        continue
      rid = sheet.get_attr('r:id')
      relationship = workbook_rels.get_relationship_by_id(rid) if rid is not None else None
      if relationship is not None:
        sheet_file = relationship.target
        name = sheet.get_attr('name')
        name0 = self.resolve_text(idx, name)
        if name != name0:
          sheet.set_attr('name', name0)
          name = name0
          rewrite = True
        sheet_files.append({'name': name, 'file': sheet_file})

    if rewrite:
      parser.write_file()

    for obj in sheet_files:
      sheet_file = obj.get('file')
      self.__process_sheet(sheet_file)

      # -- procesa los drawings
    self.__process_drawings(tempdir)

    # -- escribe las sharedstring actualizadas
    self.output_sharedstrings.write()
    # -- borrar archivo clacChain para que actualice las fórmulas
    file_util.remove_file(tempdir + EXCEL_CALC_CHAIN)

    files = file_util.list_files(tempdir + EXCEL_CHARTS)
    if files:
      for file in files:
        if file.startswith("chart") and file.endswith('.xml'):
          chart_file = tempdir + '/' + EXCEL_CHARTS + file
          self.clear_chart_values(chart_file, self)

  def extract_data(self, ref: str) -> list:
    """
    Extracts data from an Excel range.

    Args:
      ref: Reference like 'Sheet1!A1:B3'.

    Returns:
      Extracted data matrix.

    Raises:
      ExcelException: If the reference is invalid.
    """
    sheet_name, sdata, edata = parse_data_ref(ref)
    scell_num = get_cell_num(sdata)
    srow_num = get_row_num(sdata)
    ecell_num = get_cell_num(edata) if edata is not None else None
    erow_num = get_row_num(edata) if edata is not None else None
    return self.extract_from_sheet_name(sheet_name, scell_num, srow_num, ecell_num, erow_num)

  def write_cells(self, data):
    """
    Writes data to the main sheet.

    Args:
      data: Data matrix.

    Raises:
      OSError: If file writing fails.
    """
    # -- carga sharedStrings
    sharedstrings = SharedStrings(self.tempdir)
    # -- obtiene el nombre del archivo excel a partir del nombre de hoja
    sheet_file = self.tempdir + '/xl/worksheets/sheet1.xml'
    # -- procesa la hoja
    parser = XmlParser()
    try:
      worksheet = parser.parse_file(sheet_file, 'worksheet')
      sheet_data = worksheet.get_tag('sheetData')
      max_rows = len(data)
      max_num_cells = 1
      for nrow, row in enumerate(data):
        max_num_cells = max(max_num_cells, len(row))
        tag_row = sheet_data.add_tag('row', {'r': str(nrow + 1), 'spans': '1:' + str(len(row)), 'x14ac:dyDescent': '0.2'})
        for ncell, cell in enumerate(row):
          cellpos = get_cell_format_position(ncell + 1, nrow + 1, False)
          if text_util.is_numeric(cell):
            tag_cell = tag_row.add_tag('c', {'r': cellpos})
            tag_cell.add_tag_text('v', str(cell))
          else:
            tag_cell = tag_row.add_tag('c', {'r': cellpos, 't': 's'})
            idx = sharedstrings.add_string(cell)
            tag_cell.add_tag_text('v', idx)
            # -- ajusta la dimensión de la hoja
      dim_tag = worksheet.get_tag('dimension')
      if dim_tag:
        dim_tag.set_attr('ref', get_dimension(max_num_cells, max_rows))

    finally:
      if not self.test_mode:
        parser.write_file()
        sharedstrings.write()

  def extract_from_sheet_name(self, sheet_name: str, x1: int, y1: int, x2: int| None, y2: int| None) -> list:
    """
    Extracts data from a sheet by coordinates.

    Args:
      sheet_name: Sheet name.
      x1: Start column.
      y1: Start row.
      x2: End column (optional).
      y2: End row (optional).

    Returns:
      Extracted data matrix.
    """
    if x2 is None:
      x2 = x1
    # -- carga sharedStrings
    sharedstrings = SharedStrings(self.tempdir)
    # -- obtiene el nombre del archivo excel a partir del nombre de hoja
    try:
      rid = self.get_sheet_rid(sheet_name)
      sheet_file = self.get_sheet_file_by_rid(rid)
    except ExcelException:
      sheet_file = self.tempdir + '/xl/worksheets/sheet1.xml'

      # -- procesa la hoja
    parser = XmlParser()
    worksheet = parser.parse_file(sheet_file, 'worksheet')
    sheet_data = worksheet.get_tag('sheetData', False)
    data = []
    if sheet_data:
      rows = sheet_data.get_tags('row')
      rowpos = 0
      vrowpos = y1
      while rowpos < len(rows):
        row = rows[rowpos]
        rrow_num = types.to_int(row.get_attr('r'))
        if vrowpos > rrow_num:
          rowpos += 1
          continue
        # -- llena de datos vacios el resultado hasta la posición de la row
        while vrowpos < rrow_num and y2 is not None and vrowpos <= y2:
          data.append([])
          vrowpos += 1
        if y2 is not None and vrowpos > y2:
          break
        cells = row.get_tags('c')
        celldata = []
        cellpos = 0
        vcellpos = x1
        while cellpos < len(cells):
          cell = cells[cellpos]
          rvalue = cell.get_attr('r')
          rcell_num = get_cell_num(rvalue)
          if vcellpos > rcell_num:
            cellpos += 1
            continue
            # -- lleva de datos vacios el resultado hasta la posición de la celda
          while vcellpos < rcell_num and vcellpos <= x2:
            celldata.append(None)
            vcellpos += 1
          if vcellpos > x2:
            break
          ctype = cell.get_attr('t')
          value = cell.get_tag_text('v', False)
          if ctype is not None and text_util.is_numeric(value):
            idx = int(value)
            value = sharedstrings.get_string(idx)
          if value is not None:
            celldata.append(value)
          vcellpos += 1
          cellpos += 1
        data.append(celldata)
        vrowpos += 1
        rowpos += 1
    while len(data) > 0 and len(data[len(data) - 1]) == 0:
      del data[len(data) - 1]
    return data

  def get_sheet_rid(self, sheet_name: str) -> str:
    """
    Gets the rId of a sheet by name.

    Args:
      sheet_name: Sheet name.

    Returns:
      Sheet rId.

    Raises:
      ExcelException: If the sheet is not found.
    """
    workbook_file = self.tempdir + EXCEL_WORKBOOK
    parser = XmlParser()
    workbook = parser.parse_file(workbook_file, 'workbook')
    sheets_tag = workbook.get_tag('sheets')
    sheet_tags = sheets_tag.get_tags('sheet')
    for sheet in sheet_tags:
      name = sheet.get_attr('name')
      if name == sheet_name:
        return sheet.get_attr('r:id')
    raise ExcelException(f"rId not found for sheet '{sheet_name}'")

  def get_sheet_file_by_rid(self, sheer_rid: str) -> str:
    """
    Gets the sheet file from an rId.

    Args:
      sheer_rid: rId identifier.

    Returns:
      Sheet file path.

    Raises:
      ExcelException: If the relationship does not exist.
    """
    workbook_file_rels = self.tempdir + EXCEL_WORKBOOK_RELS
    relationships = Relationships(self.tempdir, workbook_file_rels)
    relationship = relationships.get_relationship_by_id(sheer_rid)
    if relationship is None:
      raise ExcelException(f"Relationship not found for sheet rId {sheer_rid}")
    return relationship.target

  def __process_sheet(self, sheet_file: str):
    spath = file_util.get_pathname(sheet_file)
    sfile = file_util.get_filename(sheet_file)
    filesheet_rel = spath + '_rels/' + sfile + '.rels'
    relationships = Relationships(self.tempdir, filesheet_rel)
    sheet_doc = Worksheet(sheet_file)
    sheet_data_tag = sheet_doc.root_tag.get_tag('sheetData')
    rows = sheet_data_tag.elements

    # -- procesa las tablas relacionadas para incluir los campos vacios de columna
    if relationships and not self.test_mode:
      relationship_list = relationships.get_relationships(RELATIONSHIP_TYPE_TABLE, None)
      for relationship in relationship_list:
        tablefile = relationship.target
        ExcelDocument.__process_sheet_table_columns(tablefile, rows)

    # -- resuelve todas las celdas con el valor real del shared-string
    for row in rows:
      if not isinstance(row, XmlTag):
        continue
      cells = row.elements
      for cell in cells:
        if not isinstance(cell, XmlTag):
          continue
        ctype = cell.get_attr("t")
        value = cell.get_tag_text('v', False)
        if ctype == 's' and text_util.is_numeric(value):
          cell_id = int(value)
          value = self.sharedstrings.get_string(cell_id)
          if value is not None:
            value = text_util.trim(value)
            cell.set_tag_text('v', value, False)
        cell.remove_attr('t')

    # -- crea las nuevas filas que harán falta para las variables múltiples
    cur_row = 0
    while cur_row < len(rows):
      row = rows[cur_row]
      if not isinstance(row, XmlTag):
        cur_row += 1
        continue
      at_row = types.to_int(row.get_attr("r"))
      expected_at_row = cur_row + 1
      # -- faltan filas respecto la fila encontrada
      if expected_at_row < at_row:
        row0 = XmlTag('row')
        row0.set_attr('r', expected_at_row)
        rows.insert(cur_row, row0)
        cur_row += 1
        continue
      cells = row.elements
      num_cell = 0
      repeating = 0
      # -- procesa las celda para ver si exiten variables de filas múltiples
      while num_cell < len(cells):
        cell = cells[num_cell]
        num_cell += 1
        if not isinstance(cell, XmlTag):
          continue
        value = cell.get_tag_text('v', False)
        if text_util.is_empty(value):
          continue
          # -- verifica si hay variable a procesar
        idx = value.find('{{')
        if idx < 0:
          continue
        repeating = max(repeating, self.__resolve_cell_repeating(value, False))
        # -- hay creación de nuevas filas vacias
      if repeating > 1:
        cur_row += 1
        first_row_at = cur_row
        for f in range(0, repeating - 1):
          at_row = cur_row + 1
          row0 = XmlTag('row')
          row0.set_attr('r', at_row)
          rows.insert(cur_row, row0)
          cells0 = row0.elements
          num_cell = 0
          while num_cell < len(cells):
            cell = cells[num_cell]
            num_cell += 1
            if not isinstance(cell, XmlTag):
              continue
            value = cell.get_tag_text('v', False)
            idx = value.find('{{') if value is not None else -1
            if idx >= 0:
              continue
            cell0 = cell.clone(True)
            rvalue = cell.get_attr('r')
            at_cell = get_cell_num(rvalue)
            cell0.set_attr('r', get_cell_format_position(at_cell, at_row))
            formula = cell.get_tag_text('f', False)
            if formula is not None:
              text_row = str(first_row_at)
              idx = formula.find(text_row)
              while idx >= 0 and idx < len(formula) - 1 and not text_util.is_numeric(formula[idx-1]) and \
              not text_util.is_numeric(formula[idx + len(text_row)]):
                formula = formula[:idx] + str(at_row) + formula[idx + len(text_row):]
                cell0.set_tag_text('f', formula)
                idx = formula.find(text_row)

            cells0.append(cell0)
          cur_row += 1

        # -- ajusta los alcances de las validaciones
        extlst_tag = sheet_doc.root_tag.get_tag('extLst', False)
        ext_tag = extlst_tag.get_tag('ext') if extlst_tag else None
        datavals_tag = ext_tag.get_tag('x14:dataValidations', False) if ext_tag else None
        if datavals_tag is not None:
          dataval_tags = datavals_tag.get_tags('x14:dataValidation') if datavals_tag else []
          for dataval_tag in dataval_tags:
            sqref = dataval_tag.get_tag_text('xm:sqref', False)
            if not sqref:
              continue
            sqrefs = sqref.split(' ')
            stext = ''
            for sqref in sqrefs:
              new_sqref = adjust_cell_row_by(sqref, first_row_at, repeating)
              stext += new_sqref + ' '
            dataval_tag.set_tag_text('xm:sqref', text_util.trim(stext))

        # -- ajusta alcance de las consolidaciones
        conso_tags = sheet_doc.root_tag.get_tags('conditionalFormatting')
        for conso in conso_tags:
          sqref = conso.get_attr('sqref')
          if not sqref:
            continue
          new_sqref = adjust_cell_row_by(sqref, first_row_at, repeating)
          conso.set_attr('sqref', text_util.trim(new_sqref))

        # -- reposiciona las posiciones de las celdas por las filas desplazadas
        back_row = cur_row
        total_sum = repeating - 1
        for f in range(cur_row, len(rows)):
          at_row += 1
          row0 = rows[f]
          at_row = types.to_int(row0.get_attr('r')) + total_sum
          # -- reposiciona la fila
          row0.set_attr('r', at_row)  
          cells = row0.elements
          num_cell = 0
          while num_cell < len(cells):
            cell = cells[num_cell]
            rvalue = cell.get_attr("r")
            at_cell = get_cell_num(rvalue)
            formula = cell.get_tag_text("f", False)
            if formula is not None:
              idx1 = formula.find('(')
              idx2 = formula.find(')')
              if idx1 > 0 and idx2 > idx1:
                inner = formula[idx1 + 1:idx2]
                outer = formula[:idx1]
                changed = False
                params = inner.split(';')
                for idx1 in range(0, len(params)):
                  param = params[idx1]
                  idx2 = param.find(':')
                  if idx2 > 0:
                    pos = get_row_num(param[idx2 + 1:])
                    if pos < at_row:
                      param = param[0:idx2 + 1] + get_cell_format_position(at_cell, pos + total_sum)
                      params[idx1] = param
                      changed = True
                if changed:
                  formula = outer + '(' + ';'.join(params) + ')'
                  cell.set_tag_text("f", formula)
                  # -- reposiciona la columna
            cell.set_attr('r', get_cell_format_position(at_cell, at_row))
            num_cell += 1
        cur_row = back_row
        continue
      cur_row += 1

    # -- procesa todas las celdas
    num_row = 0
    while num_row < len(rows):
      row = rows[num_row]
      num_row += 1
      if not isinstance(row, XmlTag):
        continue
      num_cell = 0
      cells = row.elements
      # -- procesa las celdas de la fila
      while num_cell < len(cells):
        cell = cells[num_cell]
        num_cell += 1
        if not isinstance(cell, XmlTag):
          continue
        is_formula = False
        value = cell.get_tag_text('f', False)
        if value is not None:
          is_formula = True
          cell.remove_tag('v')
        else:
          value = cell.get_tag_text('v', False)  # está resuelto
        tvalue = cell.get_attr("t")
        # -- comprueba si ya está indexada para ignorar la celda
        if tvalue is not None:
          continue
        rvalue = cell.get_attr("r")
        cstyle = cell.get_attr("s")
        at_row = get_row_num(rvalue)
        at_cell = get_cell_num(rvalue)
        if not text_util.is_empty(value):
          idx = value.find('{{')
          repeating = self.__resolve_cell_repeating(value, True) if idx >= 0 else 0
          # -- resuelve la repetición del campo
          if repeating > 0:
            idx += 2
            atpos1 = value.find('!', idx)
            atpos2 = value.find('}}', atpos1) if atpos1 >= idx else -1
            # -- comprueba reposicionamiento de variable
            if atpos2 > atpos1:
              self.set_cell_value(cstyle, at_row, at_cell, '', rows, is_formula)
              cell_value = value[idx:atpos1]
              at = value[atpos1 + 1:atpos2]
              value = "{{" + cell_value + "}}"
              at_cell = get_cell_num(at)
              at_row = get_row_num(at)
            resolved_value = self.__resolve_cell_value(None, value)
            # -- resuelve la celda en modo diagrama (si lo es)
            if isinstance(resolved_value, dict):
              diagram = resolved_value
              categories = diagram.get('categories')
              series = diagram.get('series')
              if series and categories:
                self.set_cell_value(cstyle, at_row, at_cell, 'Categories', rows, False)
                for idx2, category in enumerate(categories):
                  self.set_cell_value(cstyle, at_row + idx2 + 1, at_cell, category, rows, False)
                for idx1, serie in enumerate(series):
                  data = serie.get('data')
                  name = serie.get('name')
                  if name is None:
                    name = 'Serie' + str(idx1 + 1)
                  self.set_cell_value(cstyle, at_row, at_cell + idx1 + 1, str(name), rows, False)
                  for idx2, data_value in enumerate(data):
                    self.set_cell_value(cstyle, at_row + idx2 + 1, at_cell + idx1 + 1, str(data_value), rows, False)
              continue
              # -- resuelve la celda de forma normal
            self.set_cell_value(cstyle, at_row, at_cell, resolved_value, rows, is_formula)
            for f in range(1, repeating):
              resolved_value = self.__resolve_cell_value(f, value)
              self.set_cell_value(cstyle, at_row + f, at_cell, resolved_value, rows, is_formula)
            continue
        self.set_cell_value(cstyle, at_row, at_cell, value, rows, is_formula)

      row.remove_attr('spans')
      row.remove_attr('x14ac:dyDescent')

    if self.test_mode:
      return

    # -- ajusta altura de las rows
    max_num_cells = 0
    max_rows = 0
    for row in rows:
      if not isinstance(row, XmlTag):
        continue
      max_rows += 1
      hg = types.to_int(row.get_attr('ht'), 16)
      # -- recupera las celdas
      cells = row.elements
      # -- procesa las celdas de la fila
      num_cells = 0
      for cell in cells:
        if not isinstance(cell, XmlTag):
          continue
        num_cells += 1
        ctype = cell.get_attr("t")
        value = cell.get_tag_text('v', False)  # está resulto
        formula = cell.get_tag_text('f', False)
        if formula:
          cell.remove_tag('v')
          value = ''
        if value is None:
          value = ''
        value = text_util.trim(str(value))
        # -- si está referenciado por tipo en el archivo de strings
        if ctype is not None and text_util.is_numeric(value):
          cell_id = int(value)
          value = self.output_sharedstrings.get_string(cell_id)
          value = text_util.trim(value)
        if formula or not text_util.is_empty(value):
          max_num_cells = max(max_num_cells, num_cells)
        hg = max(hg, ExcelDocument.calculate_value_hg(value))
        # -- row.set_attr('customFormat', 1) Esto da problemas de vinculación en PPTX
      row.set_attr('ht', str(hg))

    dimension = get_dimension(max_num_cells, max_rows)
    # -- ajusta la dimensión de la hoja
    dim_tag = sheet_doc.root_tag.get_tag('dimension')
    if dim_tag:
      dim_tag.set_attr('ref', dimension)

    # -- quita la selección de hojas
    sviews_tag = sheet_doc.root_tag.get_tag('sheetViews')
    if sviews_tag:
      sview_tags = sviews_tag.get_tags('sheetView')
      for sheet_view in sview_tags:
        sheet_view.remove_tags('selection')

    # -- procesa las tablas relacionadas para incluir el rango de la tabla
    if relationships:
      relationship_list = relationships.get_relationships(RELATIONSHIP_TYPE_TABLE, None)
      for relationship in relationship_list:
        tablefile = relationship.target
        ExcelDocument.__process_sheet_table_range(tablefile, dimension)
    # -- escribe el archivo final
    sheet_doc.write_file()

  @staticmethod
  def calculate_value_hg(value: str):
    """
    Calculates an approximate height based on content.

    Args:
      value: Cell value.

    Returns:
      Approximate height in points.
    """
    # -- cuenta retornos para calcular alto
    total_lf = text_util.count_lf(value)
    if total_lf == 0:
      total_lf = 1
    return 16 * total_lf

  def set_cell_value(self, cstyle: str| None, at_row: int, at_cell: int, value: str, rows: list, is_formula: bool):
    """
    Writes a value in a specific cell.

    Args:
      cstyle: Cell style.
      at_row: Row (1-based).
      at_cell: Column (1-based).
      value: Value to assign.
      rows: List of XML rows.
      is_formula: Whether the value is a formula.
    """
    # -- asegura que la fila existe y tienen todas suficientes celdas
    len_rows = len(rows)
    while at_row > len_rows:
      row0 = XmlTag('row')
      row0.set_attr('r', str(len_rows + 1))
      rows.append(row0)
      len_rows += 1
    # -- asegura que la celda existe
    row0 = rows[at_row - 1]
    if not isinstance(row0, XmlTag):
      return
    cells = row0.elements
    num_cell = 0
    while num_cell < len(cells):
      cell = cells[num_cell]
      if not isinstance(cell, XmlTag):
        num_cell += 1
        continue
      rvalue = cell.get_attr("r")
      at_cell2 = get_cell_num(rvalue)
      # -- me he pasado la posición
      if at_cell2 > at_cell:
        cell = XmlTag('c')
        cell.set_attr('r', get_cell_format_position(at_cell, at_row))
        cells.insert(num_cell, cell)
        at_cell2 = at_cell
        # -- estoy en la posición
      if at_cell2 == at_cell:
        cell.clear_tags()
        if cstyle is not None:
          cell.add_attr('s', cstyle)
        if is_formula:
          if not text_util.is_empty(value):
            cell.add_tag_text('f', value)
        elif text_util.is_numeric(value):
          cell.remove_attr('t')
          if not text_util.is_empty(value):
            cell.add_tag_text('v', value)
        else:
          if not text_util.is_empty(value):
            idx = str(self.output_sharedstrings.add_string(value))
            cell.add_tag_text('v', str(idx))
            cell.set_attr('t', 's')
        return
      num_cell += 1
    # -- no se ha encontrado la posición
    cell0 = XmlTag('c')
    cell0.set_attr('r', get_cell_format_position(at_cell, at_row))
    if cstyle is not None:
      cell0.add_attr('s', cstyle)
    if is_formula:
      if not text_util.is_empty(value):
        cell0.add_tag_text('f', value)
    elif text_util.is_numeric(value):
      cell0.remove_attr('t')
      if not text_util.is_empty(value):
        cell0.add_tag_text('v', value)
    else:
      if not text_util.is_empty(value):
        idx = str(self.output_sharedstrings.add_string(value))
        cell0.add_tag_text('v', str(idx))
        cell0.set_attr('t', 's')
    cells.append(cell0)

  def __resolve_cell_value(self, row: int| None, value: str) -> any:
    stext = ''
    idx0 = 0
    while idx0 >= 0:
      idx1 = value.find('{{', idx0)
      idx2 = value.find('}}', idx1) if idx1 >= 0 else -1
      if idx2 <= idx1:
        break
      stext = stext + value[idx0:idx1]
      param = value[idx1 + 2:idx2]
      atpos = param.find('!')
      if atpos > 0:
        param = param[:atpos]
      resolved = self.resolve_value(row, param)
      if isinstance(resolved, dict):
        return resolved
      if isinstance(resolved, list):
        if row is None:
          row = 0;
        resolved = resolved[row]
      if resolved is not None:
        stext = stext + str(resolved)
      idx0 = idx2 + 2
    stext = stext + value[idx0:]
    return stext

  def __resolve_cell_repeating(self, value: str, include_overwrite: bool) -> int:
    idx0 = 0
    repeating = 0
    while idx0 >= 0:
      idx1 = value.find('{{', idx0)
      idx2 = value.find('}}', idx1) if idx1 >= 0 else -1
      if idx2 <= idx1:
        break
      param = value[idx1 + 2:idx2]
      atpos = param.find('!')
      if atpos > 0:
        if not include_overwrite:
          idx0 = idx2 + 2
          continue
        param = param[:atpos]
      if repeating == 0:
        repeating = 1
      rep0 = self.resolve_repeating(param)
      if rep0 is None:
        rep0 = 0
      repeating = max(repeating, rep0)
      idx0 = idx2 + 2
    return repeating

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
          if last_t is not None and ExcelDocument.__is_the_same_rpr(last_rpr, rpr):
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
  def __is_the_same_rpr(rpr1: XmlTag| None, rpr2: XmlTag| None) -> bool:
    if rpr1 is None and rpr2 is None:
      return True
    xml = XmlParser()
    rpr1 = rpr1.clone(True)
    rpr2 = rpr2.clone(True)
    ExcelDocument.__remove_ignorable_style_tags(rpr1)
    ExcelDocument.__remove_ignorable_style_tags(rpr2)
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
      ExcelDocument.__remove_ignorable_style_tags(element)
      idx += 1

  def __process_drawings(self, tempdir: str):
    xml = XmlParser()
    # -- parsea los drawings
    idx = 1
    found = True
    while found:
      drawing_file = tempdir + '/' + EXCEL_DRAWINGS + "drawing" + str(idx) + ".xml"
      found = file_util.exist(drawing_file)
      if found:
        file0_rel = tempdir + '/' + EXCEL_DRAWINGS + "_rels/drawing" + str(idx) + ".xml.rels"
        self.relationships = Relationships(tempdir, file0_rel)
        document = xml.parse_file(drawing_file, 'xdr:wsDr')
        blocks = document.elements
        self.collapse_paragraphs(blocks)
        self.max_id = max(self.max_id, xml.max_id)
        self.search_graphic_frames(tempdir, blocks)
        self.process_descr_attrs(blocks)
        self.process_paragraphs(blocks)
        self.process_html_content(blocks)
        xml_content = xml.get_pretty_xml(document, {CONFIG_PARAM_INCLUDE_DECL: True})
        file_util.write_bytes(drawing_file, bytes(xml_content, enc_util.UTF_8))
        if not self.test_mode:
          self.relationships.write_file()
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
    Finds and processes graphic frames in the document.

    Args:
      tempdir: Temporary directory.
      elements0: List of XML elements.
    """
    idx0 = 0
    while idx0 < len(elements0):
      element0 = elements0[idx0]
      if isinstance(element0, XmlTag):
        if element0.name == 'xdr:graphicFrame':
          self.__process_graphic_frame(tempdir, element0)
        else:
          self.search_graphic_frames(tempdir, element0.elements)
      idx0 += 1

  def __process_graphic_frame(self, tempdir: str, element0):
    cnvpr = element0.get_tag_path(['xdr:nvGraphicFramePr', 'xdr:cNvPr'])
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

      # -- modifica chart
    chart_file = tempdir + '/' + EXCEL_CHARTS + tfile
    parser = XmlParser()
    root_tag = parser.parse_file(chart_file)
    title_tag = root_tag.get_tag_path(['c:chart', 'c:title', 'c:tx', 'c:rich', 'a:p', 'a:r', 'a:t'], False)
    if title_tag:
      title_tag.set_text(diagram.get('title'))
    chart_tag = root_tag.get_tag_path(['c:chart', 'c:plotArea', '*Chart'], False)
    if chart_tag is not None:
      ser_tags = chart_tag.get_tags('c:ser')
      for sidx, ser_tag in enumerate(ser_tags):
        cat_tag = ser_tag.get_tag('c:cat', False)
        if not cat_tag:
          continue
        cat_tag.clear_tags()
        str_lit = cat_tag.add_element(XmlTag('c:strLit'))
        str_lit.add_element(XmlTag('c:ptCount', {'val': len(cats)}))
        for pidx in range(0, len(cats)):
          pt_tag = str_lit.add_element(XmlTag('c:pt', {'idx': pidx}))
          pt_tag.add_element(XmlTag('c:v')).set_text(cats[pidx])
        data = series[sidx].get('data')
        val_tag = ser_tag.get_tag('c:val', False)
        if not val_tag:
          continue
        val_tag.clear_tags()
        num_lit = val_tag.add_element(XmlTag('c:numLit'))
        num_lit.add_element(XmlTag('c:formatCode')).set_text('General')
        num_lit.add_element(XmlTag('c:ptCount', {'val': len(data)}))
        for pidx in range(0, len(data)):
          pt_tag = num_lit.add_element(XmlTag('c:pt', {'idx': pidx}))
          pt_tag.add_element(XmlTag('c:v')).set_text(data[pidx])
    parser.write_file()

  @staticmethod
  def __process_sheet_table_range(tablefile, maxcellrow: str):
    xml = XmlParser()
    table_tag = xml.parse_file(tablefile, 'table')
    table_tag.set_attr('ref', maxcellrow)
    autofilter = table_tag.get_tag('autoFilter', False)
    if autofilter:
      autofilter.set_attr('ref', maxcellrow)
    xml.write_file()

  @staticmethod
  def __process_sheet_table_columns(tablefile, table_rows: list):
    xml = XmlParser()
    table_tag = xml.parse_file(tablefile, 'table')
    changed = False
    columns_tag = table_tag.get_tag('tableColumns')
    if columns_tag:
      column_tags = columns_tag.get_tags('tableColumn')
      if column_tags:
        cpos = 1
        for column_tag in column_tags:
          name = column_tag.get_attr('name')
          # -- IMPORTANTE! Hay que asegurar que ningún nombre de columna de tabla sea nulo/vacío/espacio
          name = text_util.trim(name) if name is not None else ''
          if name == '':
            column_name = 'Column' + str(cpos)
            column_tag.set_attr('name', column_name)
            # -- Hay que asegurar que el nombre está en la celda de la hoja
            cells = table_rows[0].get_tags('c')
            cell = cells[cpos - 1]
            vtag = cell.get_tag('v')
            vtag.set_text(column_name)
            cell.remove_attr('t')  # Elimina el type para evitar que use el texto como índice
            changed = True
          cpos += 1
    if changed:
      xml.write_file()

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
        if element.name != 'xdr:cNvPr':
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
          if element1.name == 'xdr:spPr':
            sppr = element1
          if element1.name == 'xdr:txBody':
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
      self.max_id = ExcelDocument.__reassign_ids(tr0, self.max_id)
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
      maxid = ExcelDocument.__reassign_ids(item, maxid)
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

  def __resolve_text_value(self, row: int| None, block: XmlTag) -> (bool, bool, bool):
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
      found = ExcelDocument.__find_block(tagname, item)
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
        rblock = ExcelDocument.__find_block(TAG_R, block)
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

  def __expand_html_content(self, text_html: str) -> list| None:
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
        ExcelDocument.parse_css_properties(css, paraprops)
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
        ExcelDocument.parse_css_properties(css, props)

      content = item.elements
      self.__process_tag_content(content, runs, props)

  @staticmethod
  def parse_css_properties(css: dict| None, props):
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


def get_cell_num(rvalue: str) -> int:
  """
  Converts a cell reference to a column number.

  Args:
    rvalue: Cell reference (e.g., 'C3').

  Returns:
    Column number (1-based).
  """
  chars = get_cell_alphas(rvalue)
  mul = len(chars) - 1
  if mul < 0:
    return 0
  value = 0
  for char in chars:
    pos = ord(char) - ord('A')
    if mul == 0:
      value += pos
    else:
      value += ((pos + 1) * 26 * mul)
    mul -= 1
  return value + 1


def get_row_num(rvalue: str) -> int:
  """
  Converts a cell reference to a row number.

  Args:
    rvalue: Cell reference (e.g., 'C3').

  Returns:
    Row number (1-based).
  """
  chars = get_cell_digits(rvalue)
  if len(chars) == 0:
    return 0
  return int(chars)


def get_cell_alphas(rvalue: str) -> str:
  """
  Extracts the alphabetic part of a cell reference.

  Args:
    rvalue: Cell reference.

  Returns:
    Column letters.
  """
  if rvalue is not None:
    idx = 0
    while idx < len(rvalue) and text_util.is_alpha(rvalue[idx]):
      idx += 1
    return rvalue[0:idx]
  return ''


def get_cell_digits(rvalue: str) -> str:
  """
  Extracts the numeric part of a cell reference.

  Args:
    rvalue: Cell reference.

  Returns:
    Row digits as text.
  """
  if rvalue is not None:
    idx = len(rvalue) - 1
    while idx >= 0 and text_util.is_numeric(rvalue[idx]):
      idx -= 1
    return rvalue[idx + 1:]
  return ''


def get_cell_format_position(cell: int, row: int| None = None, dollar: bool = False) -> str:
  """
  Converts column and row to a cell reference.

  Args:
    cell: Column number (1-based).
    row: Row number (1-based) or None.
    dollar: If True, uses absolute references ($A$1).

  Returns:
    Excel-style cell reference.
  """
  cell -= 1
  chars = "$" if dollar else ""
  while cell >= 0:
    cell, rest = divmod(cell, 26)
    chars += chr(65 + rest)
    if cell == 0:
      break
  if row is not None:
    if dollar:
      chars += "$"
    chars += str(row)
  return chars


def get_dimension(max_num_cells: int, max_rows: int) -> str:
  """
  Calculates the dimension range for a sheet.

  Args:
    max_num_cells: Maximum number of columns.
    max_rows: Maximum number of rows.

  Returns:
    Range in A1:Z99 format.
  """
  dimpos = get_cell_format_position(max_num_cells, max_rows)
  if dimpos == '0':
    return 'A1'
  return 'A1:' + dimpos


def parse_data_ref(ref: str) -> tuple[str, str, str]:
  """
  Parses an Excel range reference.

  Args:
    ref: Reference like 'Sheet1!$A$1:$B$2'.

  Returns:
    Tuple (sheet_name, start_cell, end_cell).

  Raises:
    ExcelException: If the reference is invalid.
  """
  ref = text_util.trim(ref)
  idx = ref.find('!')
  error_base = f"Invalid excel data '{ref}' reference: "
  if idx < 0:
    raise ExcelException(error_base + "'!' symbol expected")
  if idx == 0:
    raise ExcelException(error_base + "No sheet name found")
  sheet_name = text_util.trim(ref[0:idx])
  ref = text_util.trim(ref[idx + 1:])
  idx = ref.find(':')
  if idx < 0:
    start_data = ref.replace('$', '')
    return sheet_name, start_data, None
  if idx == 0:
    raise ExcelException(error_base + "No start data found")
  if idx == len(ref) - 1:
    raise ExcelException(error_base + "No end data found")
  start_data = ref[0:idx].replace('$', '')
  end_data = ref[idx + 1:].replace('$', '')
  return sheet_name, start_data, end_data


class Worksheet(XmlParser):
  """
  Represents a worksheet (worksheet.xml).
  """
  def __init__(self, pathfile: str):
    super().__init__()
    self.parse_file(pathfile, 'worksheet')


def reassign_ids(block: XmlTag, attr_id: str, maxid: int) -> int:
  """
  Reassigns IDs in an XML tree.

  Args:
    block: Root tag.
    attr_id: ID attribute name.
    maxid: Current maximum ID.

  Returns:
    New maximum ID.
  """
  if not isinstance(block, XmlTag):
    return maxid
  elements = block.elements
  for item in elements:
    maxid = reassign_ids(item, attr_id, maxid)
  attrs = block.attrs
  if not attrs or len(attrs) == 0:
    return maxid
  idvalue = attrs.get(attr_id)
  if not idvalue:
    return maxid
  maxid += 1
  attrs[attr_id] = maxid
  return maxid


def adjust_cell_row_by(sqref: str, at_row: int, by: int) -> str:
  """
  Adjusts row references in a range.

  Args:
    sqref: Range or list of ranges.
    at_row: Row from which to adjust.
    by: Row increment.

  Returns:
    Adjusted range.
  """
  idx0 = sqref.find(':')
  from_cell = sqref[0:idx0] if idx0 >= 0 else sqref
  to_cell = sqref[idx0 + 1:] if idx0 >= 0 else from_cell
  to_cell_row = get_row_num(to_cell)
  to_cell_num = get_cell_num(to_cell)
  by = by - 1 if by > 0 else 0
  if to_cell_row == at_row:
    to_cell = get_cell_format_position(to_cell_num, to_cell_row + by)
    return from_cell + (':' + to_cell if to_cell != from_cell else '')
  if to_cell_row > at_row:
    from_cell_num = get_cell_num(from_cell)
    from_cell_row = get_row_num(from_cell)
    from_cell = get_cell_format_position(from_cell_num, from_cell_row + by)
    to_cell = get_cell_format_position(to_cell_num, to_cell_row + by)
    return from_cell + (':' + to_cell if to_cell != from_cell else '')
  return sqref
