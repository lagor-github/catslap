# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)
from catslap.base import utils as common
from catslap.utils import zip as zip_util
from catslap.utils import file as file_util


ZIPPED_EXTENSIONS = ['.docx', '.xlsx', '.pptx', '.zip']

class DocumentException(Exception):
  """
  Base exception for document processing.
  """
  pass


class Document:
  """
  Base class for OOXML documents.

  Attributes:
    filename: Input file path.
    value_resolver: Value resolution function.
    repeating_resolver: Repetition resolution function.
    is_zipped: Whether the document is a ZIP package.
    tempdir: Temporary working path.
    default_params: Default parameters.
    default_value_resolver: Default value resolver.
    default_repeat_resolver: Default repeat resolver.
    access_ok_param_list: List of successfully resolved parameters.
    access_err_param_list: List of unresolved parameters.
    test_mode: If True, avoids final writing.
  """
  def __init__(self, filename: str):
    """
    Creates the base document.

    Args:
      filename: Template file name to render.

    Raises:
      FileNotFoundError: If the file does not exist.
      OSError: If copy or extraction fails.
    """
    self.filename = filename
    self.value_resolver = None
    self.repeating_resolver = None
    extension = file_util.get_extension(filename).lower()
    is_dir = file_util.is_directory(filename)
    self.is_zipped = not is_dir and extension in ZIPPED_EXTENSIONS
    if self.is_zipped:
      self.tempdir = zip_util.extract_all(self.filename)
    else:
      self.tempdir = file_util.copy_to_temp_dir(filename)
    self.default_params = {}
    self.default_value_resolver = common.dict_value_resolver(self.default_params)
    self.default_repeat_resolver = common.dict_repeat_resolver(self.default_params)
    self.access_ok_param_list = []
    self.access_err_param_list = []
    self.test_mode = False
    self.config_params = {}

  def set_config_params(self, config: dict):
    self.config_params = config

  def get_document_bytes(self) -> bytes:
    """
    Returns the bytes of the generated document.

    Returns:
      Bytes of the generated document.

    Raises:
      OSError: If temporary read/write fails.
    """
    if self.is_zipped:
      zipfilename = file_util.get_temp_file(None, '.zip')
      try:
        zip_util.zip_directory(zipfilename, self.tempdir)
        return file_util.read_bytes(zipfilename)
      finally:
        file_util.remove_file(zipfilename)
    else:
      return file_util.read_bytes(self.tempdir)

  def close(self):
    """
    Closes the temporary resource (removes temp files/directories).

    Raises:
      OSError: If the temp resource cannot be removed.
    """
    if self.tempdir:
      if self.is_zipped:
        file_util.remove_dir_tree(self.tempdir)
      else:
        file_util.remove_file(self.tempdir)
      self.tempdir = None

  def create_doc_with_resolvers(self, output_file: str, __value_resolver, __repeating_resolver):
    """
    Generates a document using explicit resolvers.

    Args:
      output_file: Output path.
      __value_resolver: Value resolver.
      __repeating_resolver: Repeat resolver.

    Raises:
      OSError: If writing fails.
    """
    data = self.get_bytes_with_resolvers(__value_resolver, __repeating_resolver)
    file_util.write_bytes(output_file, data)

  def get_bytes_with_resolvers(self,  __value_resolver, __repeating_resolver) -> bytes:
    """
    Generates document bytes using explicit resolvers.

    Args:
      __value_resolver: Value resolver.
      __repeating_resolver: Repeat resolver.

    Returns:
      Bytes of the generated document.
    """
    self.value_resolver = __value_resolver
    self.repeating_resolver = __repeating_resolver
    self.process_template(self.tempdir)
    return self.get_document_bytes()

  def test_with_json(self,  json: dict) -> tuple[list, list]:
    """
    Processes the document in test mode using JSON.

    Args:
      json: Value map.

    Returns:
      Tuple with OK and error parameter lists.
    """
    return self.test_with_resolvers(common.dict_value_resolver(json), common.dict_repeat_resolver(json))

  def test_with_resolvers(self,  __value_resolver, __repeating_resolver) -> tuple[list, list]:
    """
    Processes the document in test mode using resolvers.

    Args:
      __value_resolver: Value resolver.
      __repeating_resolver: Repeat resolver.

    Returns:
      Tuple with OK and error parameter lists.
    """
    self.test_mode = True
    self.value_resolver = __value_resolver
    self.repeating_resolver = __repeating_resolver
    self.process_template(self.tempdir)
    return self.access_ok_param_list, self.access_err_param_list

  def create_doc_with_json(self, output_file: str, json: dict):
    """
    Generates a document using JSON.

    Args:
      output_file: Output path.
      json: Value map.

    Raises:
      OSError: If writing fails.
    """
    data = self.get_bytes_with_json(json)
    file_util.write_bytes(output_file, data)

  def get_bytes_with_json(self, json: dict) -> bytes:
    """
    Generates document bytes using JSON.

    Args:
      json: Value map.

    Returns:
      Bytes of the generated document.
    """
    return self.get_bytes_with_resolvers(common.dict_value_resolver(json), common.dict_repeat_resolver(json))

  def resolve_text(self, var_row: int | None, value: str) -> str:
    """
    Resolves {{ }} placeholders inside a text.

    Args:
      var_row: Current row index (if applicable).
      value: Text with placeholders.

    Returns:
      Resolved text.
    """
    rtext = ''
    idx0 = 0
    idx1 = value.find('{{')
    if idx1 < 0:
      return value
    while idx1 >= 0:
      idx2 = value.find('}}', idx1)
      if idx2 <= idx1:
        break
      rtext += value[idx0:idx1]
      param = value[idx1 + 2:idx2]
      resolved = self.resolve_value(var_row, param)
      if resolved is not None:
        rtext = rtext + str(resolved)
      idx0 = idx2 + 2
      idx1 = value.find('{{', idx0)
    return rtext

  def resolve_value(self, row: int | None, param: str) -> any:
    """
    Resolves a parameter using configured resolvers.

    Args:
      row: Current row index.
      param: Expression/key to resolve.

    Returns:
      Resolved value, or '' if not found.
    """
    value = self.default_value_resolver(row, param)
    if value is None:
      value = self.value_resolver(row, param)
    if value is None:
      if param not in self.access_err_param_list:
        self.access_err_param_list.append(param)
      return ''
    if param not in self.access_ok_param_list:
      self.access_ok_param_list.append(param)
    return value

  def resolve_repeating(self, param: str) -> int:
    """
    Resolves the number of repetitions for an expression.

    Args:
      param: Expression to evaluate.

    Returns:
      Repetition count (>=0) or -1 if invalid.
    """
    value = self.default_repeat_resolver(param)
    if value is None or value < 0:
      value = self.repeating_resolver(param)
    if value is None:
      value = 0
    if value < 0:
      if param not in self.access_err_param_list:
        self.access_err_param_list.append(param)
      return -1
    if param not in self.access_ok_param_list:
      self.access_ok_param_list.append(param)
    return value

  def process_template(self, tempdir: str):
    """
    Processes the template (abstract method).

    Args:
      tempdir: Temporary working directory.
    """
    pass
