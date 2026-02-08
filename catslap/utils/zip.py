# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

import os
import zipfile

from catslap.utils import file as file_util


def extract_all(zipfilename: str) -> str:
  """
  Extracts a ZIP to a temporary directory.

  Args:
    zipfilename: ZIP file path.

  Returns:
    Temporary directory path with extracted contents.

  Raises:
    FileNotFoundError: If the ZIP does not exist.
    zipfile.BadZipFile: If the ZIP is invalid.
    OSError: If the filesystem cannot be written.
  """
  tempdir = file_util.get_temp_dir()
  with zipfile.ZipFile(zipfilename, 'r') as zip_ref:
    zip_ref.extractall(tempdir)
  return tempdir


def zip_directory(directory: str, path: str):
  """
  Compresses a directory into a ZIP file.

  Args:
    directory: Output ZIP file path.
    path: Directory path to compress.

  Raises:
    OSError: If a read/write error occurs.
    zipfile.BadZipFile: If the ZIP cannot be created.
  """
  zipf = zipfile.ZipFile(directory, 'w', zipfile.ZIP_DEFLATED)
  with zipf:
    __zip_directory(path, path, zipf)


def __zip_directory(path, relpath, zipf):
  for root, _, files in os.walk(path):
    for file in files:
      filename = os.path.normpath(os.path.join(root, file))
      if os.path.isdir(filename):
        __zip_directory(filename, relpath, zipf)
      if file in ['.DS_Store']:
        os.remove(filename)
        continue
      arcname = filename[len(relpath) + 1:]
      zipf.write(filename, arcname)
