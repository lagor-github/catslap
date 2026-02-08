# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

import os
import shutil
import tempfile
from pathlib import Path


def get_temp_dir():
  """
  Creates a temporary directory.

  Returns:
    Path to the created temporary directory.

  Raises:
    OSError: If the temporary directory cannot be created.
  """
  return tempfile.mkdtemp()


def get_pathname(path: str):
  """
  Gets the directory portion (with trailing slash) of a path.

  Args:
    path: File or directory path.

  Returns:
    Directory portion ending with '/' or an empty string if not present.
  """
  idx = path.rfind('/')
  if idx > 0:
    path = path[0:idx + 1]
    return path
  return ''


def get_filename(path: str):
  """
  Gets the filename without its directory.

  Args:
    path: Full path.

  Returns:
    Filename.
  """
  idx = path.rfind('/')
  if idx >= 0:
    return path[idx + 1:]
  return path


def get_extension(path: str):
  """
  Gets a file extension (including the dot).

  Args:
    path: File path.

  Returns:
    Extension including the dot or an empty string if missing.
  """
  idx = path.rfind('.')
  if idx >= 0:
    return path[idx:]
  return ''


def exist(file):
  """
  Indicates whether a path exists as a file or directory.

  Args:
    file: Path to check.

  Returns:
    True if it exists as a file or directory.
  """
  return is_file(file) or is_directory(file)

def is_directory(file):
  """
  Indicates whether a path is a directory.

  Args:
    file: Path to check.

  Returns:
    True if it is a directory.
  """
  return os.path.isdir(file)

def is_file(file):
  """
  Indicates whether a path is a file.

  Args:
    file: Path to check.

  Returns:
    True if it is a file.
  """
  return os.path.isfile(file)

def get_base_dir(file):
  """
  Gets the base directory of the current module.

  Args:
    file: Ignored parameter (kept for compatibility).

  Returns:
    Directory path of the current file.
  """
  return Path(__file__).resolve().parent.as_posix()

def complete_path(path: str, relpath: str) -> str:
  """
  Resolves a relative path against a base path.

  Args:
    path: Base path.
    relpath: Relative path (supports ./ and ../).

  Returns:
    Absolute/combined path.

  Raises:
    RecursionError: If the relative path triggers infinite recursion (edge case).
  """
  if not path.endswith('/'):
    path = path + '/'
  if relpath.startswith('/'):
    return path + complete_path('/', relpath[1:])[1:]
  if relpath.startswith('./'):
    return complete_path(path, relpath[2:])
  if relpath.startswith('../'):
    if path.endswith('/'):
      path = path[0:len(path)-1]
    path = get_pathname(path)
    return complete_path(path, relpath[3:])
  idx = relpath.find('/')
  if idx > 0:
    return complete_path(path + relpath[0:idx + 1], relpath[idx + 1:])
  return path + relpath


def get_temp_file(tmpdir: str | None, extension: str) -> str:
  """
  Creates a temporary file and returns its path.

  Args:
    tmpdir: Temporary directory (None to use system default).
    extension: Temporary file suffix.

  Returns:
    Path to the created temporary file.

  Raises:
    OSError: If the file cannot be created.
  """
  (handle, name) = tempfile.mkstemp(suffix=extension, prefix=None, dir=tmpdir, text=False)
  os.close(handle)
  return name


def read_bytes(file: str) -> bytes:
  """
  Reads the binary contents of a file.

  Args:
    file: File path.

  Returns:
    Read bytes.

  Raises:
    FileNotFoundError: If the file does not exist.
    OSError: If read errors occur.
  """
  handler = open(file, "rb")
  with handler:
    return handler.read()


def read_text(file: str, encoding: str) -> str:
  """
  Reads a text file with a specific encoding.

  Args:
    file: File path.
    encoding: Text encoding.

  Returns:
    Read text.

  Raises:
    FileNotFoundError: If the file does not exist.
    UnicodeDecodeError: If decoding fails.
    OSError: If read errors occur.
  """
  data = read_bytes(file)
  return data.decode(encoding)


def write_bytes(file: str, data: bytes):
  """
  Writes bytes to a file.

  Args:
    file: File path.
    data: Binary content to write.

  Raises:
    OSError: If write errors occur.
  """
  infile_handler = open(file, "wb")
  with infile_handler:
    infile_handler.write(data)


def remove_dir_tree(directory: str):
  """
  Removes a directory recursively.

  Args:
    directory: Directory path.

  Raises:
    FileNotFoundError: If the directory does not exist.
    OSError: If it cannot be removed.
  """
  shutil.rmtree(directory)


def remove_file(filename: str):
  """
  Removes a file if it exists.

  Args:
    filename: File path.

  Raises:
    OSError: If it cannot be removed.
  """
  if os.path.exists(filename):
    os.remove(filename)


def list_files(directory: str) -> list | None:
  """
  Lists files and directories in a directory.

  Args:
    directory: Directory path.

  Returns:
    List of names or None if it does not exist.

  Raises:
    OSError: If a directory read error occurs.
  """
  try:
    return os.listdir(directory)
  except FileNotFoundError:
    return None


def copy_to_temp_dir(filename: str) -> str:
  """
  Copies a file or directory to a temporary directory.

  Args:
    filename: File or directory path to copy.

  Returns:
    Destination path inside the temporary directory.

  Raises:
    FileNotFoundError: If the source does not exist.
    OSError: If copy errors occur.
  """
  temp_dir = get_temp_dir()
  if os.path.isdir(filename):
    name = os.path.basename(os.path.normpath(filename))
    dest = os.path.join(temp_dir, name)
    shutil.copytree(name, dest)
    return dest
  dest = os.path.join(temp_dir, os.path.basename(filename))
  shutil.copy2(filename, dest)
  return dest
