# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)
from catslap.utils.dotdict import DotDict


def complete_pathfile(pathfile: str, file: str) -> str:
  """
  Ensures a path ends with a file name.

  Args:
    pathfile: Base path.
    file: File name.

  Returns:
    Full path including the file.
  """
  if not pathfile.endswith(file):
    if not pathfile.endswith('/'):
      pathfile = pathfile + '/'
    pathfile = pathfile + file
  return pathfile


def resolve_param_value(value_map: dict, row: int | None, param: str) -> any:
  """
  Resolves an expression against a value map.

  Args:
    value_map: Dictionary of available values.
    row: Row index (assigned to 'row' in the map).
    param: Expression to evaluate.

  Returns:
    Evaluated value, or None if it cannot be resolved.

  Raises:
    SyntaxError: If the expression is invalid.
  """
  if not param:
    return None
  # -- define el parámetro row en las variable locales
  if (row is not None and value_map):
    value_map['row'] = row
  try:
    param = param.replace('‘', '\'').replace('’', '\'').replace('“', '\"').replace('”', '\"')
    value = eval(param, {"__builtins__": None}, DotDict.create(value_map))
  except AttributeError:
    value = None
  except NameError:
    value = None
  except TypeError:
    value = None
  return value


def resolve_param_repeating(value_map: dict, param: str) -> any:
  """
  Computes the repetition count if a value is a list.

  Args:
    value_map: Dictionary of available values.
    param: Expression to evaluate.

  Returns:
    List length, or -1 if the value is not a list.
  """
  value = resolve_param_value(value_map, 0, param)
  if isinstance(value, list):
    return len(value)
  return -1


def dict_value_resolver(value_map):
  """
  Creates a value resolver for a map.

  Args:
    value_map: Dictionary of available values.

  Returns:
    Resolver function (row, param) -> value.
  """
  def __value_resolver(var_row: int | None, vparam: str) -> any:
    return resolve_param_value(value_map, var_row, vparam)
  return __value_resolver


def dict_repeat_resolver(value_map):
  """
  Creates a repetition resolver for a map.

  Args:
    value_map: Dictionary of available values.

  Returns:
    Resolver function (param) -> int.
  """
  def __repeat_resolver(param: str) -> int:
    return resolve_param_repeating(value_map, param)
  return __repeat_resolver
