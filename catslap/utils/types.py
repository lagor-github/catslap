# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

from decimal import Decimal


def length(obj: any):
  """
  Returns the length of an object or 0 if None.

  Args:
    obj: Object with __len__.

  Returns:
    Length or 0.

  Raises:
    TypeError: If the object is not len-able.
  """
  if obj is None:
    return 0
  return len(obj)


def to_bool(string: str) -> bool:
  """
  Converts a value to boolean by comparing with 'true'.

  Args:
    string: Value to convert.

  Returns:
    True if the value is 'true' (case-insensitive), False otherwise.
  """
  if string is None:
    return False
  try:
    return str(string).lower() == 'true'
  except ValueError:
    return False


def to_int(obj: any, default_value: int | None = None) -> int:
  """
  Converts a value to int with fallback.

  Args:
    obj: Value to convert.
    default_value: Default value if conversion fails.

  Returns:
    Resulting int or 0.
  """
  if isinstance(obj, int):
    return int(obj)
  if isinstance(obj, float):
    return int(round(float(obj)))
  if isinstance(obj, Decimal):
    return int(obj)
  if isinstance(obj, str):
    try:
      return int(obj)
    except ValueError:
      pass
  if default_value is not None:
    return to_int(default_value, 0)
  return 0


def to_float(obj: any, default_value: int | None = None) -> float:
  """
  Converts a value to float with fallback.

  Args:
    obj: Value to convert.
    default_value: Default value if conversion fails.

  Returns:
    Resulting float or 0.0.
  """
  if isinstance(obj, int):
    return float(int(obj))
  if isinstance(obj, float):
    return float(obj)
  if isinstance(obj, Decimal):
    return float(obj)
  if isinstance(obj, str):
    try:
      return float(obj)
    except ValueError:
      pass
  if default_value is not None:
    return to_float(default_value, 0.0)
  return 0.0


def merge_list_unique(list1: list, list2: list) -> list:
  """
  Merges lists avoiding duplicates (mutates list1).

  Args:
    list1: Base list.
    list2: List to add.

  Returns:
    Combined list without duplicates.
  """
  if list1 is None:
    list1 = []
  if list2 is None:
    list2 = []
  for item in list2:
    if item not in list1:
      list1.append(item)
  return list1


def merge_dicts(dict1: dict, dict2: dict | None) -> dict:
  """
  Merges dictionaries without overwriting existing keys.

  Args:
    dict1: Base dictionary.
    dict2: Dictionary to merge.

  Returns:
    Merged dictionary.
  """
  dict3 = {}
  if dict1:
    for item, value in dict1.items():
      dict3[item] = value
  if dict2 is not None:
    for item, value in dict2.items():
      if dict3.get(item) is None:
        dict3[item] = value
  return dict3
