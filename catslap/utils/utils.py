# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026


def repeat(char: str, repeating: int) -> str:
  """
  Repeats a character N times.

  Args:
    char: Character to repeat.
    repeating: Number of repetitions.

  Returns:
    Resulting string.
  """
  string = ''
  if repeating > 0:
    for _ in range(0, int(repeating)):
      string = string + char
  return string


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


def split_no_empty(string: str, sep: str) -> list:
  """
  Splits a string and removes empty elements.

  Args:
    string: String to split.
    sep: Separator.

  Returns:
    List of non-empty elements, trimmed.
  """
  slist = string.split(sep)
  olist = []
  for item in slist:
    if item == '':
      continue
    olist.append(item.strip())
  return olist


def merge_dicts(dict1: dict, dict2: dict | None) -> dict:
  """
  Merges two dictionaries without overwriting existing keys.

  Args:
    dict1: Base dictionary.
    dict2: Dictionary to merge (adds only missing keys).

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
