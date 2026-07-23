# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026


class DotDict(dict):
  """
  Dictionary with attribute access (dot notation).
  """
  __getattr__ = dict.get
  __setattr__ = dict.__setitem__
  __delattr__ = dict.__delitem__

  @staticmethod
  def _convert_value(value):
    if isinstance(value, dict):
      return DotDict.create(value)
    if isinstance(value, list):
      return [DotDict._convert_value(item) for item in value]
    if isinstance(value, tuple):
      return tuple(DotDict._convert_value(item) for item in value)
    return value

  @staticmethod
  def create(value_map):
    """
    Recursively converts dictionaries to DotDict.

    Args:
      value_map: Base dictionary.

    Returns:
      DotDict with converted sub-dictionaries.

    Raises:
      AttributeError: If non-existent attributes are accessed.
      TypeError: If the input is not a dict with iterable elements.
    """
    for key in list(value_map.keys()):
      value_map[key] = DotDict._convert_value(value_map[key])
    return DotDict(value_map)
