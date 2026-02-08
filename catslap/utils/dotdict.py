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
    for key in value_map.keys():
      vkey = value_map[key]
      if isinstance(vkey, dict):
        value_map[key] = DotDict.create(vkey)
    return DotDict(value_map)
