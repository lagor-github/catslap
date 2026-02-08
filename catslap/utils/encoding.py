# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

import base64


UTF_8 = "UTF-8"
ISO_8859_1 = "iso-8859-1"


def from_base64(string: str) -> bytes:
  """
  Decodes a Base64 string to bytes.

  Args:
    string: Base64 string (ISO-8859-1 encoded).

  Returns:
    Decoded bytes.

  Raises:
    binascii.Error: If the Base64 string is invalid.
    ValueError: If input conversion fails.
  """
  return base64.standard_b64decode(bytes(string, ISO_8859_1))
