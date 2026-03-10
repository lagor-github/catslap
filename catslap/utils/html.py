# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026


from catslap.utils import encoding as enc_util
from catslap.utils import text as text_util
from catslap.utils import css_color


class HtmlException(Exception):
  """
  HTML processing exception.
  """
  pass


def extract_image_data(image: str) -> tuple[str, str, bytes]:
  """
  Extracts metadata and bytes from a data URI embedded image.

  Args:
    image: Data URI string (data:<mime>;base64,<data>).

  Returns:
    Tuple (media_type, encoding, bytes) for the image.

  Raises:
    HtmlException: If the format, media type, or base64 are invalid.
  """
  idx0 = 5
  idx = image.find(';', idx0)
  if idx < 0:
    raise HtmlException('Invalid image media-type: ' + image[0:20] + '...')
  media_type = image[idx0:idx]
  idx2 = media_type.find('/')
  if idx2 < 0:
    raise HtmlException('Unsuppported image media-type: ' + media_type)
  idx0 = idx + 1
  idx = image.find(',', idx0)
  if idx < 0:
    raise HtmlException('Invalid image data encoding: ' + image[0:20] + '...')
  encoding = image[idx0:idx]
  if encoding.lower() != 'base64':
    raise HtmlException('Unsupported image data encoding: ' + encoding)
  image = image[idx+1:]
  try:
    base64bytes = enc_util.from_base64(image)
  except Exception:
    raise HtmlException('Invalid base64 image data: ' + image[0:20] + '...')
  return media_type, encoding, base64bytes


def parse_css(css: str) -> dict:
  """
  Parses an inline CSS block into a dictionary.

  Args:
    css: Style string (e.g., 'color: red; font-size: 12px;').

  Returns:
    Dictionary of properties normalized to lowercase.
  """
  stylemap = {}
  if css:
    lines = css.split(';')
    for line in lines:
      idx = line.find(':')
      if idx > 0:
        key = text_util.trim(line[0:idx]).lower()
        value = text_util.trim(line[idx+1:]).lower()
        stylemap[key] = value
  return stylemap

def get_rgb_color(color: str|None) -> str|None:
  return css_color.get_rgb_color(color)


