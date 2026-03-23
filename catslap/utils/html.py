# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026


import ssl
import urllib.request

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


def fetch_blob_image(image: str) -> tuple[str, bytes]:
  """
  Fetches image bytes from a blob URL (blob:https://host/uuid).

  Strips the 'blob:' prefix and performs an HTTP/HTTPS GET request.
  For HTTPS requests, certificate verification is disabled to allow
  self-signed certificates on local servers.

  Args:
    image: blob URL string (blob:https://...).

  Returns:
    Tuple (media_type, bytes) for the image.

  Raises:
    HtmlException: If the URL cannot be fetched or returns no data.
  """
  url = image[5:]  # strip 'blob:' prefix
  try:
    if url.startswith('https://'):
      ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
      ctx.check_hostname = False
      ctx.verify_mode = ssl.CERT_NONE
      response = urllib.request.urlopen(url, context=ctx)
    else:
      response = urllib.request.urlopen(url)
    with response:
      content_type = response.headers.get('Content-Type', 'image/png')
      mediatype = content_type.split(';')[0].strip()
      data = response.read()
  except HtmlException:
    raise
  except Exception as e:
    raise HtmlException(f'Cannot fetch blob image from {url}: {e}')
  if not data:
    raise HtmlException(f'Empty response fetching blob image from {url}')
  return mediatype, data


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


