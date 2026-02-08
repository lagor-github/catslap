# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026




def count_lf(text: str) -> int:
  """
  Counts line breaks (\\n) in a text.

  Args:
    text: Input text.

  Returns:
    Number of '\\n' characters.
  """
  return text.count('\n')


def is_alpha(text: str) -> bool:
  """
  Checks whether a text contains only A-Z letters.

  Args:
    text: Text to evaluate.

  Returns:
    True if the trimmed text is alphabetic, False otherwise.
  """
  text = trim(text)
  if is_empty(text):
    return False
  for char in text:
    if (char < 'a' or char > 'z') and (char < 'A' or char > 'Z'):
      return False
  return True


def is_numeric(text: any) -> bool:
  """
  Checks whether a text represents an integer/decimal number.

  Args:
    text: Text or number.

  Returns:
    True if numeric (supports sign and decimal point).
  """
  if isinstance(text, int) or isinstance(text, float):
    return True
  text = trim(text)
  if is_empty(text):
    return False
  idx = 0
  if text[idx] == '+' or text[idx] == '-':
    idx += 1
  dot = False
  while idx < len(text):
    char = text[idx]
    if not dot and char == '.':
      dot = True
      idx += 1
      continue
    if char < '0' or char > '9':
      return False
    idx += 1
  return True


def is_hex(text: str) -> bool:
  """
  Checks whether a text is hexadecimal.

  Args:
    text: Text to evaluate.

  Returns:
    True if it contains only hexadecimal characters.
  """
  text = trim(text)
  if is_empty(text):
    return False
  for char in text:
    if (char < '0' or char > '9') and (char < 'a' or char > 'f') and (char < 'A' or char > 'F'):
      return False
  return True


def is_decimal(text: str) -> bool:
  """
  Checks whether a text is a valid decimal.

  Args:
    text: Text to evaluate.

  Returns:
    True if it has decimal format (supports sign).
  """
  text = trim(text)
  if is_empty(text):
    return False
  idx = 0
  if text[idx] == '-' or text[idx] == '+':
    idx += 1
  dot = False
  while idx < len(text):
    char = text[idx]
    if char == '.' and not dot:
      dot = True
      idx += 1
      continue
    if char < '0' or char > '9':
      return False
    idx += 1
  return True


def trim(string: str|None):
  """
  Trims spaces and line breaks from both ends.

  Args:
    string: Text to clean.

  Returns:
    Text without leading/trailing whitespace.
  """
  if string is None or not isinstance(string, str) or len(string) == 0:
    return ''
  if not isinstance(string, str):
    string = str(string)
  if string[0] == '\uFEFF':
    string = string[1:]
  return string.strip(' \t\r\n')


def ltrim(string: str|None):
  """
  Trims spaces/line breaks from the start.

  Args:
    string: Text to clean.

  Returns:
    Text without leading whitespace.
  """
  if string is None or not isinstance(string, str) or len(string) == 0:
    return ''
  if not isinstance(string, str):
    string = str(string)
  if string[0] == '\uFEFF':
    string = string[1:]
  return string.lstrip(' \t\r\n')


def rtrim(string: str|None):
  """
  Trims spaces/line breaks from the end.

  Args:
    string: Text to clean.

  Returns:
    Text without trailing whitespace.
  """
  if string is None or not isinstance(string, str) or len(string) == 0:
    return ''
  if not isinstance(string, str):
    string = str(string)
  if string[0] == '\uFEFF':
    string = string[1:]
  return string.rstrip(' \t\r\n')


def remove_quotes(text: str|None) -> str:
  """
  Removes wrapping quotes if present.

  Args:
    text: Text to process.

  Returns:
    Text without outer quotes.
  """
  if not text:
    return ''
  text = text.strip()
  if (text.startswith('"') and text.endswith('"')) \
  or (text.startswith("'") and text.endswith("'")) \
  or (text.startswith("`") and text.endswith("`")) \
  or (text.startswith("´") and text.endswith("´")) \
  or (text.startswith('“') and text.endswith('”')) \
  or (text.startswith("‘") and text.endswith("’")):
    text = text[1:-1]
  return text


def is_empty(value: any) -> bool:
  """
  Checks whether a value is empty (None or empty string after trim).

  Args:
    value: Value to evaluate.

  Returns:
    True if empty.
  """
  return value is None or trim(str(value)) == ''


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


def split_no_empty(string: str, sep: str) -> list:
  """
  Splits a string and removes empty elements.

  Args:
    string: String to split.
    sep: Separator.

  Returns:
    List of non-empty elements with trim applied.
  """
  slist = string.split(sep)
  olist = []
  for item in slist:
    if item == '':
      continue
    olist.append(item.strip())
  return olist
