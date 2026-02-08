# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)
from catslap.base.document import Document, DocumentException
from catslap.utils import encoding as enc_util
from catslap.utils import file as file_util
from catslap.utils import text as text_util
from catslap.utils.sentence import Sentence


class HtmlDocument(Document):
  """
  Processes HTML templates with placeholders and directives.
  """
  def __init__(self, file: str):
    super().__init__(file)

  def process_template(self, tmpfile: str):
    """
    Processes the HTML template.

    Args:
      tmpfile: Temporary output file.
    """
    self.__process_html_file(tmpfile)

  def __process_html_file(self, tmpfile: str):
    html_content = text_util.trim(str(file_util.read_bytes(self.filename), enc_util.UTF_8))
    out = self.__process_html_fragment(html_content, None)
    if not self.test_mode:
      file_util.write_bytes(tmpfile, bytes(out, enc_util.UTF_8))

  def __process_html_fragment(self, html_content: str, row: int | None) -> str:
    out = ''
    sentence = Sentence(html_content)
    while not sentence.is_eos():
      if sentence.match('{{'):
        keyword = sentence.parse_until_word('}}')
        if not keyword:
          raise DocumentException('"}}" expression expected')
        keyword = keyword.strip()
        value = str(self.resolve_value(row, keyword))
        out += value
        continue
      if sentence.match('{%'):
        keyword = sentence.parse_until_word('%}')
        if not keyword:
          raise DocumentException('"%}" expression expected')
        keyword = keyword.strip()
        if keyword.startswith("for "):
          keyword = keyword[4:]
          idx2 = keyword.find(' in ')
          if idx2 < 0:
            raise DocumentException("'in' clause expected in 'for' expression")
          varname1 = text_util.trim(keyword[0:idx2])
          varname2 = text_util.trim(keyword[idx2 + 4:])
          idx0 = sentence.idx;
          idx3 = HtmlDocument.find_end_for(sentence)
          if idx3 < 0:
            raise DocumentException("'{% endfor %}' clause expected")
          sub_content = sentence.substring(idx0, idx3)
          total = self.resolve_repeating(varname2)
          if self.test_mode:
            total = 1
          value = self.resolve_value(None, varname2)
          pos = 1
          value2 = {}
          for row2 in range(0, total):
            if not self.test_mode:
              value2 = value[row2] if isinstance(value, list) and row2 < len(value) else {}
              self.default_params[varname1] = value2
              value2['row'] = pos
              pos +=1
            out += self.__process_html_fragment(sub_content, row2)
          value2['row'] = pos
          continue
        if keyword.startswith("if "):
          varname1 = text_util.trim(keyword[3:])
          idx0 = sentence.idx
          idx3 = HtmlDocument.find_end_if(sentence)
          if idx3 < 0:
            raise DocumentException("'{% endif %}' clause expected")
          value = self.resolve_value(row, varname1)
          if (isinstance(value, str) and not text_util.is_empty(value)) \
          or (isinstance(value, bool) and value is True) \
          or (isinstance(value, int) and value != 0) \
          or (isinstance(value, list) and len(value) > 0) \
          or (isinstance(value, dict) and value is not None) \
          or (isinstance(value, float) and value != 0):
            sub_content = sentence.substring(idx0, idx3)
            out += self.__process_html_fragment(sub_content, row)
          continue
        raise DocumentException("Clause '{% " + keyword + " %}' unexcepted")
      out += sentence.peek_next()
    return out

  @staticmethod
  def find_end_for(sentence) -> int:
    """
    Finds the end of a for block.

    Args:
      sentence: Sentence parser.

    Returns:
      Index of the '{% endfor %}' start or end of text.
    """
    nfors = 1
    while not sentence.is_eos() and nfors > 0:
      if sentence.match('{%'):
        idx3 = sentence.idx - 2;
        keyword = sentence.parse_until_word('%}')
        if not keyword:
          raise DocumentException('"%}" expression expected')
        keyword = keyword.strip()
        pos = keyword.find(' ')
        keyword = keyword[0:pos] if pos >= 0 else keyword
        if keyword == 'endfor':
          nfors -= 1
          if nfors == 0:
            return idx3
        if keyword == 'for':
          nfors += 1
      else:
        sentence.next()
    return sentence.idx

  @staticmethod
  def find_end_if(sentence) -> int:
    """
    Finds the end of an if block.

    Args:
      sentence: Sentence parser.

    Returns:
      Index of the '{% endif %}' start or end of text.
    """
    nifs = 1
    while not sentence.is_eos() and nifs > 0:
      if sentence.match('{%'):
        idx3 = sentence.idx - 2;
        keyword = sentence.parse_until_word('%}')
        if not keyword:
          raise DocumentException('"%}" expression expected')
        keyword = keyword.strip()
        pos = keyword.find(' ')
        keyword = keyword[0:pos] if pos >= 0 else keyword
        if keyword == 'endif':
          nifs -= 1
          if nifs == 0:
            return idx3
        if keyword == 'if':
          nifs += 1
      else:
        sentence.next()
    return sentence.idx
