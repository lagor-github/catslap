# -*- coding: utf-8 -*-
"""Catslap HTML templates: expressions, scoped loops, conditions and assignments."""

import re
from collections.abc import Iterable
from html import escape

from catslap.base.document import Document, DocumentException
from catslap.utils import file as file_util


class HtmlDocument(Document):
  def __init__(self, file: str):
    super().__init__(file)
    # Rich HTML stays raw by default; scalar metadata can be escaped explicitly.
    self.default_params['escape_html'] = lambda value: escape(str(value) if value is not None else '', quote=True)

  def process_template(self, tmpfile: str):
    content = file_util.read_bytes(self.filename).decode('utf-8-sig')
    tokens = re.split(r'(\{\{.*?\}\}|\{%.*?%\})', content, flags=re.DOTALL)
    nodes, position, stop = self.__parse_nodes(tokens, 0, set())
    if stop is not None or position != len(tokens):
      raise DocumentException('Unexpected closing directive')
    output = self.__render_nodes(nodes, None)
    if not self.test_mode:
      file_util.write_bytes(tmpfile, output.encode('utf-8'))

  def __parse_nodes(self, tokens, position, stops):
    nodes = []
    while position < len(tokens):
      token = tokens[position]
      position += 1
      if token.startswith('{{'):
        if not token.endswith('}}'):
          raise DocumentException('Unclosed HTML template expression')
        expression = token[2:-2].strip()
        if not expression:
          raise DocumentException('Empty HTML expression')
        nodes.append(('value', expression))
        continue
      if not token.startswith('{%'):
        if '{{' in token or '{%' in token:
          raise DocumentException('Unclosed HTML template expression or directive')
        nodes.append(('text', token))
        continue
      directive = token[2:-2].strip()
      keyword = directive.split(None, 1)[0] if directive else ''
      if keyword in stops:
        return nodes, position, directive
      if keyword == 'for':
        match = re.fullmatch(r'for\s+([A-Za-z_]\w*)\s+in\s+(.+)', directive, re.DOTALL)
        if not match:
          raise DocumentException('Expected: {% for name in expression %}')
        body, position, closing = self.__parse_nodes(tokens, position, {'endfor'})
        if closing != 'endfor':
          raise DocumentException("Missing '{% endfor %}'")
        nodes.append(('for', match.group(1), match.group(2), body))
      elif keyword == 'if':
        expression = directive[2:].strip()
        if not expression:
          raise DocumentException('Expected a condition after if')
        branches = []
        while True:
          body, position, closing = self.__parse_nodes(tokens, position, {'elif', 'else', 'endif'})
          branches.append((expression, body))
          if closing is None:
            raise DocumentException("Missing '{% endif %}'")
          if closing.startswith('elif '):
            expression = closing[5:].strip()
            continue
          if closing == 'else':
            body, position, closing = self.__parse_nodes(tokens, position, {'endif'})
            branches.append((None, body))
          if closing != 'endif':
            raise DocumentException("Expected '{% endif %}'")
          break
        nodes.append(('if', branches))
      elif keyword == 'set':
        match = re.fullmatch(r'set\s+([A-Za-z_]\w*)\s*=\s*(.+)', directive, re.DOTALL)
        if not match:
          raise DocumentException('Expected: {% set name = expression %}')
        nodes.append(('set', match.group(1), match.group(2)))
      else:
        raise DocumentException("Unexpected HTML directive: {% " + directive + " %}")
    return nodes, position, None

  def __render_scoped(self, nodes, row, variables=None):
    previous = dict(self.default_params)
    try:
      self.default_params.update(variables or {})
      return self.__render_nodes(nodes, row)
    finally:
      self.default_params.clear()
      self.default_params.update(previous)

  def __render_nodes(self, nodes, row):
    output = []
    for node in nodes:
      kind = node[0]
      if kind == 'text':
        output.append(node[1])
      elif kind == 'value':
        value = self.resolve_value(row, node[1])
        output.append('' if value is None else str(value))
      elif kind == 'set':
        self.default_params[node[1]] = self.resolve_value(row, node[2])
      elif kind == 'if':
        for condition, body in node[1]:
          if condition is None or bool(self.resolve_value(row, condition)):
            output.append(self.__render_scoped(body, row))
            break
      elif kind == 'for':
        values = self.resolve_value(row, node[2])
        if self.test_mode:
          values = [{}]
        elif values is None or values == '':
          values = []
        elif isinstance(values, (str, bytes)) or not isinstance(values, Iterable):
          raise DocumentException('HTML for expression must return an iterable: ' + node[2])
        values = list(values)
        total = 1 if self.test_mode else self.default_repeat_resolver(node[2])
        if not self.test_mode and (total is None or total < 0):
          total = self.repeating_resolver(node[2])
        if total is None or total < 0:
          # Generic iterables (e.g. dict keys) are supported even when a legacy
          # list-only resolver cannot determine their length.
          total = len(values)
        parent_loop = self.default_params.get('loop')
        for index in range(total):
          item = values[index] if index < len(values) else {}
          if isinstance(item, dict):
            item = dict(item)
            item['row'] = index + 1
          loop = {'index': index + 1, 'index0': index, 'first': index == 0,
                  'last': index == total - 1, 'length': total, 'parent': parent_loop}
          output.append(self.__render_scoped(node[3], index, {node[1]: item, 'loop': loop}))
    return ''.join(output)

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
