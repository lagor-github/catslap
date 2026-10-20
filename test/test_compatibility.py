"""Regression checks for templates and resolver APIs used by consumer projects."""
import copy
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from catslap.base.utils import dict_repeat_resolver, dict_value_resolver
from catslap.catslap import Catslap
from catslap.docx.styles import Styles
from catslap.html.document import HtmlDocument


class CompatibilityTests(unittest.TestCase):
  def setUp(self):
    self.directory = tempfile.TemporaryDirectory()
    self.addCleanup(self.directory.cleanup)
    self.root = Path(self.directory.name)

  def document(self, template):
    source = self.root / 'template.html'
    source.write_text(template, encoding='utf-8')
    document = HtmlDocument(str(source))
    self.addCleanup(document.close)
    return document

  def test_numbered_colors_and_list_styles(self):
    source = self.root / 'styles.xml'
    source.write_text('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    styles = Styles(str(source))
    styles.find_styles()
    styles.setStyle('table_cell_bgcolor2', '#E8E8E8')
    styles.setStyle('table_cell_color2', '#123456')
    styles.setStyle('heading', 'Heading 1')
    styles.setStyle('heading2', 'Custom heading')
    styles.setStyle('list_bullet', 'Bullet 1')
    styles.setStyle('list_bullet2', 'Custom bullet')
    self.assertEqual(styles.style_map['table_cell_bgcolor2'], 'E8E8E8')
    self.assertEqual(styles.style_map['table_cell_color2'], '123456')
    self.assertEqual(styles.style_map['heading'][1], 'Customheading')
    self.assertEqual(styles.style_map['list_bullet'][1], 'Custombullet')

  def test_explicit_resolvers_combine_nested_scope_and_external_data(self):
    document = self.document('{% set prefix = "[" %}{% for entry in entries %}'
      '{% for child in entry.children %}{{ prefix + entry.name + child + suffix }}'
      '{% endfor %}{% endfor %}')
    data = {'entries': [{'name': 'A', 'children': ['x', 'y']}], 'suffix': '!'}
    original = copy.deepcopy(data)
    result = document.get_bytes_with_resolvers(dict_value_resolver(data), dict_repeat_resolver(data))
    self.assertEqual(result, b'[Ax![Ay!')
    self.assertEqual(data, original)

  def test_plain_custom_resolvers_honor_repeat_count(self):
    for count, expected in [(0, b''), (1, b'1:A;'), (2, b'1:A;2:B;')]:
      with self.subTest(count=count):
        document = self.document('{% for entry in entries %}{{ loop.index }}:{{ entry.name }};{% endfor %}')
        calls = []
        def repeat(expression):
          calls.append(expression)
          return count
        result = document.get_bytes_with_resolvers(
          lambda row, expression: [{'name': 'A'}, {'name': 'B'}] if expression == 'entries' else None,
          repeat)
        self.assertEqual(result, expected)
        self.assertEqual(calls, ['entries'])

  def test_generic_iterables_work_with_list_only_repeat_resolver(self):
    document = self.document('{% for key in values.keys() %}{{ key }};{% endfor %}')
    data = {'values': {'A': 1, 'B': 2}}
    self.assertEqual(document.get_bytes_with_resolvers(
      dict_value_resolver(data), dict_repeat_resolver(data)), b'A;B;')

  def test_html_diagnostic_visits_missing_loop_fields(self):
    document = self.document('{% for entry in entries %}{{ entry.name }}{% endfor %}')
    _, errors = document.test_with_json({})
    self.assertIn('entries', errors)
    self.assertIn('entry.name', errors)

  def test_example_documents_render(self):
    import json
    root = Path(__file__).resolve().parent
    for extension in ('docx', 'xlsx'):
      with self.subTest(extension=extension):
        data = json.loads((root / 'data/Catslap_sales.json').read_text(encoding='utf-8'))
        result = Catslap(data).process(str(root / f'templates/{extension}/Catslap_example.{extension}'))
        with ZipFile(BytesIO(result)) as archive:
          self.assertIsNone(archive.testzip())
          part = 'word/document.xml' if extension == 'docx' else 'xl/workbook.xml'
          self.assertIn(part, archive.namelist())


if __name__ == '__main__':
  unittest.main()
