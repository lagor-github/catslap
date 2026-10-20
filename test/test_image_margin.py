"""Check image margins from template directives through generated Word XML."""
import base64
import os
import subprocess
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from PIL import Image, ImageChops
from catslap.catslap import Catslap
from catslap.docx import elements
from catslap.docx.styles import Styles
from catslap.base.relationships import Relationships
from catslap.base.types import ContentTypes
from catslap.utils.xml import XmlTag, XmlText

# Match AI Document Miller's compatibility initialization for detached nodes.
for element_type, probe in ((XmlTag, XmlTag('root')), (XmlText, XmlText(''))):
    if not hasattr(probe, 'parent'):
        element_type.parent = None

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'


class ImageMarginTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.styles_xml = f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
        styles_path = self.root / 'styles.xml'
        styles_path.write_text(self.styles_xml)
        self.styles = Styles(str(styles_path))
        self.styles.find_styles()
        stream = BytesIO()
        Image.new('RGB', (2, 2), 'red').save(stream, format='PNG')
        self.image = stream.getvalue()
        self.src = 'data:image/png;base64,' + base64.b64encode(self.image).decode()

    def render(self, margin=None, content=None):
        directive = '' if margin is None else '<w:p><w:r><w:t>{% style image_margin=' + margin + ' %}</w:t></w:r></w:p>'
        document = f'''<w:document xmlns:w="{W}" xmlns:wp="{WP}"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing">
<w:body>{directive}<w:p><w:r><w:t>{{{{ content }}}}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'''
        template = self.root / 'template.docx'
        with ZipFile(template, 'w') as archive:
            archive.writestr('word/document.xml', document)
            archive.writestr('word/styles.xml', self.styles_xml)
            archive.writestr('word/numbering.xml', f'<w:numbering xmlns:w="{W}"/>')
            archive.writestr('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
            archive.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
            archive.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        data = Catslap({'content': content or f'<p><img src="{self.src}"></p>'}).process(str(template))
        self.rendered = data
        with ZipFile(BytesIO(data)) as archive:
            root = ET.fromstring(archive.read('word/document.xml'))
            anchors = root.findall(f'.//{{{WP}}}anchor')
            return root, [tuple(int(a.get(key)) for key in ('distT', 'distR', 'distB', 'distL')) for a in anchors]

    def test_requested_example(self):
        root, margins = self.render('1 0 0 0')
        self.assertEqual(margins, [(360000, 0, 0, 0)])
        self.assertNotIn('image_margin', ''.join(root.itertext()))

    def test_top_margin_moves_anchor_and_reserves_vertical_flow(self):
        root, margins = self.render('4 0 0 0')
        self.assertEqual(margins, [(1440000, 0, 0, 0)])
        anchor = root.find(f'.//{{{WP}}}anchor')
        self.assertEqual(anchor.find(f'{{{WP}}}positionV/{{{WP}}}posOffset').text, '1440000')
        self.assertIsNotNone(anchor.find(f'{{{WP}}}wrapTopAndBottom'))
        self.assertIsNone(anchor.find(f'{{{WP}}}wrapSquare'))

    def test_bottom_margin_also_keeps_following_text_below_image(self):
        root, _ = self.render('0 0 2 0')
        anchor = root.find(f'.//{{{WP}}}anchor')
        self.assertEqual(anchor.find(f'{{{WP}}}positionV/{{{WP}}}posOffset').text, '0')
        self.assertIsNotNone(anchor.find(f'{{{WP}}}wrapTopAndBottom'))

    def test_unconfigured_anchor_keeps_original_position_and_wrapping(self):
        root, _ = self.render()
        anchor = root.find(f'.//{{{WP}}}anchor')
        self.assertEqual(anchor.find(f'{{{WP}}}positionV/{{{WP}}}posOffset').text, '0')
        self.assertIsNotNone(anchor.find(f'{{{WP}}}wrapSquare'))

    @unittest.skipUnless(os.environ.get('CATSLAP_TEST_SOFFICE'), 'Set CATSLAP_TEST_SOFFICE to enable visual layout regression')
    def test_rendered_four_cm_gap_and_following_text_order(self):
        stream = BytesIO()
        Image.new('RGB', (120, 60), 'red').save(stream, format='PNG')
        src = 'data:image/png;base64,' + base64.b64encode(stream.getvalue()).decode()
        content = f'<p>BEFORE IMAGE</p><p><img src="{src}"></p><p>AFTER IMAGE</p>'
        for name, margin in (('zero', '0 0 0 0'), ('four', '4 0 0 0')):
            self.render(margin, content)
            (self.root / (name + '.docx')).write_bytes(self.rendered)
        subprocess.run([
            os.environ['CATSLAP_TEST_SOFFICE'],
            '-env:UserInstallation=' + (self.root / 'lo-profile').as_uri(),
            '--headless', '--convert-to', 'png', '--outdir', str(self.root),
            str(self.root / 'zero.docx'), str(self.root / 'four.docx'),
        ], check=True, capture_output=True, timeout=45)
        boxes = []
        for name in ('zero', 'four'):
            with Image.open(self.root / (name + '.png')) as rendered:
                rgb = rendered.convert('RGB')
            red_mask = Image.eval(rgb.getchannel('R'), lambda v: 255 if v > 180 else 0)
            for channel in ('G', 'B'):
                red_mask = ImageChops.multiply(red_mask, Image.eval(rgb.getchannel(channel), lambda v: 255 if v < 80 else 0))
            box = red_mask.getbbox()
            self.assertIsNotNone(box)
            boxes.append(box)
            if name == 'four':
                below_image = rgb.crop((0, box[3], rgb.width, rgb.height)).convert('L')
                text_mask = Image.eval(below_image, lambda v: 255 if v < 80 else 0)
                self.assertIsNotNone(text_mask.getbbox(), 'Following text must appear below the image')
        # Writer's default PNG export uses 96 DPI: 4 cm is about 151 px.
        self.assertAlmostEqual(boxes[1][1] - boxes[0][1], 4 * 96 / 2.54, delta=3)

    def test_decimal_values_and_order(self):
        self.assertEqual(self.render('0.1 0.25 1.5 2')[1], [(36000, 90000, 540000, 720000)])

    def test_explicit_zero_overrides_default_sides(self):
        self.assertEqual(self.render('0 0 0 0')[1], [(0, 0, 0, 0)])

    def test_defaults_unchanged(self):
        self.assertEqual(self.render()[1], [(0, 114300, 0, 114300)])

    def test_every_generated_image_including_table_cells(self):
        img = f'<img src="{self.src}">'
        content = f'<p>{img}</p><p>Text {img}</p><table><tr><td>{img}</td></tr></table>'
        self.assertEqual(self.render('1 2 3 4', content)[1], [(360000, 720000, 1080000, 1440000)] * 3)

    def test_invalid_margins_leave_previous_setting_unchanged(self):
        expected = self.styles.setStyle('image_margin', '1 2 3 4')
        for value in ('', '1', '1 2 3', '1 2 3 4 5', '-1 0 0 0', 'x 0 0 0',
                      'NaN 0 0 0', 'Infinity 0 0 0', '1e999 0 0 0'):
            with self.subTest(value=value):
                with self.assertRaisesRegex(ValueError, 'image_margin'):
                    self.styles.setStyle('image_margin', value)
                self.assertEqual(self.styles.image_margin, expected)

    def test_fractional_emu_rounding_and_style_prefix(self):
        self.assertEqual(self.styles.setStyle('style_image_margin', '0.000001 0 0 0'), (0, 0, 0, 0))
        self.assertEqual(self.styles.setStyle('image_margin', '0.000002 0 0 0'), (1, 0, 0, 0))
        self.assertEqual(self.styles.setStyle('paragraph', 'My Style'), 'MyStyle')

    def test_blob_images_receive_margins(self):
        expected = self.styles.setStyle('image_margin', '1 0.5 0.25 0')
        types_path = self.root / '[Content_Types].xml'
        types_path.write_text('<Types/>')
        relationships = Relationships(str(self.root), str(self.root / 'word/_rels/document.xml.rels'), load=False)
        with patch.object(elements.html, 'fetch_blob_image', return_value=('image/png', self.image)):
            run = elements.create_run(XmlTag('w:r'), '', {'image': 'blob:test'}, relationships, ContentTypes(str(types_path)), self.styles)
        anchor = run.get_tag('w:drawing').get_tag('wp:anchor')
        self.assertEqual(tuple(int(anchor.get_attr(key)) for key in ('distT', 'distR', 'distB', 'distL')), expected)


if __name__ == '__main__':
    unittest.main()
