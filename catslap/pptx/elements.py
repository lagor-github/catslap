# -*- coding: utf-8 -*-
# Catslap
# Author: Luis A. González
# MIT License (view LICENSE file)
# Copyright (c) 2026

# OOXML Python API by Luis Alberto González (SIA)

from catslap.base.relationships import Relationships
from catslap.base.types import ContentTypes
from catslap.utils.xml import XmlParserException, XmlTag

TAG_BU_FONT = 'a:buFont'
TAG_BU_AUTO_NUM = 'a:buAutoNum'

ATTR_TYPEFACE_OL = '+mj-lt'
ATTR_TYPEFACE_DOTTED = '•'


def create_paragraph(props: dict, content: str or list, relationships: Relationships, types: ContentTypes):
  """
  Creates a PowerPoint paragraph with properties and content.

  Args:
    props: Paragraph/run properties.
    content: Text or list of XmlTag runs.
    relationships: Document relationships.
    types: Document ContentTypes.

  Returns:
    Paragraph XmlTag.

  Raises:
    XmlParserException: If content is invalid.
  """
  para_tag = XmlTag('a:p')
  p_props_tag = XmlTag('a:pPr')
  align = props.get('align')
  if align:
    p_props_tag.add_attr('algn', align)
  ul_list = props.get('list')
  if ul_list:
    p_props_tag.add_attr('marL', 171450)
    p_props_tag.add_attr('indent', -171450)
    if ul_list == 'number-dot':
      p_props_tag.add_tag(XmlTag(TAG_BU_FONT, {'typeface': ATTR_TYPEFACE_OL}))
      p_props_tag.add_tag(XmlTag(TAG_BU_AUTO_NUM, {'type': 'arabicParenR'}))
    elif ul_list == 'number':
      p_props_tag.add_tag(XmlTag(TAG_BU_FONT, {'typeface': ATTR_TYPEFACE_OL}))
      p_props_tag.add_tag(XmlTag(TAG_BU_AUTO_NUM, {'type': 'arabicPeriod'}))
    elif ul_list == 'alpha-dot':
      p_props_tag.add_tag(XmlTag(TAG_BU_FONT, {'typeface': ATTR_TYPEFACE_OL}))
      p_props_tag.add_tag(XmlTag(TAG_BU_AUTO_NUM, {'type': 'alphaLcParenR'}))
    elif ul_list == 'alpha':
      p_props_tag.add_tag(XmlTag(TAG_BU_FONT, {'typeface': ATTR_TYPEFACE_OL}))
      p_props_tag.add_tag(XmlTag(TAG_BU_AUTO_NUM, {'type': 'alphaLcPeriod'}))
    else:
      p_props_tag.add_tag(XmlTag('a:buChar', {'char': ATTR_TYPEFACE_DOTTED}))
  para_tag.add_tag(p_props_tag)
  if isinstance(content, str):
    runprops = props.copy()
    run_tag = create_run(content, runprops, relationships, types)
    para_tag.add_tag(run_tag)
    return para_tag
  if isinstance(content, list):
    for item in content:
      para_tag.add_tag(item)
    return para_tag
  raise XmlParserException("str o list expected: " + str(type(content)))


def create_run(text: str, runprops: dict or None, relationships: Relationships, types: ContentTypes) -> XmlTag:
  """
  Creates a text run for PowerPoint.

  Args:
    text: Run text.
    runprops: Formatting properties.
    relationships: Document relationships.
    types: Document ContentTypes.

  Returns:
    Run XmlTag.
  """
  rpr_tag = XmlTag('a:rPr')
  url = None
  if runprops:
    url = runprops.get('link')
    if url and relationships:
      relationship = relationships.add_relationship_hyperlink(url)
      rpr_tag.add_tag(XmlTag('a:hlinkClick', {'r:id': relationship.rid}))
    if runprops.get('bold'):
      rpr_tag.add_attr('b', '1')
    if runprops.get('italic'):
      rpr_tag.add_attr('i', '1')
    if runprops.get('underline'):
      rpr_tag.add_attr('u', 'sng')
    if runprops.get('strike'):
      rpr_tag.add_attr('strike', 'sngStrike')
    if runprops.get('code'):  # -- <a:latin typeface="Courier" pitchFamily="2" charset="0"/>
      rpr_tag.add_tag(XmlTag('a:latin', {'typeface': 'Courier', 'pitchFamily': 2, 'charset': 0}))
      # -- tamaño de fuente (por defecto siempre 800 -> 8pt)
    size = runprops.get('size')
    if not size or not isinstance(size, int):
      size = 800
    elif size < 700:
      size = 700
    if size:
      rpr_tag.add_attr('sz', size)

    color = runprops.get('color')
    if color:
      if color.startswith('#'):
        color = color[1:]
      solig_tag = rpr_tag.add_tag(XmlTag('a:solidFill'))
      solig_tag.add_tag(XmlTag('a:srgbClr', {'val': color}))
  run_tag = XmlTag('a:r')
  run_tag.add_tag(rpr_tag)
  if text is None:
    text = ''
  t_tag = run_tag.add_tag(XmlTag('a:t', {'xml:space': 'preserve'}))
  t_tag.add_text(text)
  return run_tag
