#!/usr/bin/env python
#-*- coding:utf-8 -*-

import sys
from sw_openoffice_xml import OOWriterXML, NS_MAP

# xPath for the offending property
charPropsXPath = './/office:styles/style:style[@style:family="text"]' \
                        '/style:text-properties[@style:font-size-asian]'

source_file = sys.argv[1]

# open the ODT file and get the styles XML
odt = OOWriterXML(source_file)
styles = odt.get_styles_xml()

# ditch 'style:font-size-asian' attribs from all 'style:text-properties'
#   elements which have them
for el in styles.xpath(charPropsXPath, namespaces=NS_MAP):
    del el.attrib['{{{style}}}font-size-asian'.format(**NS_MAP)]

# save the modified ODT file
odt.save_styles(styles)
odt.save_as('%s_fixed.odt' % source_file[:-4])
