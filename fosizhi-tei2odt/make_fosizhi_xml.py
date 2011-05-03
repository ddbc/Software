#!/usr/bin/env python
#-*- coding:utf-8 -*-

import os
import sys

from sw_xml import *
from copy import deepcopy

XML_NS = 'http://www.w3.org/XML/1998/namespace'
TEI_NS = 'http://www.tei-c.org/ns/1.0'


def formatTree(el, indent='  ', level=0):
    i = '\n%s' % (level*indent)
    if el.tag == etree.Comment:
        el.getparent().remove(el)
    elif len(el):
        if not el.text or not el.text.strip():
            el.text = '%s%s' % (i, indent)
        for e in el:
            formatTree(e, indent, level + 1)
            if not e.tail or not e.tail.strip():
                e.tail = '%s%s' % (i, indent)
        if not e.tail or not e.tail.strip():
            e.tail = i
    else:
        if level and (not el.tail or not el.tail.strip()):
            el.tail = i
        if el.text and el.text.strip():
            txt = el.text.strip('%s%s' % ('\n', string.whitespace))
            if txt.count('\n'):
                el.text = '%s%s%s%s' % (i, indent, txt.replace('\n', '%s%s' % (i, indent)), i)
            else:
                el.text = txt
        if el.tail and el.tail.strip():
            txt = el.tail.strip('%s%s' % ('\n', string.whitespace))
            if txt.count('\n'):
                el.tail = '%s%s%s%s' % (i, indent, txt.replace('\n', '%s%s' % (i, indent)), i)
            else:
                el.tail = txt


def removeUnusedCharDecl(tei):
    """ get rid of glyphs declared in the master charDecl section which
        are not actually used in the present 志
    """

    refs = []

    wrapper = tei.find('.//{%s}div[@type="wrapper"]' % TEI_NS)
    for g in wrapper.iter('{%s}g' % TEI_NS):
        ref = g.get('ref')[1:]
        if ref not in refs:
                refs.append(ref)

    charDecl = tei.find('.//{%s}charDecl' % TEI_NS)
    if charDecl is not None:
        for glyph in charDecl.iter('{%s}glyph' % TEI_NS):
            ref = glyph.get('{%s}id' % XML_NS)
            if ref in refs:
                refs.remove(ref)
            else:
                charDecl.remove(glyph)

    """ report any missing Gaiji """
    if len(refs):
        logging.warning(
            'The folling Gaijis are not found in the encodingDesc:\n\t%s' % \
                    '\n\t'.join(refs))


def dumpTEI(tei):
    """ remove unused glyph declarations """
    removeUnusedCharDecl(tei)
    encodingDesc = tei.find(
            './{%s}teiHeader/{%s}encodingDesc' % (TEI_NS, TEI_NS))
    if encodingDesc is not None:
        #del encodingDesc.attrib['{%s}base' % XML_NS]
        if len(encodingDesc) < 1:
            tei.remove(encodingDesc)

    """ cleanup namespaces """
    tei_new = etree.Element(
        '{%s}TEI' % TEI_NS, nsmap={None: TEI_NS, 'xml': XML_NS})
    tei_new[:] = deepcopy(tei[:])
    etree.cleanup_namespaces(tei_new)

    """ output a nice clean, complete TEI doc """
    stripComments(tei_new)
    #formatTree(tei_new)
    teiXML = etree.tostring(
        tei_new, pretty_print=True, xml_declaration=False, encoding='utf-8')

    print teiXML


if __name__ == '__main__':

    iFile = sys.argv[1]

    """ get the TEI """
    try:
        tei = etree.parse(iFile)
    except IOError:
        print >> sys.stderr, 'bad input file %s' % iFile
        raise SystemExit

    """ parse XIncludes """
    tei.xinclude()
    tei = tei.getroot()

    dumpTEI(tei)

