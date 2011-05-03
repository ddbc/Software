#!/usr/bin/env python
#-*- coding:utf-8 -*-

""" Build Fosizhi ODTs """

import sys
import os
from lxml import etree
import logging
from sw_xml import stripNamespaces
from sw_misc import prep_logging, get_parser
from sw_uno import OOWriter

import settings

PROCESSED_TAGS = []


def load_settings(settings):
    """ Function to load settings from external module. """
    _thismodule = sys.modules[__name__]
    _m = settings
    for _k in dir(_m):
        if _k.isupper() and not _k.startswith('__'):
            setattr(_thismodule, _k, getattr(_m, _k))


def render_elm(writer, elm, stack=None, context=None):
    """ Renders a TEI element in OpenOffice. """

    if stack is None:
        stack = []

    if elm.tag in FOOTNOTE_TAGS:
        # create a footnote, set the cursor context...
        footnote = writer.create_footnote()
        saved_context = writer.context
        writer.set_context(footnote)
        #  ...and flush the para_style stack
        stack = ['Footnote']
    else:
        # otherwise we're still in the main context, so just append the
        #  tage name to the para_style stack
        stack.append(elm.tag)

    if elm.tag in CHARSTYLE_TAGS:
        char_style = elm.tag
        if elm.get('rend') is not None:
            char_style = '_'.join([char_style, elm.get('rend')])

        if char_style not in PROCESSED_TAGS:
            writer.check_style_name(char_style, 'Character')
            PROCESSED_TAGS.append(char_style)

        # char_style is opened here, but not closed, in case there are
        #  sub-elements
        # (note that OOo nested tags don't work, so this may have to be
        #  implemented as a stack, at some point)
        # ALSO: this is dodgy anyway, probably a stack in the writer
        #       class is a better idea
        saved_charstyle = writer.cursor.CharStyleName
        writer.cursor.CharStyleName = char_style
        if elm.text:
            writer.write_string(elm.text)

    elif elm.tag in PARASTYLE_TAGS or elm.tag in FOOTNOTE_TAGS:
        if elm.tag not in PROCESSED_TAGS:
            PROCESSED_TAGS.append(elm.tag)

        para_style = '-'.join(stack)

        if para_style not in PROCESSED_TAGS:
            PROCESSED_TAGS.append(para_style)

        writer.open_para(para_style, elm.tag)
        if elm.text:
            # strip newlines needed for tag = 'item'
            writer.write_string(elm.text.strip('\n'))

    elif elm.tag in PASS_TAGS:
        if elm.text:
            writer.write_string(elm.text)

    elif elm.tag in IGNORE_TAGS:
        pass

    else:
        if elm.tag not in PROCESSED_TAGS and elm.tag not in dir(settings):
            logging.warning('tag "{0}" has no processing instuction and is not'
                            ' "pass"ed (behaviour undefined)!'.format(elm.tag))
            PROCESSED_TAGS.append(elm.tag)

    for sub_elm in elm.iterchildren(tag=etree.Element):
        render_elm(writer, sub_elm, stack, context)

    if elm.tag in PARASTYLE_TAGS:
        writer.close_para()

    if elm.tag in CHARSTYLE_TAGS:
        writer.cursor.CharStyleName = saved_charstyle \
                            if saved_charstyle != '' else 'Default'

    # check if there's a function in `settings` to process this element
    if elm.tag in dir(settings):
        func = getattr(settings, elm.tag)
        if callable(func):
            # the function should return True on success, or False on failure
            if not func(writer, elm, stack):
                logging.error('malformed %s element!\n%s',
                        elm.tag, etree.tostring(elm, pretty_print=True))
                raise SystemExit

    if elm.tag in FOOTNOTE_TAGS:
        # return the context after a footnote
        writer.set_context(saved_context)

    if elm.tail:
        writer.write_string(elm.tail)

    # pop the tag name back off the para_style stack
    del stack[-1]


def main():
    """ Process a TEI document. """

    load_settings(settings)

    parser = get_parser()

    parser.add_option('-H', '--headless', dest='headless', action='store_true',
                        default=False, help='Headless Operation')

    parser.add_option('-k', '--keepopen', dest='keepopen', action='store_true',
                        default=False, help='Keep the OpenOffice server open '
                                    'when finished (if started by the script)')

    parser.add_option('-g', '--gazetteer', dest='gazetteer', action='store',
                        help='Gazetteer to process (e.g. g008)')

    parser.add_option('--teiBase', dest='teiBase', action='store',
                        default=TEI_BASE,
                        help='path to TEI files (eXist dump) ({0})'\
                                                            .format(TEI_BASE))

    parser.add_option('-s', '--styles', dest='stylesFile', action='store',
                        help='ODT or OTT file to read styles from')

    #parser.add_option('-o', '--output', dest='destFile', action='store',
                        #help='''output file''')

    opts = parser.parse_args()[0]

    if opts.gazetteer is None:
        parser.print_help()
        raise SystemExit

    gaz = opts.gazetteer

    prep_logging(opts.verbose, opts.quiet)

    logging.debug('Begin!')

    # get the TEI
    try:
        xml_file = os.path.join(opts.teiBase, gaz, '{0}_main.xml'.format(gaz))
        tei = etree.parse(xml_file)
    except IOError:
        logging.error('''
        file "%s" could not be found!  make sure the tei is available in
            "%s", or specify the --teiBase option correctly''',
                xml_file, opts.teiBase)
        raise SystemExit

    # parse XIncludes
    tei.xinclude()
    tei = tei.getroot()

    # strip namespaces for clarity and cleanliness :)
    tei = stripNamespaces(tei)

    logging.debug('Successfully loaded and parsed XML for %s', gaz)

    # Initialize OOWriter class, and connect to OOo
    writer = OOWriter(headless=opts.headless, keepopen=opts.keepopen)

    # if a styles template file has been specified, load the styles now
    if opts.stylesFile:
        styles_file_path = os.path.abspath(
                            os.path.join(os.getcwd(), opts.stylesFile))
        writer.load_styles_from_file(styles_file_path)

    # get the main TEI body and start work!
    wrapper = tei.find('.//div[@type="wrapper"]')
    stack = []
    for elm in wrapper.iterchildren(tag=etree.Element):
        render_elm(writer, elm, stack)

    #render_elm(wrapper.find('./div[@id="g008_00.xml"]'))

    logging.debug('End!')


if __name__ == '__main__':
    main()
