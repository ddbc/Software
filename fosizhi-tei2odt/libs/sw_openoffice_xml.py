#!/usr/bin/env python

""" OpenOffice XML manipulation class. """

__program_name__ = 'sw_openoffice_xml'
__version__ = '0.1'
__author__ = 'Simon Wiles'
__email__ = 'simonjwiles@gmail.com'
__copyright__ = 'Copyright (c) 2010-2011, Simon Wiles'
__license__ = 'GPL http://www.gnu.org/licenses/gpl.txt'
__date__ = 'April, 2011'

import os
import zipfile
import codecs
import logging

from lxml import etree
from sw_misc import make_temp_dir, zip_dir


NS_MAP = {
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    'xlink': 'http://www.w3.org/1999/xlink',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
    'number': 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
    'svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
    'chart': 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
    'dr3d': 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0',
    'math': 'http://www.w3.org/1998/Math/MathML',
    'form': 'urn:oasis:names:tc:opendocument:xmlns:form:1.0',
    'script': 'urn:oasis:names:tc:opendocument:xmlns:script:1.0',
    'ooo': 'http://openoffice.org/2004/office',
    'ooow': 'http://openoffice.org/2004/writer',
    'oooc': 'http://openoffice.org/2004/calc',
    'dom': 'http://www.w3.org/2001/xml-events',
    'xforms': 'http://www.w3.org/2002/xforms',
    'xsd': 'http://www.w3.org/2001/XMLSchema',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    'rpt': 'http://openoffice.org/2005/report',
    'of': 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
    'xhtml': 'http://www.w3.org/1999/xhtml',
    'grddl': 'http://www.w3.org/2003/g/data-view#',
    'officeooo': 'http://openoffice.org/2009/office',
    'tableooo': 'http://openoffice.org/2009/table',
    'css3t': 'http://www.w3.org/TR/css3-text/',
}


class OOWriterXML():
    """ Class to access the contents of OpenOffice ODT documents. """

    def __init__(self, source_file, temp_dir=None):
        # an ODT file is just a regular Zip file...
        try:
            if source_file[-4:].lower() != '.odt':
                raise zipfile.BadZipfile
            source_zip_file = zipfile.ZipFile(source_file)
        except zipfile.BadZipfile:
            logging.fatal('File "%s" is not a valid ODT file!', source_file)
            raise SystemExit
        except IOError:
            logging.fatal(
                'File "%s" does not exist or cannot be read!', source_file)
            raise SystemExit
        except:
            logging.fatal('No input file specified?')
            raise SystemExit

        self.source_file = source_file

        # we need a temp folder
        if temp_dir is None:
            self.temp_dir = make_temp_dir(cleanup=True)
        else:
            self.temp_dir = temp_dir

        # extract to the temp folder
        source_zip_file.extractall(path=self.temp_dir)

    def save(self, output_file=None):
        """ Saves the ODT document (if no `output_file` is passed,
            then over-writes the original).
        """
        if output_file is None:
            output_file = self.source_file
        zip_dir(self.temp_dir, output_file)

    def save_as(self, output_file):
        """ Convenience method for backwards-compatibility. """
        self.save(output_file)

    def get_styles_path(self):
        """ Returns the path to the extracted `styles.xml`. """
        return os.path.join(self.temp_dir, 'styles.xml')

    def get_styles_xml(self):
        """ Returns the XML tree for the ODT styles. """
        return etree.parse(self.get_styles_path())

    def save_styles(self, styles):
        """ Writes the modified XML back to the temp file. """
        with codecs.open(self.get_styles_path(), 'w', 'utf-8') as output_file:
            output_file.write(
                    unicode(etree.tostring(styles, encoding='utf-8'), 'utf8'))

    def get_content_path(self):
        """ Returns the path to the extracted `content.xml`. """
        return os.path.join(self.temp_dir, 'content.xml')

    def get_content_xml(self):
        """ Returns the XML tree for the ODT content. """
        return etree.parse(self.get_content_path())

    def save_content(self, content):
        """ Writes the modified XML back to the temp file. """
        with codecs.open(self.get_content_path(), 'w', 'utf-8') as output_file:
            output_file.write(
                    unicode(etree.tostring(content, encoding='utf-8'), 'utf8'))
