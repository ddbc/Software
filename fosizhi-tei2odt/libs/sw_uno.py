#-*- coding:utf-8 -*-

""" Library of OpenOffice / Uno / PyUno related functions. """

__program_name__ = 'sw_uno'
__version__ = '0.2'
__author__ = 'Simon Wiles'
__email__ = 'simonjwiles@gmail.com'
__copyright__ = 'Copyright (c) 2010-2011, Simon Wiles'
__license__ = 'GPL http://www.gnu.org/licenses/gpl.txt'
__date__ = 'April, 2011'

import os
import sys
import subprocess
import time
import atexit
import logging

try:
    import uno
    import unohelper
except ImportError:
    print >> sys.stderr, 'Unable to find pyUno -- aborting!',


# Note on com.sun.star.* imports -- using the uno.getClass() and
#  uno.getContantByName() methods is necessary for compatibility with
#  cx_freeze -- not too sure if I care about this or not...

# Exceptions
UnoException = uno.getClass('com.sun.star.uno.Exception')
NoConnectException = uno.getClass('com.sun.star.connection.NoConnectException')
RuntimeException = uno.getClass('com.sun.star.uno.RuntimeException')
IllegalArgumentException = uno.getClass(
                    'com.sun.star.lang.IllegalArgumentException')
DisposedException = uno.getClass('com.sun.star.lang.DisposedException')
IOException = uno.getClass('com.sun.star.io.IOException')
NoSuchElementException = uno.getClass(
                    'com.sun.star.container.NoSuchElementException')


# Control Characters
PARAGRAPH_BREAK = uno.getConstantByName(
                    'com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK')
LINE_BREAK = uno.getConstantByName(
                    'com.sun.star.text.ControlCharacter.LINE_BREAK')
HARD_HYPHEN = uno.getConstantByName(
                    'com.sun.star.text.ControlCharacter.HARD_HYPHEN')
SOFT_HYPHEN = uno.getConstantByName(
                    'com.sun.star.text.ControlCharacter.SOFT_HYPHEN')
HARD_SPACE = uno.getConstantByName(
                    'com.sun.star.text.ControlCharacter.HARD_SPACE')
APPEND_PARAGRAPH = uno.getConstantByName(
                    'com.sun.star.text.ControlCharacter.APPEND_PARAGRAPH')


# Styles
SLANT_ITALIC = uno.getConstantByName('com.sun.star.awt.FontSlant.ITALIC')
SLANT_NONE = uno.getConstantByName('com.sun.star.awt.FontSlant.NONE')
WEIGHT_BOLD = uno.getConstantByName('com.sun.star.awt.FontWeight.BOLD')
WEIGHT_NORMAL = uno.getConstantByName('com.sun.star.awt.FontWeight.NORMAL')
UNDERLINE_SINGLE = uno.getConstantByName(
                    'com.sun.star.awt.FontUnderline.SINGLE')
UNDERLINE_NONE = uno.getConstantByName('com.sun.star.awt.FontUnderline.NONE')
CASE_UPPER = uno.getConstantByName('com.sun.star.style.CaseMap.UPPERCASE')
CASE_SMALLCAPS = uno.getConstantByName('com.sun.star.style.CaseMap.SMALLCAPS')
CASE_NON = uno.getConstantByName('com.sun.star.style.CaseMap.NONE')


# Misc
URL = uno.getClass('com.sun.star.util.URL')
PropertyValue = uno.getClass('com.sun.star.beans.PropertyValue')
Locale = uno.getClass('com.sun.star.lang.Locale')
DIRECT_VALUE = uno.getConstantByName(
                    'com.sun.star.beans.PropertyState.DIRECT_VALUE')


def whereis(program):
    """ Generic function to find the location of a binary executable
        on the system path.
    """
    for path in os.environ.get('PATH', '').split(':'):
        if os.path.exists(os.path.join(path, program)) and \
           not os.path.isdir(os.path.join(path, program)):
            return os.path.join(path, program)
    return None


def connectOO(headless=False, keepopen=False,
                oo_host='127.0.0.1', oo_port='8100'):
    """ Open and/or connect to an OpenOffice server.  Returns an
        instance of com.sun.star.frame.Desktop.
    """
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
                    'com.sun.star.bridge.UnoUrlResolver', local)

    connection_string = \
        'uno:socket,host={0},port={1};urp;StarOffice.ComponentContext'

    try:
        office = None
        context = resolver.resolve(connection_string.format(oo_host, oo_port))
        logging.debug('Connected to OpenOffice on %s:%s', oo_host, oo_port)
    except NoConnectException:
        office = subprocess.Popen([
            whereis('soffice'),
            '-headless' if headless else '-invisible',
            '-nofirststartwizard',
            '-norestore',
            '-nologo',
            '-accept=socket,host={0},port={1};urp;'.format(oo_host, oo_port),
        ])
        time.sleep(3)

        def cleanup():
            """ Helper function to shut down the OpenOffice server on
                completion, if required.
            """
            office.terminate()
            logging.debug('Closed OpenOffice with PID %d', office.pid)
            office.wait()

        if not keepopen:
            atexit.register(cleanup)

        context = resolver.resolve(connection_string.format(oo_host, oo_port))
        logging.debug(
            'Opened and connected to OpenOffice on %s:%s (with PID %d)',
                            oo_host, oo_port, office.pid)

    desktop = context.ServiceManager.createInstanceWithContext(
                'com.sun.star.frame.Desktop', context)

    return desktop


class OOWriter():
    """ Class to manipulate OpenOffice Writer using the Uno API. """

    def __init__(self, **args):
        self.desktop = connectOO(**args)
        self.document = self.desktop.getCurrentComponent()

        if self.document is None:
            self.document = self.desktop.loadComponentFromURL(
                                'private:factory/swriter', '_blank', 0, ())
            #document = desktop.loadComponentFromURL(
                #'file:///home/simon/test.odt' ,'_blank', 0, ())

        self.saved_char_style_name = None

        self.contexts = [self.document.Text]
        self.cursors = {0: self.document.Text.createTextCursor()}
        self.cursor = None
        self.context = None
        self.set_context()

        self.view_cursor = self.document.getCurrentController().getViewCursor()

    def set_context(self, context=None):
        """ Set the document context in the OpenOffice document. """
        if context is None:
            self.context = self.document.Text
        else:
            self.context = context

        if self.context in self.contexts:
            self.cursor = self.cursors[self.contexts.index(self.context)]
        else:
            self.contexts.append(self.context)
            self.cursor = self.context.createTextCursor()
            self.cursors[self.contexts.index(self.context)] = self.cursor

    def check_style_name(self, style_name, style_type, parent_style_name=None):
        """ Checks if a style name exists, and creates it if not.  This is
            necessary as Uno will crash if attempting to assign a style which
            doesn't already exist!
        """
        styles = self.document.StyleFamilies.getByName(
                                                '{0}Styles'.format(style_type))

        if not styles.hasByName(style_name):
            style = self.document.createInstance(
                    'com.sun.star.style.{0}Style'.format(style_type))

            if parent_style_name is not None:
                self.check_style_name(parent_style_name, style_type)
                style.ParentStyle = parent_style_name
            elif style_type == 'Paragraph':
                style.ParentStyle = 'Default'

            styles.insertByName(style_name, style)

    def open_para(self, style_name='Default', parent_style_name=None):
        """ Begins a new paragraph, with specified style. """
        try:
            self.cursor.ParaStyleName = style_name
        except UnoException:
            self.check_style_name(style_name, 'Paragraph', parent_style_name)
            self.cursor.ParaStyleName = style_name

    def close_para(self):
        """ Closes the current paragraph (i.e. inserts a paragraph break). """
        self.context.insertControlCharacter(
                self.cursor, PARAGRAPH_BREAK, False)

    def write_para(self, text, style_name='Default', parent_style_name=None):
        """ Writes an entire paragraph in one go. """
        try:
            self.cursor.ParaStyleName = style_name
        except UnoException:
            self.check_style_name(style_name, 'Paragraph', parent_style_name)
            self.cursor.ParaStyleName = style_name

        self.context.insertString(self.cursor, text, False)
        self.context.insertControlCharacter(
                self.cursor, PARAGRAPH_BREAK, False)

    def write_string(self, text, style_name=None):
        """ Writes a simple string in the current Context, with the specified
            Character Style.
        """
        if style_name is not None:
            saved_char_style = self.cursor.CharStyleName
            try:
                self.cursor.CharStyleName = style_name
            except UnoException:
                self.check_style_name(style_name, 'Character')
                self.cursor.CharStyleName = style_name

        self.context.insertString(self.cursor, text, False)

        if style_name is not None:
            # ugh! don't seem to be able to "unset" this property..?
            self.cursor.CharStyleName = \
                saved_char_style if saved_char_style != '' else 'Default'

    def open_char_style(self, style_name='Default'):
        """ Begins a new Character Style. """
        self.saved_char_style_name = self.cursor.CharStyleName
        try:
            self.cursor.CharStyleName = style_name
        except UnoException:
            self.check_style_name(style_name, 'Character')
            self.cursor.CharStyleName = style_name
            self.cursor.CharStyleName = style_name

    def insert_footnote(self, text):
        """ quick convenience function """
        footnote = self.create_footnote()
        footnote_cursor = footnote.createTextCursor()
        footnote.insertString(footnote_cursor, text, False)

    def create_footnote(self):
        """ Returns a new footnote anchored at the current cursor. """
        footnote = self.document.createInstance('com.sun.star.text.Footnote')
        self.context.insertTextContent(self.cursor, footnote, False)
        return footnote

    def load_styles_from_file(self, file_path):
        """ Loads styles from a specified ODT file. """
        # Available options:
        #  LoadCellStyles
        #  LoadTextStyles
        #  LoadFrameStyles
        #  LoadPageStyles
        #  LoadNumberingStyles
        #  OverwriteStyles
        #
        # all default to True

        properties = (PropertyValue('OverwriteStyles', 0, True, 0),)
        url = unohelper.systemPathToFileUrl(file_path)
        self.document.StyleFamilies.loadStylesFromURL(url, properties)

    def create_index(self, anchor=None, index_type='toc', index_name=None,
                        index_title=''):
        """
            com.sun.star.text.DocumentIndex         alphabetical index
            com.sun.star.text.ContentIndex          table of contents
            com.sun.star.text.UserIndex             user defined index
            com.sun.star.text.IllustrationIndex     illustrations index
            com.sun.star.text.ObjectIndex           objects index
            com.sun.star.text.TableIndex            (text) tables index
            com.sun.star.text.Bibliography          bibliographical index

level_format = (
    (
        PropertyValue('TokenType', 0, 'TokenEntryText', DIRECT_VALUE),
        PropertyValue('CharacterStyleName', 0, '', DIRECT_VALUE),
    ),
    (
        PropertyValue('TokenType', 0, 'TokenTabStop', DIRECT_VALUE),
        PropertyValue('TabStopRightAligned', 0, True, DIRECT_VALUE),
        PropertyValue('TabStopFillCharacter', 0, ' ', DIRECT_VALUE),
        PropertyValue('CharacterStyleName', 0, '', DIRECT_VALUE),
        PropertyValue('WithTab', 0, True, DIRECT_VALUE),
    ),
    (
        PropertyValue('TokenType', 0, 'TokenPageNumber', DIRECT_VALUE),
        PropertyValue('CharacterStyleName', 0, '', DIRECT_VALUE)
    ),
)

# these two lines are a work around for bug #12504. See python-uno FAQ.
# since index.LevelFormat.replaceByIndex(2,level_format) does not work
level_format = uno.Any('[][]com.sun.star.beans.PropertyValue', level_format)
uno.invoke(index.LevelFormat, 'replaceByIndex', (2, level_format))


        """
        index_types = {
            'alpha': 'com.sun.star.text.DocumentIndex',
            'toc': 'com.sun.star.text.ContentIndex',
            'user': 'com.sun.star.text.UserIndex',
        }
        index = self.document.createInstance(index_types[index_type])
        index.Name = index_name or index_type
        index.Title = index_title
        if anchor is None:
            anchor = self.cursor
        anchor.Text.insertTextContent(anchor, index, False)
        return index

    def create_index_entry(self, entry, anchor=None,
                            mark_type='user', level=0, first_key=None):
        """
            com.sun.star.text.DocumentIndexMark     (for alphabetical indexes)
            com.sun.star.text.UserIndexMark         (for user defined indexes)
            com.sun.star.text.ContentIndexMark
                (for entries in TOCs which are independent of chapter headings)
        """
        mark_types = {
            'alpha': 'com.sun.star.text.DocumentIndexMark',
            'toc': 'com.sun.star.text.ContentIndexMark',
            'user': 'com.sun.star.text.UserIndexMark',
        }
        mark = self.document.createInstance(mark_types[mark_type])
        mark.setMarkEntry(entry)
        if mark_type == 'alpha' and first_key is not None:
            mark.PrimaryKey = first_key
        if mark_type != 'alpha':
            mark.setPropertyValue('Level', level)
        if anchor is None:
            anchor = self.cursor
        anchor.Text.insertTextContent(anchor, mark, False)
        return mark

    def save_odt(self, file_path):
        """ Save the ODT file. """
        #document.store()
        url = unohelper.systemPathToFileUrl(file_path)
        self.document.storeAsURL(url, ())
