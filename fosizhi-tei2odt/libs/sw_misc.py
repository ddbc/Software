# coding=utf-8

""" Library of miscellaneous functions. """

from __future__ import with_statement


def zip_dir(basedir, archivename):
    """ Zips an entire folder. """
    import os
    from contextlib import closing
    from zipfile import ZipFile, ZIP_DEFLATED

    assert os.path.isdir(basedir)
    with closing(ZipFile(archivename, 'w', ZIP_DEFLATED)) as zip_file:
        for root, dirs, files in os.walk(basedir):
            # NOTE: ignores empty directories
            for file_path in files:
                abs_file_path = os.path.join(root, file_path)
                zip_file_path = abs_file_path[len(basedir) + len(os.sep):]
                zip_file.write(abs_file_path, zip_file_path)


def make_temp_dir(cleanup=True):
    """ Returns the path to valid temp folder (on any platform),
        and removes it on exit, if requested.
    """
    import atexit
    import tempfile
    import shutil

    temp_dir = tempfile.mkdtemp()

    def cleanup_func():
        """ Helper function to remove the temp folder. """
        shutil.rmtree(temp_dir)

    if cleanup:
        atexit.register(cleanup_func)

    return temp_dir


def print_timing(func):
    """ Decorator function to provide timing analysis. """
    import time

    def wrapper(*arg):
        """ Decorator wrapper. """
        time1 = time.time()
        res = func(*arg)
        time2 = time.time()
        print '%s took %0.3f ms' % (func.func_name, (time2 - time1) * 1000.0)
        return res

    return wrapper


def prep_logging(verbose=False, quiet=False):
    """ Configures a logging instance (boilerplate code). """
    import logging

    log_level = logging.DEBUG if verbose else logging.INFO
    log_level = logging.CRITICAL if quiet else log_level
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s.%(msecs)03d: %(levelname)s: %(message)s',
        datefmt='%H:%M:%S',
    )


def get_parser():
    """ Returns a OptionParser instance with some default config. """
    import optparse

    parser = optparse.OptionParser()
    parser.add_option('-v', '--verbose', dest='verbose', action='store_true',
                        default=False, help="Increase verbosity")
    parser.add_option('-q', '--quiet', dest='quiet', action='store_true',
                        help="quiet operation")
    return parser
