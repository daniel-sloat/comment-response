"""Logging support functions."""
import functools
import logging
from pprint import pformat

from xlsx_rich_text.sheets.newdatasheet import NewDataSheet


def log(print_console=True):
    def inner(func):
        @functools.wraps(func)
        def wrapper(sheet: NewDataSheet, **config):
            logger_start(print_console)
            logging.info(
                "Reading sheet '%s' from '%s'...", sheet.sheetname, sheet.workbook.file
            )
            logging.info("Using configuration:\n%s", pformat(config, sort_dicts=False))
            result = func(sheet, **config)
            return result

        return wrapper

    return inner


def logger_start(print_console=True):
    msg_fmt = "%(asctime)s (elapsed: %(relativeCreated)dms) [%(levelname)s] %(message)s"
    logging.basicConfig(
        filename="log.log",
        filemode="w",
        level=logging.INFO,
        datefmt=r"%Y-%m-%d %H:%M:%S",
        format=msg_fmt,
    )
    if print_console:
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        formatter = logging.Formatter(msg_fmt)
        console.setFormatter(formatter)
        logging.getLogger().addHandler(console)
    logging.info("Logging initialized.")


def logger_quit():
    logging.info("Logging ended.")
    logging.shutdown()


def log_write(func):
    @functools.wraps(func)
    def wrapper(self, filename, *args, **kwargs):
        logging.info("Writing '%s'...", filename)
        result = func(self, filename, *args, **kwargs)
        logging.info("Saved '%s'.", filename)
        return result

    return wrapper
