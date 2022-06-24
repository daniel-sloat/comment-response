import logging
import functools


def initialize_logging(
    start_text: str="Logging initialized."
):
    logging.basicConfig(
        filename="log.log",
        filemode="w",
        level=logging.INFO,
        datefmt=r"%Y-%m-%d %H:%M:%S",
        format="%(asctime)s.%(msecs)03d [%(levelname)s] %(message)s",
    )
    logging.info(start_text)
    return None


def log_read_file(func):
    @functools.wraps(func)
    def wrapper(file_path):
        logging.info(f"Reading file: {file_path}")
        result = func(file_path)
        return result

    return wrapper

def log_automark(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        logging.info(f"Creating automark document...")
        result = func(*args, **kwargs)
        logging.info(f"Automark document created: {result}")
        return None

    return wrapper

def log_write_docx(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        logging.info(f"Creating comment-response document...")
        result = func(*args, **kwargs)
        logging.info(f"Comment-response document created: {result}")
        return None

    return wrapper

def log_exception(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
            return result
        except Exception as e:
            logging.exception(
                f"Exception raised in {func.__name__}. exception: {str(e)}"
            )
            raise e

    return wrapper


def quit_logging() -> None:
    logging.info("Finished.")
    logging.shutdown()
    return None
