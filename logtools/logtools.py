import logging
import functools


def initialize_logging(
    start_text: str="Logging initialized.",
    print_console: bool=True,    
) -> None:
    logging.basicConfig(
        filename="log.log",
        filemode="w",
        level=logging.INFO,
        datefmt=r"%Y-%m-%d %H:%M:%S",
        format="%(asctime)s.%(msecs)03d [%(levelname)s] %(message)s",
    )
    if print_console:
        # Print logger message to console
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
        console.setFormatter(formatter)
        logging.getLogger().addHandler(console)
    
    logging.info(start_text)
    return None


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
