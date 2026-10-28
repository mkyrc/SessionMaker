"""Logging setup library"""

import logging


def init_logging(verbose: int):
    """Initiate logging system.

    Args:
        verbose (int, default=0):
            0=critical,error,warning
            1=critical,error,warning,info
            2=critical,error,warning,info,debug
    """
    if type(verbose) is not int:
        verbose=0
        
    if verbose >= 2:
        logging.basicConfig(
            format="  %(levelname)s: %(message)s", level=logging.DEBUG, force=True
        )
    elif verbose == 1:
        logging.basicConfig(
            format="  %(levelname)s: %(message)s", level=logging.INFO, force=True
        )
    else:
        logging.basicConfig(
            format="  %(levelname)s: %(message)s", level=logging.WARNING, force=True
        )
