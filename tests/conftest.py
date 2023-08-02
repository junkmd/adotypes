import logging

import pytest


@pytest.fixture(autouse=True, scope="package")
def suppress_comtypes_logging():
    comtypes_logger = logging.getLogger("comtypes")
    orig_lv = comtypes_logger.level
    comtypes_logger.setLevel(logging.ERROR)
    yield
    comtypes_logger.setLevel(orig_lv)
