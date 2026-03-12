from .main import *

import importlib
import importlib.metadata

__version__ = None
try:
    __version__ = importlib.metadata.version("program_completion_check")
except importlib.metadata.PackageNotFoundError:
    pass

__all__ = [
    "get_client",
    "validate_program_students",
    "validate_utas_grade",
    "validate_program_subjects",
    "do_check",
]
