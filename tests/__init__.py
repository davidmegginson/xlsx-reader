""" Common functions and variables for unit tests.
"""

import os

def resolve_path(filename):
    """Resolve a pathname for a test input file."""
    return os.path.join(os.path.dirname(__file__), 'files', filename)
