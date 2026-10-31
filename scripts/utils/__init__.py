"""
Shared helpers for the weekly automation scripts.

Submodules are imported directly (`from scripts.utils.openlp import ...`) rather
than re-exported here: a package-level import of .openlp pulls in sqlite3 and the
config singleton, which is why welcome.py needs a second import shim to run inside
GIMP's interpreter.
"""
