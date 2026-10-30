"""
Vercel serverless entry point.

Vercel's Python runtime looks for a WSGI callable named `app` in this file.
Everything (templates, static, Excel, Word template, ready book) lives in the
project root, so no absolute folder paths are needed on the server.
"""

import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from web_app import app  # noqa: E402  (Flask WSGI app)

__all__ = ["app"]
