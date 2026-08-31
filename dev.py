"""Serve the add-in's files locally, over plain HTTP, while working on them.

Usage:  python dev.py [--port 5174]

This is for looking at pages in a browser — chiefly `pdffinder.html`, which is
an app in its own right and renders perfectly well with no Excel behind it (the
column simply stays empty). It is *not* for loading the add-in into Excel: Office
requires HTTPS, which is what `serve.py` is for.

The one thing a naive static server gets wrong here is MIME types. pdf.js ships
as `.mjs` and Tesseract as `.wasm`; a browser refuses both unless they arrive
labelled properly, and the finder then fails to load with nothing said.
"""

import argparse
import http.server
import socketserver
from functools import partial
from pathlib import Path

ROOT = Path(__file__).parent


class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        ".mjs": "text/javascript",
        ".js": "text/javascript",
        ".wasm": "application/wasm",
        ".traineddata": "application/octet-stream",
    }

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, *args):
        pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=5174)
    args = ap.parse_args()

    handler = partial(Handler, directory=str(ROOT))
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.ThreadingTCPServer(("127.0.0.1", args.port), handler) as httpd:
        print(f"  http://localhost:{args.port}/pdffinder.html")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
