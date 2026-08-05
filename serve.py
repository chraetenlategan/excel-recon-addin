"""
Serve the Reconcile add-in's static files over HTTPS for internal (LAN) hosting.

Office add-ins must be served over HTTPS (only localhost is exempt), so this
script needs a certificate + key. For a real network rollout, use a cert your
client machines trust (an internal AD CA cert, or a real cert if the host has a
resolvable DNS name) — a bare self-signed cert will make Excel silently fail to
load the pane unless it's in each machine's Trusted Root store.

Usage:
    pip install flask waitress
    python serve.py --cert cert.pem --key key.pem --port 443

To generate a quick self-signed cert for testing (PowerShell / OpenSSL):
    openssl req -x509 -newkey rsa:2048 -nodes -keyout key.pem -out cert.pem \
        -days 825 -subj "/CN=addins.corp.local" \
        -addext "subjectAltName=DNS:addins.corp.local"

The manifest's URLs must match the host/port you serve on. See
manifest.network.xml — replace https://YOUR-SERVER/recon-addin with, e.g.,
https://addins.corp.local  (if you serve at the site root, drop the
/recon-addin path segment from every URL).
"""

import argparse
import os
from flask import Flask, send_from_directory

HERE = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, static_folder=None)


@app.after_request
def add_headers(resp):
    # Office loads the pane in an iframe; make sure nothing blocks that and
    # that clients don't cache stale JS after you push an update.
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers.pop("X-Frame-Options", None)
    return resp


@app.route("/")
def root():
    return send_from_directory(HERE, "index.html")


@app.route("/<path:filename>")
def static_files(filename):
    # Serves taskpane.html, *.js, functions.json, taskpane.css, assets/*, etc.
    return send_from_directory(HERE, filename)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Serve the Reconcile add-in over HTTPS.")
    parser.add_argument("--cert", required=True, help="Path to the TLS certificate (PEM).")
    parser.add_argument("--key", required=True, help="Path to the TLS private key (PEM).")
    parser.add_argument("--host", default="0.0.0.0", help="Bind address (default: all interfaces).")
    parser.add_argument("--port", type=int, default=443, help="HTTPS port (default: 443).")
    args = parser.parse_args()

    # Flask's built-in server is fine for a small internal team. For heavier
    # use, put this behind IIS/nginx as a reverse proxy and let that terminate
    # TLS instead, or run under a production WSGI server (waitress/gunicorn).
    app.run(host=args.host, port=args.port, ssl_context=(args.cert, args.key))
