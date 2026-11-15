#!/usr/bin/env python3
"""Tiny static server for the dashboard.

Serves this folder with caching disabled so the dashboard always sees the
latest data/dashboard.xlsx. Usage:  python3 serve.py [port]
"""
import http.server
import os
import socketserver
import sys

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
os.chdir(os.path.dirname(os.path.abspath(__file__)))


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".js": "text/javascript",
    }

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, must-revalidate")
        super().end_headers()


class Server(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True


if __name__ == "__main__":
    with Server(("127.0.0.1", PORT), NoCacheHandler) as httpd:
        print(f"Dashboard running at http://localhost:{PORT}  (Ctrl+C to stop)")
        print("Edit data/dashboard.xlsx and save - the dashboard refreshes automatically.")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            pass
