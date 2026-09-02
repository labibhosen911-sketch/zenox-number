import os
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import main as bot_main


class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path in ('/', '/health'):
            body = b'ZENOX bot is running.'
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        return


def run():
    bot_thread = threading.Thread(target=bot_main.main, daemon=True)
    bot_thread.start()

    port = int(os.environ.get('PORT', '10000'))
    server = HTTPServer(('0.0.0.0', port), HealthHandler)
    print(f'Health server listening on 0.0.0.0:{port}')
    server.serve_forever()


if __name__ == '__main__':
    run()
