from http.server import BaseHTTPRequestHandler, HTTPServer

HOST = "0.0.0.0"
PORT = 3000

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain")
        self.end_headers()
        self.wfile.write(b"Ola mundo")

server = HTTPServer((HOST, PORT), Handler)

print(f"Servidor rodando em http://{HOST}:{PORT}")
server.serve_forever()