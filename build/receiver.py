# Tiny local receiver: the browser page POSTs an image blob here and we save it.
import http.server, sys, os, time
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "creatives")
class H(http.server.BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*"); self.send_header("Access-Control-Allow-Headers", "*"); self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
    def do_OPTIONS(self):
        self.send_response(204); self._cors(); self.end_headers()
    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0)); data = self.rfile.read(n)
        name = self.headers.get("X-Name") or self.path.strip("/").split("?")[0] or f"img_{int(time.time())}.png"
        name = os.path.basename(name)
        with open(os.path.join(OUT, name), "wb") as f: f.write(data)
        self.send_response(200); self._cors(); self.end_headers(); self.wfile.write(b"ok")
        sys.stderr.write(f"saved {name} {n} bytes\n")
    def log_message(self, *a): pass
http.server.ThreadingHTTPServer(("127.0.0.1", 8766), H).serve_forever()
