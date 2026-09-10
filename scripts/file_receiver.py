# -*- coding: utf-8 -*-
"""
本地文件接收服务：浏览器自动化导出的报表经页面内 fetch POST 落盘。
用法：
  python file_receiver.py [保存目录] [端口]
默认：保存目录 = 当前工作目录/data，端口 = 8765
"""
import http.server, socketserver, os, sys, urllib.parse

SAVE_DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.getcwd(), 'data')
PORT = int(sys.argv[2]) if len(sys.argv) > 2 else 8765
os.makedirs(SAVE_DIR, exist_ok=True)

class Handler(http.server.BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.send_header("Access-Control-Allow-Private-Network", "true")

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        self.send_response(200)
        self._cors()
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write("receiver ok".encode("utf-8"))

    def do_POST(self):
        qs = urllib.parse.urlparse(self.path)
        name = urllib.parse.parse_qs(qs.query).get("name", ["upload.bin"])[0]
        name = os.path.basename(name)
        length = int(self.headers.get("Content-Length", 0))
        data = self.rfile.read(length)
        path = os.path.join(SAVE_DIR, name)
        with open(path, "wb") as f:
            f.write(data)
        self.send_response(200)
        self._cors()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(('{"ok":true,"path":"%s","size":%d}' % (path.replace("\\", "/"), len(data))).encode("utf-8"))
        print("SAVED:", path, len(data), flush=True)

    def log_message(self, fmt, *args):
        print(fmt % args, flush=True)

with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
    print("listening on", PORT, "save_dir:", SAVE_DIR, flush=True)
    httpd.serve_forever()

