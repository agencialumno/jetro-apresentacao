"""
Servidor local de desenvolvimento (alternativa ao `vercel dev`, sem precisar de login/variáveis de ambiente).

  python dev_server.py            -> http://localhost:3000

- Serve os arquivos estáticos desta pasta (index.html, app.js, ...).
- POST /api/generate roda api/generate_pptx.py direto (mesmo código da função da Vercel).
- GET /api/superlogica/* é repassado para a API de produção. Motivo: essa API só libera CORS para o
  domínio de produção, então o navegador bloqueia chamadas feitas direto de http://localhost.
  Repassando pelo servidor local, o navegador só fala com a própria origem e o bloqueio some.
"""
import json
import os
import sys
import urllib.error
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

RAIZ = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(RAIZ, "api"))
from generate_pptx import gerar_apresentacao_bytes  # noqa: E402

API_REMOTA = "https://jetro-superlogica-api-agencia-lumno.vercel.app"
MODELO = os.path.join(RAIZ, "api", "modelo.pptx")
PORTA = int(os.environ.get("PORT", "3000"))


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=RAIZ, **kw)

    def do_GET(self):
        if self.path.startswith("/api/superlogica/"):
            return self._proxy()
        return super().do_GET()

    def _proxy(self):
        try:
            req = urllib.request.Request(API_REMOTA + self.path, headers={"User-Agent": "jetro-dev-proxy"})
            with urllib.request.urlopen(req, timeout=60) as r:
                corpo, status = r.read(), r.status
        except urllib.error.HTTPError as e:
            corpo, status = e.read(), e.code
        except Exception as e:  # rede/timeout
            corpo, status = json.dumps({"msg": "proxy local: " + str(e)}).encode(), 502
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(corpo)))
        self.end_headers()
        self.wfile.write(corpo)

    def do_POST(self):
        if self.path != "/api/generate":
            self.send_error(404)
            return
        try:
            dados = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))))
            pptx = gerar_apresentacao_bytes(dados, template_path=MODELO)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
            self.send_header("Content-Length", str(len(pptx)))
            self.end_headers()
            self.wfile.write(pptx)
        except Exception as e:
            corpo = json.dumps({"erro": str(e)}).encode("utf-8")
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(corpo)


if __name__ == "__main__":
    print(f"Servidor local em http://localhost:{PORTA}")
    ThreadingHTTPServer(("127.0.0.1", PORTA), Handler).serve_forever()
