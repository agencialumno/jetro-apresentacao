"""
Vercel Python Serverless Function.
Endpoint: POST /api/generate
Recebe o JSON com os dados da assembleia, gera o .pptx a partir do modelo.pptx
(usando generate_pptx.py) e devolve o arquivo binário pra download.
"""
import json
import os
import sys
from http.server import BaseHTTPRequestHandler

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from generate_pptx import gerar_apresentacao_bytes

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "modelo.pptx")


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            dados = json.loads(body)

            pptx_bytes = gerar_apresentacao_bytes(dados, template_path=TEMPLATE_PATH)

            nome_arquivo = "Apresentacao_" + "".join(
                c if c.isalnum() else "_" for c in dados.get("condominio", "Assembleia")
            )[:40] + ".pptx"

            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{nome_arquivo}"')
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(pptx_bytes)

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"erro": str(e)}).encode("utf-8"))

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()