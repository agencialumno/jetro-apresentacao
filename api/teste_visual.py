from generate_pptx import gerar_apresentacao

dados = {
    "condominio": "Condomínio Teste",
    "tipo": "Assembleia Geral Ordinária",
    "pauta": [
        {"texto": "Prestação de contas da administração", "role": "prestacao"},
        {"texto": "Eleição de síndico, subsíndico e conselho fiscal", "role": "topico", "conteudo": "Síndico: João da Silva\nSubsíndico: Maria Souza"},
        {"texto": "Previsão orçamentária", "role": "previsao"},
        {"texto": "Assuntos gerais", "role": "assuntos"},
    ],
    "a11": {
        "months": ["Mai/25", "Jun/25", "Jul/25", "Ago/25", "Set/25", "Out/25"],
        "saldoAnterior": [1000, 2000, 1500, 1800, 2200, 2500],
        "totalReceitas": [15000, 15200, 14800, 15500, 15300, 15600],
        "totalDespesas": [14000, 15700, 14500, 14800, 15000, 15300],
        "saldoFinal": [2000, 1500, 1800, 2200, 2500, 2800],
        "energiaEletrica": [1200, 1350, 900, 1100, 1250, 1180],
        "aguaEsgotoDespesa": [2000, 2100, 1950, 2050, 2150, 2080],
    },
    "inadimplencia": {
        "valor": "85.253,35",
        "cobrancas": 229,
        "unidades": 63,
        "organica": "22,00",
        "garantidora": "39,00",
        "total": "44,00",
        "aging": [
            {"label": "1-5 dias", "valor": 500, "quantidade": 3},
            {"label": "6-15 dias", "valor": 1200, "quantidade": 5},
            {"label": "16-30 dias", "valor": 2300, "quantidade": 8},
            {"label": "31-60 dias", "valor": 4800, "quantidade": 12},
            {"label": "61-90 dias", "valor": 3600, "quantidade": 9},
            {"label": "90+ dias", "valor": 72853.35, "quantidade": 26},
        ],
    },
    "b16": {
        "contas": [
            {
                "nome": "Conta Principal",
                "saldoAnt": 5000,
                "creditos": 15000,
                "debitos": 14500,
                "saldoFinal": 5500,
                "movimentos": [
                    {"data": "01/10/2025", "descricao": "Pagamento taxa condominial", "tipo": "C", "valor": 500},
                    {"data": "05/10/2025", "descricao": "Pagamento conta de luz", "tipo": "D", "valor": 1200},
                ],
            }
        ]
    },
    "periodo16b": "De 01/10/2025 até 31/10/2025",
    "previsao": {
        "itensCompletos": [
            {"tipo": "grupo", "nome": "CONTRATOS"},
            {"tipo": "item", "nome": "Limpeza", "atualRaw": 5000, "previstoRaw": 5500},
            {"tipo": "item", "nome": "Manutenção", "atualRaw": 5000, "previstoRaw": 5500},
            {"tipo": "subtotal", "nome": "Sub-Total", "atualRaw": 10000, "previstoRaw": 11000},
            {"tipo": "grupo", "nome": "PESSOAL"},
            {"tipo": "item", "nome": "Porteiro", "atualRaw": 12000, "previstoRaw": 13000},
            {"tipo": "item", "nome": "Zelador", "atualRaw": 8000, "previstoRaw": 8500},
            {"tipo": "subtotal", "nome": "Sub-Total", "atualRaw": 20000, "previstoRaw": 21500},
        ],
        "categories": [],
    },
    "canais": {
        "whats": "(21) 99866-6640",
        "tel": "(21) 3596-6640",
        "email": "contato@jetroadm.com.br",
        "site": "www.jetroadministradora.com.br",
        "endereco": "Avenida Ayrton Senna - 5500 - Bloco 3",
    },
}

gerar_apresentacao(dados, output_path="teste_visual.pptx")
print("gerado com sucesso")