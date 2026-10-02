"""Gera data/base_demo.xlsx com dados 100% FICTÍCIOS e as mesmas abas/colunas da
planilha real (Planilha1, Tratativas, Motivos).

A aba Tratativas já traz casos em cada etapa da cadência (datas relativas a
HOJE), para demonstrar tudo sem esperar 10 dias:
  - sem histórico                -> Envio 1
  - Envio 1 há 15 dias           -> Envio 2 (lembrete)
  - Envio 1 há 3 dias            -> aguardando prazo
  - Envio 2 há 12 dias           -> escalar para humano
  - já respondida / em tratativa -> automação não mexe

Uso: python scripts/gerar_base_demo.py [caminho_saida]
Por ser criado pelo próprio Python, o arquivo não tem rótulo de confidencialidade
nem proteção - evita os bloqueios de política em arquivos baixados/salvos pelo Excel.
"""
import random
import sys
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

RAIZ = Path(__file__).resolve().parent.parent
SAIDA = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "data" / "base_demo.xlsx"
HOJE = datetime.now()

random.seed(7)

CONTATOS = {  # fictícios; domínio exemplo.com não existe de verdade
    "ana": ("Ana Souza", "(11) 90000-0001", "ana.souza@exemplo.com", "Gestora de Loja"),
    "carlos": ("Carlos Lima", "(11) 90000-0002", "carlos.lima@exemplo.com", "Gerente"),
    "paula": ("Paula Rocha", "(21) 90000-0003", "paula.rocha@exemplo.com", "Coordenadora de TI"),
    "rafael": ("Rafael Costa", "(31) 90000-0004", "rafael.costa@exemplo.com", "Supervisor"),
    "julia": ("Juliana Alves", "(41) 90000-0005", "juliana.alves@exemplo.com", "Analista de Suporte"),
    "marcos": ("Marcos Dias", "(51) 90000-0006", "marcos.dias@exemplo.com", "Gerente Regional"),
}
RUAS = ["Av. Paulista", "Rua das Flores", "Av. Brasil", "Rua Augusta", "Av. Rio Branco",
        "Rua XV de Novembro", "Av. Getúlio Vargas", "Rua Oscar Freire"]
BAIRROS = ["Centro", "Vila Nova", "Jardim América", "Boa Vista", "Santa Cruz", "Moema"]
UFS = {"ana": "SP", "carlos": "SP", "paula": "RJ", "rafael": "MG", "julia": "PR", "marcos": "RS"}
MODELOS = ["HP LaserJet M428", "HP LaserJet M507", "HP Color M479", "HP 4003dw"]

# (contato, status na base, histórico em Tratativas: [(dias_atras, canal, status, obs)])
AGUARD, RESP = "Aguardando Retorno", "Respondido"
CENARIOS = [
    ("ana", "Sem Monitoramento", []),                                        # Envio 1 (grupo Ana)
    ("ana", "Sem Monitoramento", []),
    ("ana", "Sem Monitoramento", []),
    ("carlos", "Sem Monitoramento", []),                                     # Envio 1 (Carlos)
    ("paula", "Sem Monitoramento", [(15, "Email", AGUARD, "[AUTO] Envio 1")]),   # lembrete
    ("paula", "Sem Monitoramento", [(15, "Email", AGUARD, "[AUTO] Envio 1")]),
    ("rafael", "Sem Monitoramento", [(3, "Email", AGUARD, "[AUTO] Envio 1")]),   # aguardando prazo
    ("rafael", "Sem Monitoramento", [(4, "Email", AGUARD, "[AUTO] Envio 2")]),   # aguardando prazo
    ("julia", "Sem Monitoramento", [(26, "Email", AGUARD, "[AUTO] Envio 1"),
                                    (12, "Email", AGUARD, "[AUTO] Envio 2")]),   # escalar
    ("marcos", "Sem Monitoramento", [(30, "Email", AGUARD, "[AUTO] Envio 1"),
                                     (18, "Email", AGUARD, "[AUTO] Envio 2"),
                                     (2, "Email", "Encaminhado ao Humano",
                                      "[AUTO] Escalado p/ humano após 2 envio(s)")]),
    ("carlos", "Sem Monitoramento", [(9, "Email", AGUARD, "[AUTO] Envio 1"),
                                     (5, "Email", RESP, "Cliente informou: Loja em reforma (M01)")]),
    ("julia", "Sem Monitoramento", [(7, "Telefone", "Em Tratativa", "Verificando cabo de rede")]),
    ("paula", "Sem Monitoramento", [(6, "Email", "Concluído", "Impressora reconectada")]),
    ("sem_email", "Sem Monitoramento", []),                                  # sem e-mail válido
    ("ana", "Perdeu Monitoramento", []),
    ("rafael", "Perdeu Monitoramento", []),
    ("julia", "Parou Recentemente", []),
]
FILLER = [("Monitorada", 28)]

wb = Workbook()
fino = Side(style="thin", color="BFBFBF")
borda = Border(left=fino, right=fino, top=fino, bottom=fino)


def estilizar(ws, larguras):
    for c in ws[1]:
        c.font, c.border = Font(bold=True), borda
        c.fill = PatternFill("solid", fgColor="D9E1F2")
        c.alignment = Alignment(horizontal="center")
    for linha in ws.iter_rows(min_row=2):
        for c in linha:
            c.border = borda
    for letra, w in zip("ABCDEFGHIJK", larguras):
        ws.column_dimensions[letra].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def data_txt(dias_atras):
    return (HOJE - timedelta(days=dias_atras)).strftime("%d/%m/%Y")


# ------------------------------- Planilha1 -------------------------------
ws = wb.active
ws.title = "Planilha1"
ws.append(["Serial", "Status", "Data Identificação", "Endereço", "Número Para Contato",
           "Email", "Nome", "Cargo", "Cliente", "UF", "Modelo"])
tratativas, n = [], 0
linhas = [(c, s, h) for c, s, h in CENARIOS] + \
         [(random.choice(list(CONTATOS)), s, []) for s, qtd in FILLER for _ in range(qtd)]
for chave, status, hist in linhas:
    n += 1
    serial = f"BR{n:03d}"
    if chave == "sem_email":
        nome, tel, email, cargo, uf = "Contato sem e-mail", "(11) 90000-0099", "", "Atendente", "SP"
    else:
        nome, tel, email, cargo = CONTATOS[chave]
        uf = UFS[chave]
    ws.append([serial, status, data_txt(random.randint(1, 20)),
               f"{random.choice(RUAS)}, {random.randint(10, 999)} - {random.choice(BAIRROS)}",
               tel, email, nome, cargo, "Cliente Demo", uf, random.choice(MODELOS)])
    for dias, canal, st, obs in hist:
        tratativas.append((dias, [serial, data_txt(dias), canal, st, obs]))
dv = DataValidation(type="list", formula1='"Monitorada,Com Monitoramento,Sem Monitoramento,'
                    'Perdeu Monitoramento,Parou Recentemente"')
ws.add_data_validation(dv)
dv.add(f"B2:B{ws.max_row + 500}")
estilizar(ws, [10, 22, 18, 46, 20, 30, 20, 22, 14, 6, 20])

# ------------------------------- Tratativas -------------------------------
wt = wb.create_sheet("Tratativas")
wt.append(["Serial", "Data Contato", "Canal", "Status Tratativa", "Observação"])
for _, linha in sorted(tratativas, key=lambda x: -x[0]):     # mais antigas primeiro
    wt.append(linha)
estilizar(wt, [10, 16, 14, 24, 46])

# --------------------------------- Motivos ---------------------------------
wm = wb.create_sheet("Motivos")
wm.append(["Código", "Motivo"])
for linha in [("M01", "Loja em reforma"), ("M02", "Impressora trocada"), ("M03", "Sem rede"),
              ("M04", "Perda IP"), ("M05", "Sem energia"), ("M06", "Equipamento removido"),
              ("M07", "HP 4003dw perdeu configuração")]:
    wm.append(linha)
estilizar(wm, [10, 34])

SAIDA.parent.mkdir(exist_ok=True)
wb.save(SAIDA)
print(f"Base demo criada: {SAIDA}  ({n} impressoras, {wt.max_row - 1} tratativas)")
