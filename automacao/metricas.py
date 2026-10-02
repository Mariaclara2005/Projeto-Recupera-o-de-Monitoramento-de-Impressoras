"""Foto diária dos números (data/snapshot_diario.csv) - alimenta o Power BI.

Uma linha por dia, regravada se rodar de novo no mesmo dia. É o histórico que
permite mostrar "Dados Atuais -> Dados Esperados -> Pós-Automatização".
"""
import pandas as pd

from automacao import config
from automacao.planilha import achar_coluna, norm
from automacao.regras import RE_ENVIO, RE_ESCALADO, _historico

GRUPOS = {
    "monitoradas": ("com monitoramento", "monitorada", "monitoradas"),
    "sem_monitoramento": ("sem monitoramento",),
    "perderam": ("perdeu monitoramento", "perderam monitoramento"),
    "pararam": ("parou recentemente", "pararam recentemente"),
}
EM_ANDAMENTO = {"respondido", "em tratativa"}
FECHADO = {"concluido", "resolvido"}


def calcular(base, trat, hoje):
    status = base[achar_coluna(base, "status")].map(norm)
    linha = {"data": hoje.strftime("%Y-%m-%d"), "total": len(base)}
    for nome, valores in GRUPOS.items():
        linha[nome] = int(status.isin(valores).sum())
    linha["pct_sem_monitoramento"] = round(
        linha["sem_monitoramento"] / linha["total"], 4) if linha["total"] else 0

    contatadas = aguardando = em_andamento = fechadas = escaladas = 0
    for h in _historico(trat).values():
        if any(RE_ENVIO.search(x["obs"]) for x in h):
            contatadas += 1
        if any(RE_ESCALADO.search(x["obs"]) for x in h):
            escaladas += 1
        ultimo = h[-1]["status"]
        if ultimo == "aguardando retorno":
            aguardando += 1
        elif ultimo in EM_ANDAMENTO:
            em_andamento += 1
        elif ultimo in FECHADO:
            fechadas += 1
    linha.update(contatadas=contatadas, aguardando_retorno=aguardando,
                 em_tratativa=em_andamento, concluidas=fechadas, escaladas=escaladas)
    return linha


def gravar(base, trat, hoje, caminho=None):
    caminho = caminho or config.SNAPSHOT_CSV
    nova = pd.DataFrame([calcular(base, trat, hoje)])
    if caminho.exists():
        antiga = pd.read_csv(caminho, dtype={"data": str})
        antiga = antiga[antiga["data"] != nova.loc[0, "data"]]
        nova = pd.concat([antiga, nova], ignore_index=True)
    caminho.parent.mkdir(exist_ok=True)
    nova.sort_values("data").to_csv(caminho, index=False, encoding="utf-8-sig")
    return caminho
