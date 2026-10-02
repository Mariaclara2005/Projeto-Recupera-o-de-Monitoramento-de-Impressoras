"""Regras de negócio (sem enviar nada, sem mexer em arquivo - só decide).

Cadência (lida do histórico da aba Tratativas):
  1) impressora SEM monitoramento e sem histórico        -> Envio 1
  2) sem resposta após REENVIO_DIAS                      -> Envio 2 (lembrete)
  3) sem resposta após REENVIO_DIAS do último envio      -> escalar para HUMANO
  Se o último status da tratativa for "Respondido / Em Tratativa / Concluído /
  Resolvido", a automação para: alguém já está cuidando.
"""
import re
from dataclasses import dataclass, field

import pandas as pd

from automacao.planilha import achar_coluna, norm

RE_ENVIO = re.compile(r"\[AUTO\]\s*Envio\s*(\d+)", re.IGNORECASE)
RE_ESCALADO = re.compile(r"\[AUTO\]\s*Escalado", re.IGNORECASE)


@dataclass
class Impressora:
    serial: str
    endereco: str
    nome: str
    email: str
    telefone: str = ""
    envios: int = 0


@dataclass
class Acao:
    tipo: str                 # "envio" | "escalar"
    numero: int               # nº do envio (1, 2...) ou nº de envios já feitos (escalar)
    destinatario: str         # e-mail do contato (envio) ou "" (escalar)
    nome: str
    impressoras: list = field(default_factory=list)


def parse_data(texto):
    if texto is None or str(texto).strip().lower() in ("", "nan", "nat"):
        return None
    t = str(texto).strip()
    iso = re.match(r"^\d{4}-\d{2}-\d{2}", t)
    d = pd.to_datetime(t, dayfirst=not iso, errors="coerce")
    return None if pd.isna(d) else d.to_pydatetime()


def _historico(trat):
    c_serial = achar_coluna(trat, "serial")
    c_data = achar_coluna(trat, "data")
    c_status = achar_coluna(trat, "status")
    c_obs = achar_coluna(trat, "observ")
    hist = {}
    for _, r in trat.iterrows():
        serial = str(r[c_serial]).strip()
        if not serial or serial.lower() == "nan":
            continue
        hist.setdefault(serial, []).append({
            "data": parse_data(r[c_data]),
            "status": norm(r[c_status]),
            "obs": "" if pd.isna(r[c_obs]) else str(r[c_obs]),
        })
    return hist


def decidir(historico, hoje, cfg):
    """Decide o próximo passo de UMA impressora.
    Retorna (acao, numero, motivo): acao em envio | escalar | aguardar | parar."""
    if not historico:
        return "envio", 1, "sem histórico"
    ultimo = historico[-1]
    if ultimo["status"] in cfg.STATUS_PARAR:
        return "parar", 0, f"último status: {ultimo['status']}"
    if any(RE_ESCALADO.search(h["obs"]) for h in historico):
        return "parar", 0, "já escalada para humano"

    envios = [(int(m.group(1)), h["data"]) for h in historico
              if (m := RE_ENVIO.search(h["obs"]))]
    if not envios:
        return "envio", 1, "sem envio automático anterior"

    n = max(k for k, _ in envios)
    datas = [d for k, d in envios if k == n and d]
    if datas:
        dias = (hoje.date() - max(datas).date()).days
        if dias < cfg.REENVIO_DIAS:
            return "aguardar", n, f"envio {n} há {dias} dia(s); espera {cfg.REENVIO_DIAS}"
    if n < cfg.MAX_ENVIOS:
        return "envio", n + 1, f"sem resposta ao envio {n}"
    return "escalar", n, f"sem resposta após {n} envio(s)"


def planejar(base, trat, hoje, cfg):
    """Cruza base + histórico e devolve (acoes, resumo)."""
    c_serial = achar_coluna(base, "serial")
    c_status = achar_coluna(base, "status")
    c_email = achar_coluna(base, "email")
    c_nome = achar_coluna(base, "nome")
    c_end = achar_coluna(base, "endere")
    c_tel = achar_coluna(base, "telefone", "para contato", "numero",
                         obrigatoria=False)
    hist = _historico(trat)

    resumo = {"total_base": len(base), "sem_monitoramento": 0,
              "sem_email": [], "aguardando": [], "parados": []}
    alvo = base[base[c_status].map(norm).isin([norm(s) for s in cfg.STATUS_ALVO])]
    resumo["sem_monitoramento"] = len(alvo)

    envios, escalar = {}, []
    for _, r in alvo.iterrows():
        serial = str(r[c_serial]).strip()
        email = norm(r[c_email])
        acao, numero, motivo = decidir(hist.get(serial, []), hoje, cfg)
        if acao == "aguardar":
            resumo["aguardando"].append(f"{serial} ({motivo})")
            continue
        if acao == "parar":
            resumo["parados"].append(f"{serial} ({motivo})")
            continue
        imp = Impressora(
            serial=serial,
            endereco=str(r[c_end]).strip(),
            nome="" if pd.isna(r[c_nome]) else str(r[c_nome]).strip(),
            email=email,
            telefone="" if not c_tel or pd.isna(r[c_tel]) else str(r[c_tel]).strip(),
            envios=numero if acao == "escalar" else numero - 1,
        )
        if acao == "escalar":
            escalar.append(imp)
        elif "@" not in email:
            resumo["sem_email"].append(serial)
        else:
            envios.setdefault((email, numero), []).append(imp)

    acoes = [Acao("envio", numero, email, imps[0].nome, imps)
             for (email, numero), imps in envios.items()]
    acoes.sort(key=lambda a: (a.numero, a.destinatario))
    if escalar:
        acoes.append(Acao("escalar", max(i.envios for i in escalar), "", "", escalar))
    return acoes, resumo
