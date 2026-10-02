"""Linha de comando: python run.py [--modo ...] [--loop] [--diagnostico] [--snapshot]"""
import argparse
import logging
import sys
import time
from logging.handlers import RotatingFileHandler

from automacao import servico
from automacao import config
from automacao import planilha

log = logging.getLogger("automacao")


def _logging():
    config.PASTA_LOGS.mkdir(exist_ok=True)
    if hasattr(sys.stdout, "reconfigure"):      # emojis no terminal do Windows (cp1252)
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%d/%m/%Y %H:%M:%S")
    arquivo = RotatingFileHandler(config.PASTA_LOGS / "automacao.log", maxBytes=1_000_000,
                                  backupCount=5, encoding="utf-8")
    console = logging.StreamHandler(sys.stdout)
    for h in (arquivo, console):
        h.setFormatter(fmt)
    log.setLevel(logging.INFO)
    log.addHandler(arquivo)
    log.addHandler(console)


def _rodada(modo, arquivo):
    try:
        servico.rodar_uma_vez(modo=modo, arquivo=arquivo)
    except planilha.ErroPlanilha as e:
        log.error("%s", e)
    except Exception:  # noqa: BLE001 - o modo contínuo não pode morrer
        log.exception("Erro inesperado na rodada")


def main(argv=None):
    p = argparse.ArgumentParser(description="Automação de recuperação de monitoramento")
    p.add_argument("--modo", choices=["simular", "rascunho", "enviar"],
                   help=f"padrão: {config.MODO}")
    p.add_argument("--arquivo", help=f"planilha (padrão: {config.ARQUIVO})")
    p.add_argument("--loop", action="store_true", help="fica rodando de tempos em tempos")
    p.add_argument("--intervalo", type=int, default=config.INTERVALO_MIN,
                   help="minutos entre rodadas no --loop")
    p.add_argument("--diagnostico", action="store_true",
                   help="explica por que a planilha abre ou não, e sai")
    p.add_argument("--snapshot", action="store_true",
                   help="grava a foto diária para o Power BI e sai")
    a = p.parse_args(argv)

    _logging()
    arquivo = a.arquivo or config.ARQUIVO

    if a.diagnostico:
        print("\n".join(planilha.diagnosticar(arquivo)))
        return
    if a.snapshot:
        servico.gravar_snapshot(arquivo)
        return
    modo = a.modo.upper() if a.modo else config.MODO

    if not a.loop:
        _rodada(modo, arquivo)
        return
    log.info("Automação ativa (a cada %d min). Ctrl+C para parar.", a.intervalo)
    try:
        while True:
            _rodada(modo, arquivo)
            time.sleep(a.intervalo * 60)
    except KeyboardInterrupt:
        log.info("Automação encerrada.")
