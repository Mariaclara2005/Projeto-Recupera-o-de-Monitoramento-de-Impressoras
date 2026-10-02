"""Configuração central do projeto.

Valores pessoais ou sensíveis (e-mails, modo de envio) vêm de variáveis de
ambiente ou do arquivo `.env` na raiz do projeto. O `.env` NÃO vai para o Git;
use o `.env.example` como modelo.
"""
import os
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent


def _carregar_env():
    arquivo = RAIZ / ".env"
    if not arquivo.exists():
        return
    for linha in arquivo.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = linha.split("=", 1)
        os.environ.setdefault(chave.strip(), valor.strip().strip("\"").strip("\x27"))


_carregar_env()


def _env(nome, padrao=""):
    return os.environ.get(nome, padrao).strip()


def _caminho(valor, padrao):
    p = Path(valor) if valor else padrao
    return p if p.is_absolute() else RAIZ / p


# ----------------------------- Arquivos -----------------------------
ARQUIVO = _caminho(_env("AUTOMACAO_ARQUIVO"), RAIZ / "data" / "base_demo.xlsx")
PASTA_LOGS = RAIZ / "logs"
SNAPSHOT_CSV = RAIZ / "data" / "snapshot_diario.csv"   # alimenta o Power BI
ABA_BASE = "Planilha1"
ABA_TRAT = "Tratativas"
ABA_MOTIVOS = "Motivos"

# ------------------------------ Envio -------------------------------
# SIMULAR = só mostra | RASCUNHO = salva no Outlook | ENVIAR = envia de verdade
MODO = _env("AUTOMACAO_MODO", "SIMULAR").upper()
EMAIL_TESTE = _env("AUTOMACAO_EMAIL_TESTE")      # se preenchido, TUDO vai para ele
EMAIL_HUMANO = _env("AUTOMACAO_EMAIL_HUMANO")    # quem recebe os casos escalados
CONTA_REMETENTE = _env("AUTOMACAO_CONTA_REMETENTE")   # conta do Outlook (opcional)
# Trava: ENVIAR sem EMAIL_TESTE só funciona com AUTOMACAO_ENVIO_REAL=1
PERMITIR_ENVIO_REAL = _env("AUTOMACAO_ENVIO_REAL") == "1"

# ------------------------------ Regras ------------------------------
STATUS_ALVO = ("sem monitoramento",)   # acrescente "perdeu monitoramento", "parou recentemente"
REENVIO_DIAS = int(_env("AUTOMACAO_REENVIO_DIAS", "10"))   # espera entre tentativas
MAX_ENVIOS = int(_env("AUTOMACAO_MAX_ENVIOS", "2"))        # 1º envio + lembretes
LIMITE_EMAILS = int(_env("AUTOMACAO_LIMITE_EMAILS", "10"))  # trava por execução
INTERVALO_MIN = int(_env("AUTOMACAO_INTERVALO_MIN", "2"))  # modo contínuo

# Último status da tratativa que PARA a automação (alguém já cuidou do caso)
STATUS_PARAR = ("respondido", "em tratativa", "concluido", "resolvido",
                "encaminhado ao humano")
MARCADOR = "[AUTO]"   # marca, na aba Tratativas, o que foi feito pela automação
