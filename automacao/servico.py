"""Orquestra uma rodada completa: ler -> decidir -> enviar -> registrar -> métricas."""
import logging
from datetime import datetime

from automacao import regras
from automacao.canais import email_outlook
from automacao import config
from automacao import mensagens
from automacao import metricas
from automacao import planilha

log = logging.getLogger("automacao")


def _sufixo():
    return " (TESTE)" if config.EMAIL_TESTE else ""


def rodar_uma_vez(modo=None, arquivo=None, hoje=None, enviar_fn=None):
    """Executa uma rodada. `enviar_fn(destino, assunto, corpo)` é injetável (testes)."""
    modo = (modo or config.MODO).upper()
    arquivo = arquivo or config.ARQUIVO
    hoje = hoje or datetime.now()

    if modo == "ENVIAR" and not config.EMAIL_TESTE and not config.PERMITIR_ENVIO_REAL:
        raise planilha.ErroPlanilha(
            "Trava de segurança: modo ENVIAR sem AUTOMACAO_EMAIL_TESTE enviaria para os "
            "contatos REAIS da planilha. Defina AUTOMACAO_EMAIL_TESTE no .env "
            "(ou AUTOMACAO_ENVIO_REAL=1 se for intencional).")

    planilha.checar_leitura(arquivo)
    if modo == "ENVIAR":
        planilha.checar_escrita(arquivo)
    base, trat = planilha.carregar(arquivo)
    motivos = planilha.carregar_motivos(arquivo)

    acoes, resumo = regras.planejar(base, trat, hoje, config)
    log.info("Base: %d impressoras | sem monitoramento: %d", resumo["total_base"],
             resumo["sem_monitoramento"])
    for chave, rotulo in (("sem_email", "sem e-mail válido (ignoradas)"),
                          ("aguardando", "aguardando prazo"),
                          ("parados", "já em tratativa/encerradas")):
        if resumo[chave]:
            log.info("  - %d %s: %s", len(resumo[chave]), rotulo, ", ".join(resumo[chave]))

    envios = [a for a in acoes if a.tipo == "envio"][:config.LIMITE_EMAILS]
    escalar = [a for a in acoes if a.tipo == "escalar"]
    fila = envios + escalar
    log.info("Modo: %s | mensagens nesta rodada: %d%s", modo, len(fila),
             f" | tudo redirecionado para {config.EMAIL_TESTE}" if config.EMAIL_TESTE else "")
    if not fila:
        log.info("Nada a enviar.")
    elif modo == "ENVIAR":
        log.info("Backup criado: %s", planilha.backup(arquivo).name)

    enviar_fn = enviar_fn or (lambda d, a, c: email_outlook.enviar(
        d, a, c, modo, config.CONTA_REMETENTE))

    for acao in fila:
        seriais = [i.serial for i in acao.impressoras]
        if acao.tipo == "envio":
            destino = config.EMAIL_TESTE or acao.destinatario
            assunto = mensagens.assunto_envio(len(seriais), acao.numero)
            corpo = mensagens.montar_envio(acao.nome, acao.impressoras, acao.numero, motivos)
            status, obs = "Aguardando Retorno", f"{config.MARCADOR} Envio {acao.numero}{_sufixo()}"
            titulo = f"ENVIO {acao.numero} -> {destino} | {acao.nome}"
        else:
            destino = config.EMAIL_TESTE or config.EMAIL_HUMANO
            assunto = mensagens.assunto_escalonamento(len(seriais))
            corpo = mensagens.montar_escalonamento(acao.impressoras)
            status = "Encaminhado ao Humano"
            obs = (f"{config.MARCADOR} Escalado p/ humano após {acao.numero} envio(s)"
                   f"{_sufixo()}")
            titulo = f"ESCALONAR -> {destino or '(defina AUTOMACAO_EMAIL_HUMANO)'}"
        log.info("%s | séries: %s", titulo, ", ".join(seriais))

        if modo == "SIMULAR":
            log.info("\n%s\n%s", corpo, "-" * 60)
            continue
        if not destino:
            log.error("   sem destinatário configurado; pulado.")
            continue
        try:
            enviar_fn(destino, assunto, corpo)
            if modo == "ENVIAR":
                planilha.registrar(arquivo, [[s, hoje.strftime("%d/%m/%Y"), "Email",
                                              status, obs] for s in seriais])
                log.info("   enviado e registrado em Tratativas.")
            else:
                log.info("   salvo em Rascunhos.")
        except Exception as e:  # noqa: BLE001 - uma falha não pode derrubar a rodada
            log.error("   ERRO: %s: %s", type(e).__name__, e)

    if modo == "ENVIAR":
        gravar_snapshot(arquivo, hoje)
    return resumo


def gravar_snapshot(arquivo=None, hoje=None):
    arquivo = arquivo or config.ARQUIVO
    base, trat = planilha.carregar(arquivo)
    caminho = metricas.gravar(base, trat, hoje or datetime.now())
    log.info("Snapshot atualizado: %s", caminho)
