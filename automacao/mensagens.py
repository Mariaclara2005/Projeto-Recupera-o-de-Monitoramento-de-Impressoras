"""Textos das mensagens. Alterar redação = mexer só neste arquivo."""

# Usadas se a aba "Motivos" da planilha não existir. Mantenha a ordem: a opção 1
# do cliente = M01, 2 = M02... (assim a resposta vira o código do motivo).
MOTIVOS_PADRAO = [
    ("M01", "Loja em reforma"), ("M02", "Impressora trocada"), ("M03", "Sem rede"),
    ("M04", "Perda IP"), ("M05", "Sem energia"), ("M06", "Equipamento removido"),
    ("M07", "HP 4003dw perdeu configuração"),
]


def assunto_envio(qtd, numero):
    base = ("Impressora Simpress sem comunicação" if qtd == 1
            else f"Impressoras Simpress sem comunicação ({qtd})")
    return base if numero == 1 else f"Lembrete: {base}"


def montar_envio(nome, impressoras, numero=1, motivos=None):
    motivos = motivos or MOTIVOS_PADRAO
    saudacao = f"Olá, {nome}!😊" if nome else "Olá!😊"
    if len(impressoras) == 1:
        i = impressoras[0]
        intro = (f"Identificamos que uma impressora Simpress (série {i.serial}), "
                 f"localizada em {i.endereco}, está sem comunicação com a nossa central.")
        instrucao = "Basta responder este e-mail com o número da opção."
    else:
        lista = "\n".join(f"  - {i.serial} | {i.endereco}" for i in impressoras)
        intro = ("Identificamos que as impressoras Simpress abaixo estão sem "
                 f"comunicação com a nossa central:\n\n{lista}")
        instrucao = ("Basta responder este e-mail com a série e o número da opção "
                     f"(ex.: {impressoras[0].serial} - 3).")
    lembrete = ("" if numero == 1 else
                "Retomando o nosso contato anterior: ainda não recebemos retorno.\n\n")
    opcoes = "\n\n".join(f"{n}- {texto}" for n, (_, texto) in enumerate(motivos, 1))
    return (f"{saudacao}\n\n{lembrete}{intro}\n\n"
            f"🧐 Você poderia nos informar se :\n\n{opcoes}\n\n"
            f"{instrucao}\n\nObrigado!🖨️\n")


def assunto_escalonamento(qtd):
    return f"Ação necessária: {qtd} impressora(s) sem retorno do cliente"


def montar_escalonamento(impressoras):
    linhas = []
    for i in impressoras:
        contato = ", ".join(x for x in (i.nome, i.email, i.telefone) if x)
        linhas.append(f"  - {i.serial} | {i.endereco}\n"
                      f"    Contato: {contato}\n"
                      f"    Tentativas automáticas sem resposta: {i.envios}")
    return ("Olá!\n\nAs impressoras abaixo continuam SEM monitoramento e o contato "
            "não respondeu às mensagens automáticas. É preciso tratar o caso "
            "manualmente (ligação, visita ou outro canal):\n\n"
            + "\n".join(linhas) + "\n\nMensagem gerada pela automação de recuperação "
            "de monitoramento.\n")
