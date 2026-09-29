# ===================== CONFIGURAÇÃO =====================
MODO = "ENVIAR"         # "SIMULAR" | "RASCUNHO" | "ENVIAR"
DESTINO = "maria.camposprofeta@gmail.com"
NOME_DESTINATARIO = "Maria"
# ========================================================


def montar_mensagem(nome):
    saudacao = f"Olá, {nome}!" if isinstance(nome, str) and nome.strip() else "Olá!"
    return f"""{saudacao}

Identificamos que uma impressora Simpress está sem comunicação com a nossa central.

Você poderia nos informar se:

1 - Está em funcionamento
2 - Em reforma
3 - Foi substituída
4 - Precisa de suporte

Basta responder este e-mail com o número da opção.

Obrigado!
"""


def enviar_outlook(destino, assunto, corpo):
    import win32com.client  # só é necessário fora do modo SIMULAR
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To, mail.Subject, mail.Body = destino, assunto, corpo
    if MODO == "ENVIAR":
        mail.Send()
    else:
        mail.Save()  # vai para Rascunhos


def main():
    assunto = "Impressora Simpress sem comunicação"
    corpo = montar_mensagem(NOME_DESTINATARIO)

    print(f"Modo: {MODO} | Destino: {DESTINO}\n")
    if MODO == "SIMULAR":
        print(corpo)
        return

    enviar_outlook(DESTINO, assunto, corpo)
    print("E-mail salvo em Rascunhos." if MODO == "RASCUNHO" else "E-mail enviado.")


if __name__ == "__main__":
    main()