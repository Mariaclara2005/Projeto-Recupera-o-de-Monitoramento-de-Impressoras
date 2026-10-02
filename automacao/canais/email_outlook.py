"""Canal de e-mail via Outlook clássico (pywin32). Só funciona no Windows com o
Outlook desktop instalado. Importado apenas quando realmente vai enviar."""


def enviar(destino, assunto, corpo, modo, conta=""):
    """modo 'ENVIAR' envia; qualquer outro valor salva em Rascunhos."""
    import win32com.client  # noqa: PLC0415 - import tardio de propósito

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To, mail.Subject, mail.Body = destino, assunto, corpo
    if conta:
        for c in outlook.Session.Accounts:
            if c.SmtpAddress.lower() == conta.lower():
                mail.SendUsingAccount = c
                break
        else:
            raise RuntimeError(f"Conta '{conta}' não encontrada no Outlook.")
    if modo == "ENVIAR":
        mail.Send()
    else:
        mail.Save()
