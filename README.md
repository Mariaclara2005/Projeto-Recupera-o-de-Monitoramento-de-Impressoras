# Projeto-Recupera-o-de-Monitoramento-de-Impressoras

## Envio de e-mail pelo Gmail

O script `Automacao.py` envia e-mails pelo Gmail usando SMTP com conexão TLS. Ele
não precisa do Outlook nem de `pywin32`.

1. Na conta Google remetente, ative a verificação em duas etapas e crie uma
   [senha de app](https://myaccount.google.com/apppasswords).
2. Configure as variáveis de ambiente `GMAIL_ADDRESS` (endereço Gmail remetente)
   e `GMAIL_APP_PASSWORD` (senha de app). Não coloque a senha no código nem a
   compartilhe.
3. Execute `python Automacao.py`. O modo `ENVIAR` envia a mensagem para o
   endereço definido em `DESTINO`.

Para configurar as variáveis no PowerShell antes de executar:

```powershell
$env:GMAIL_ADDRESS = "seu-endereco@gmail.com"
$env:GMAIL_APP_PASSWORD = "sua-senha-de-app"
python .\Automacao.py
```

O modo `RASCUNHO` abre a composição de uma mensagem no Gmail no navegador
padrão; o modo `SIMULAR` apenas exibe a mensagem no terminal.