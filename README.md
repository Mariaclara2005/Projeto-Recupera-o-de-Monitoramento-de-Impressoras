# Recuperação de Monitoramento de Impressoras (Simpress) — DEMO

Automação que identifica impressoras **sem monitoramento**, contata o responsável
pela unidade e acompanha a resposta, escalando para uma pessoa quando não há retorno.
Os números gerados alimentam o Power BI (antes × depois da automação).

> **Status: demonstração.** Usa dados fictícios. Antes de usar dados reais:
> aprovação do gestor/TI, revisão de LGPD, caixa de e-mail corporativa e execução
> em servidor (ver "Próximos passos").

## Fluxo (conforme o desenho do projeto)

```
Planilha de controle → Python lê → Filtra "Sem Monitoramento" → Busca contato
→ Mensagem automática → Atualiza Tratativas → Aguarda retorno
   ├─ respondeu → humano trata o caso (status "Respondido / Em Tratativa")
   ├─ sem resposta em 10 dias → Lembrete (Envio 2)
   └─ sem resposta após o lembrete + 10 dias → Encaminha para HUMANO
```

## Estrutura de pastas

```
automacao/            código (um arquivo por responsabilidade)
  config.py           configurações (lê o .env)
  planilha.py         ler/gravar Excel + diagnóstico de erros de acesso
  regras.py           REGRAS: quem contatar, reenvio, escalonamento
  mensagens.py        textos dos e-mails
  canais/email_outlook.py   envio pelo Outlook (WhatsApp entrará em canais/)
  metricas.py         foto diária para o Power BI
  servico.py          orquestra uma rodada completa
  cli.py              linha de comando
scripts/gerar_base_demo.py   cria a planilha fictícia de teste
tests/                testes automatizados
data/                 planilhas e CSVs (fora do Git)
logs/                 histórico de execução (fora do Git)
run.py                ponto de entrada
```

## Como rodar (Windows, PowerShell, na raiz do projeto)

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe scripts\gerar_base_demo.py     # cria data\base_demo.xlsx
.\.venv\Scripts\python.exe run.py                         # SIMULAR (não envia nada)
```

Copie `.env.example` para `.env` e preencha `AUTOMACAO_EMAIL_TESTE` (seu e-mail).
Depois suba de nível: `--modo rascunho` → `--modo enviar`.

| Comando | O que faz |
|---|---|
| `python run.py` | Uma rodada em modo SIMULAR |
| `python run.py --modo rascunho` | Salva os e-mails em Rascunhos do Outlook |
| `python run.py --modo enviar` | Envia e registra em Tratativas (exige `AUTOMACAO_EMAIL_TESTE`) |
| `python run.py --loop --intervalo 15` | Fica rodando, verificando a cada 15 min |
| `python run.py --diagnostico` | Explica por que a planilha abre (ou não) |
| `python run.py --snapshot` | Atualiza `data/snapshot_diario.csv` (Power BI) |
| `python -m unittest discover -s tests -t .` | Roda os testes |

## Planilha esperada

Abas e colunas (a ordem das colunas não importa; os nomes são encontrados sem
diferenciar maiúsculas/acentos):

- **Planilha1**: Serial, Status, Data Identificação, Endereço, Número Para Contato, Email, Nome, Cargo
- **Tratativas**: Serial, Data Contato, Canal, Status Tratativa, Observação
- **Motivos**: Código, Motivo (a ordem define as opções 1, 2, 3… enviadas ao cliente)

A automação marca o que enviou com `[AUTO]` na coluna Observação. É assim que ela
sabe, nas próximas rodadas, quantas tentativas já foram feitas.
Para **parar** a cadência de uma impressora, registre em Tratativas um status como
`Respondido`, `Em Tratativa`, `Concluído` ou `Resolvido`.

## Segurança e LGPD (demo)

- Trava: `ENVIAR` só funciona com `AUTOMACAO_EMAIL_TESTE` preenchido (ou com `AUTOMACAO_ENVIO_REAL=1`).
- Backup automático da planilha antes de cada gravação.
- Dados reais, `.env` e logs ficam fora do Git (`.gitignore`).
- Para produção: base legal e opt-out, minimização dos dados no e-mail, controle de
  acesso, retenção dos logs, e-mail corporativo compartilhado.

## Próximos passos

1. Ler as respostas (1–7) da caixa de entrada e atualizar o status sozinho.
2. Canal WhatsApp (API oficial) em `automacao/canais/whatsapp.py`.
3. Planilha no SharePoint/lista + execução em servidor ou Power Automate.
4. Power BI: abas *Monitoramento Atual / Resultados Esperados / Pós Automatização*
   usando `data/snapshot_diario.csv` e a aba Tratativas.
