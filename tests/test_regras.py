"""Testes das regras. Rodar: python -m unittest discover -s tests -t ."""
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace

import pandas as pd

from automacao import servico
from automacao import mensagens
from automacao import metricas
from automacao import planilha
from automacao import regras

RAIZ = Path(__file__).resolve().parent.parent
HOJE = datetime(2026, 10, 10)
CFG = SimpleNamespace(STATUS_ALVO=("sem monitoramento",), STATUS_PARAR=(
    "respondido", "em tratativa", "concluido", "resolvido", "encaminhado ao humano"),
    REENVIO_DIAS=10, MAX_ENVIOS=2)


def h(dias, status, obs):
    return {"data": HOJE - timedelta(days=dias), "status": status, "obs": obs}


class TestDecidir(unittest.TestCase):
    def test_sem_historico_envia_1(self):
        self.assertEqual(regras.decidir([], HOJE, CFG)[:2], ("envio", 1))

    def test_aguarda_prazo(self):
        r = regras.decidir([h(3, "aguardando retorno", "[AUTO] Envio 1")], HOJE, CFG)
        self.assertEqual(r[0], "aguardar")

    def test_reenvio_apos_prazo(self):
        r = regras.decidir([h(12, "aguardando retorno", "[AUTO] Envio 1")], HOJE, CFG)
        self.assertEqual(r[:2], ("envio", 2))

    def test_escala_apos_segundo_envio_sem_resposta(self):
        hist = [h(25, "aguardando retorno", "[AUTO] Envio 1"),
                h(11, "aguardando retorno", "[AUTO] Envio 2")]
        self.assertEqual(regras.decidir(hist, HOJE, CFG)[:2], ("escalar", 2))

    def test_resposta_para_a_automacao(self):
        hist = [h(12, "aguardando retorno", "[AUTO] Envio 1"), h(2, "respondido", "M01")]
        self.assertEqual(regras.decidir(hist, HOJE, CFG)[0], "parar")

    def test_ja_escalada_nao_repete(self):
        hist = [h(30, "aguardando retorno", "[AUTO] Envio 2"),
                h(1, "aguardando retorno", "[AUTO] Escalado p/ humano")]
        self.assertEqual(regras.decidir(hist, HOJE, CFG)[0], "parar")

    def test_linha_manual_antiga_nao_bloqueia_envio_1(self):
        r = regras.decidir([h(5, "aguardando retorno", "ligou ontem")], HOJE, CFG)
        self.assertEqual(r[:2], ("envio", 1))


class TestPlanejar(unittest.TestCase):
    def setUp(self):
        self.base = pd.DataFrame({
            "Serial": ["A1", "A2", "B1", "C1", "D1"],
            "Status": ["Sem Monitoramento", "sem monitoramento ", "Sem Monitoramento",
                       "Com Monitoramento", "Sem Monitoramento"],
            "Endereço": ["r1", "r2", "r3", "r4", "r5"],
            "Número Para Contato": ["1", "1", "2", "3", "4"],
            "Email": ["Ana@x.com", "ana@x.com", "b@x.com", "c@x.com", ""],
            "Nome": ["Ana", "Ana", "Bia", "Cris", "Duda"],
        })
        self.trat = pd.DataFrame({"Serial": [], "Data Contato": [],
                                  "Status Tratativa": [], "Observação": []})

    def test_agrupa_por_contato_e_filtra_status_e_email(self):
        acoes, resumo = regras.planejar(self.base, self.trat, HOJE, CFG)
        self.assertEqual(resumo["sem_monitoramento"], 4)       # C1 (monitorada) fica de fora
        self.assertEqual(resumo["sem_email"], ["D1"])
        grupos = {a.destinatario: [i.serial for i in a.impressoras] for a in acoes}
        self.assertEqual(grupos, {"ana@x.com": ["A1", "A2"], "b@x.com": ["B1"]})


class TestMensagens(unittest.TestCase):
    def test_instrucao_com_serie_quando_ha_varias(self):
        imps = [regras.Impressora("A1", "r1", "Ana", "a@x"), regras.Impressora("A2", "r2", "Ana", "a@x")]
        corpo = mensagens.montar_envio("Ana", imps, 1)
        self.assertIn("série e o número da opção", corpo)
        self.assertIn("7- HP 4003dw perdeu configuração", corpo)

    def test_lembrete_no_envio_2(self):
        imps = [regras.Impressora("A1", "r1", "Ana", "a@x")]
        self.assertIn("Retomando", mensagens.montar_envio("Ana", imps, 2))
        self.assertTrue(mensagens.assunto_envio(1, 2).startswith("Lembrete"))


class TestPlanilhaEFluxo(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.arq = Path(self.tmp.name) / "base.xlsx"
        subprocess.run([sys.executable, str(RAIZ / "scripts" / "gerar_base_demo.py"),
                        str(self.arq)], check=True, capture_output=True)
        from automacao import config
        self._orig = (config.EMAIL_TESTE, config.EMAIL_HUMANO, config.SNAPSHOT_CSV)
        config.EMAIL_TESTE = "teste@exemplo.com"
        config.EMAIL_HUMANO = "humano@exemplo.com"
        config.SNAPSHOT_CSV = Path(self.tmp.name) / "snapshot.csv"   # não suja data/
        self.enviados = []

    def tearDown(self):
        from automacao import config
        config.EMAIL_TESTE, config.EMAIL_HUMANO, config.SNAPSHOT_CSV = self._orig
        self.tmp.cleanup()

    def _rodar(self, modo="ENVIAR"):
        servico.rodar_uma_vez(modo, self.arq, enviar_fn=lambda d, a, c: self.enviados.append((d, a)))

    def test_diagnostico_ok(self):
        texto = "\n".join(planilha.diagnosticar(self.arq))
        self.assertIn("formato .xlsx normal", texto)

    def test_diagnostico_arquivo_criptografado(self):
        falso = Path(self.tmp.name) / "cript.xlsx"
        falso.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"0" * 50)
        self.assertIn("CRIPTOGRAFADO", "\n".join(planilha.diagnosticar(falso)))
        with self.assertRaises(planilha.ErroPlanilha):
            planilha.checar_leitura(falso)

    def test_envio_registra_e_nao_repete(self):
        self._rodar()
        # demo: 2 envios 1 (Ana, Carlos) + 1 lembrete (Paula) + 1 escalonamento
        self.assertEqual(len(self.enviados), 4)
        self.assertTrue(all(d == "teste@exemplo.com" for d, _ in self.enviados))
        self._rodar()                       # segunda rodada no mesmo dia: nada novo
        self.assertEqual(len(self.enviados), 4)
        base, trat = planilha.carregar(self.arq)
        self.assertGreater(trat["Observação"].str.contains("[AUTO]", regex=False).sum(), 8)

    def test_snapshot_do_power_bi(self):
        base, trat = planilha.carregar(self.arq)
        linha = metricas.calcular(base, trat, HOJE)
        self.assertEqual(linha["total"], len(base))
        self.assertEqual(linha["sem_monitoramento"], 14)
        self.assertGreaterEqual(linha["escaladas"], 1)


if __name__ == "__main__":
    unittest.main()
