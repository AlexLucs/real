# -*- coding: utf-8 -*-
"""
Testes do atualizar_ipca.py (so biblioteca padrao, sem rede).

    python -m unittest discover -s testes -v

Rodam sobre copias temporarias da planilha, do modelo.html e do script; nada
no repositorio e alterado. Os dados de teste sao montados a partir da propria
planilha (ultimos 24 meses) mais dois meses ficticios, entao continuam validos
conforme a serie cresce. A rede e simulada trocando baixar_json/hoje_utc.
"""
import contextlib
import importlib.util
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ODS = "Desvalorização_Real_Gráfico.ods"

spec = importlib.util.spec_from_file_location("atualizar_ipca", os.path.join(RAIZ, "atualizar_ipca.py"))
A = importlib.util.module_from_spec(spec)
spec.loader.exec_module(A)
A.NO_ACTIONS = False            # testes nao geram anotacoes no Actions


def serie_planilha():
    return A.ler_serie(os.path.join(RAIZ, ODS))


def mes_mais(cod, k):
    a, m = A.mes_menos(cod // 100, cod % 100, -k)
    return a * 100 + m


def registros_sgs(serie, novos=(0.30, 0.50)):
    """Ultimos 24 meses da planilha (variacao implicita) + meses novos, formato SGS."""
    reg = []
    for i in range(len(serie) - 24, len(serie)):
        c, v = serie[i][0], round((serie[i][1] / serie[i - 1][1] - 1) * 100, 2)
        reg.append({"data": f"01/{c % 100:02d}/{c // 100}", "valor": f"{v:.2f}"})
    ult = serie[-1][0]
    for k, v in enumerate(novos, 1):
        c = mes_mais(ult, k)
        reg.append({"data": f"01/{c % 100:02d}/{c // 100}", "valor": f"{v:.2f}"})
    return reg


def formato_ibge(reg):
    return [{"resultados": [{"series": [{"serie": {
        r["data"][6:] + r["data"][3:5]: r["valor"] for r in reg}}]}]}]


class Base(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="ipca_teste_")
        for nome in (ODS, "modelo.html", "atualizar_ipca.py"):
            shutil.copy(os.path.join(RAIZ, nome), self.dir)
        self.ods = os.path.join(self.dir, ODS)
        self.serie = serie_planilha()
        self.reg = registros_sgs(self.serie)
        self._orig = (A.baixar_json, A.hoje_utc)
        ult = self.serie[-1][0]
        A.hoje_utc = lambda: (ult // 100, ult % 100, 1)   # "hoje" sem atraso

    def tearDown(self):
        A.baixar_json, A.hoje_utc = self._orig
        shutil.rmtree(self.dir, ignore_errors=True)

    def json(self, nome, dados):
        caminho = os.path.join(self.dir, nome)
        with open(caminho, "w", encoding="utf-8") as fh:
            json.dump(dados, fh)
        return caminho

    def cli(self, *args):
        env = {k: v for k, v in os.environ.items() if k != "GITHUB_ACTIONS"}
        r = subprocess.run([sys.executable, "-I", "atualizar_ipca.py", *args], cwd=self.dir,
                           capture_output=True, text=True, env=env)
        return r.returncode, r.stdout + r.stderr

    def bytes_ods(self):
        with open(self.ods, "rb") as fh:
            return fh.read()

    def em_processo(self, funcao, *args):
        """Roda uma funcao do script e devolve (codigo_de_saida, texto impresso)."""
        out = io.StringIO()
        codigo = 0
        with contextlib.redirect_stdout(out):
            try:
                funcao(*args)
            except SystemExit as e:
                codigo = 1 if e.code not in (0, None) else 0
                print(e.code)
        return codigo, out.getvalue()


class TestPlanilha(Base):
    def test_serie_basica(self):
        s = self.serie
        self.assertEqual(s[0], [A.MES_INICIAL, 915.93])
        self.assertEqual(A.lacunas([(c // 100, c % 100) for c, _ in s]), [])
        self.assertEqual(A.checar_indices(s), [])

    def test_base_reproduz_ibge(self):
        """Base jun/1994 = 857,29: julho/1994 = 6,84% e jul/1994-jun/2019 = 508,23%
        (acumulado do Plano Real divulgado pelo IBGE nos 25 anos do Real)."""
        d = dict(self.serie)
        self.assertEqual(round((d[199407] / A.BASE - 1) * 100, 2), 6.84)
        self.assertEqual(round((d[201906] / A.BASE - 1) * 100, 2), 508.23)

    def test_checar_indices_pega_problemas(self):
        ruim = [list(x) for x in self.serie]
        ruim[10][1] = float("nan")
        ruim[20][1] = ruim[19][1] * 1.5
        erros = A.checar_indices(ruim)
        self.assertTrue(any("invalido" in e for e in erros))
        self.assertTrue(any("implausivel" in e for e in erros))
        self.assertTrue(A.checar_indices(ruim[1:]))       # nao comeca em jul/1994


class TestCLI(Base):
    def test_atualizacao_normal_idempotente_e_verificada(self):
        f = self.json("normal.json", self.reg)
        cod, out = self.cli("--fonte", f, "--conferir-com", f)
        self.assertEqual(cod, 0, out)
        s = A.ler_serie(self.ods)
        self.assertEqual(len(s), len(self.serie) + 2)
        meta = zipfile.ZipFile(self.ods).read("meta.xml").decode()
        self.assertIn('IPCA_conferencia" meta:value-type="string">ok com', meta)
        antes = self.bytes_ods()
        cod, out = self.cli("--fonte", f)
        self.assertEqual(cod, 0, out)
        self.assertIn("Nada a fazer", out)
        self.assertEqual(self.bytes_ods(), antes)
        cod, out = self.cli("--verificar")
        self.assertEqual(cod, 0, out)
        with open(os.path.join(self.dir, "index.html"), encoding="utf-8") as fh:
            html = fh.read()
        self.assertIn(f"/*__SERIE__*/[[{A.MES_BASE},{A.BASE:g}],[{A.MES_INICIAL},915.93]", html)

    def test_entradas_invalidas_nao_gravam(self):
        sem_penultimo = self.reg[:-2] + self.reg[-1:]
        vazio = self.reg[:-2] + [dict(self.reg[-2], valor="")]
        absurdo = self.reg[:-2] + [dict(self.reg[-2], valor="48")]
        antes = self.bytes_ods()
        for nome, dados in (("lacuna", sem_penultimo), ("vazio", vazio), ("absurdo", absurdo),
                            ("so_ultimo", self.reg[-1:])):
            cod, out = self.cli("--fonte", self.json(nome + ".json", dados))
            self.assertEqual(cod, 1, f"{nome}: {out}")
            self.assertEqual(self.bytes_ods(), antes, nome)

    def test_conferencia_divergente_bloqueia(self):
        outra = [dict(r) for r in self.reg]
        outra[-1]["valor"] = "0.55"
        antes = self.bytes_ods()
        cod, out = self.cli("--fonte", self.json("a.json", self.reg),
                            "--conferir-com", self.json("b.json", outra))
        self.assertEqual(cod, 1, out)
        self.assertIn("As fontes divergem", out)
        self.assertEqual(self.bytes_ods(), antes)

    def test_verificar_detecta_pagina_desatualizada(self):
        self.assertEqual(self.cli("--so-html")[0], 0)
        caminho = os.path.join(self.dir, "index.html")
        with open(caminho, encoding="utf-8") as fh:
            html = fh.read()
        with open(caminho, "w", encoding="utf-8") as fh:
            fh.write(re.sub(r",\[\d{6},[\d.]+\]\];", "];", html, count=1))
        cod, out = self.cli("--verificar")
        self.assertEqual(cod, 1, out)


class TestOnline(Base):
    """BCB e IBGE simulados."""
    def rodar(self, bcb, ibge):
        def falso(url, tentativas=2, espera=10):
            r = bcb if "bcb.gov.br" in url else ibge
            if r is None:
                raise A.FalhaDownload("HTTP Error 503")
            return r
        A.baixar_json = falso
        return self.em_processo(A.atualizar, self.ods, self.ods)

    def conferencia(self):
        return re.findall(r'IPCA_conferencia"[^>]*>([^<]*)<',
                          zipfile.ZipFile(self.ods).read("meta.xml").decode())

    def test_fontes_iguais(self):
        cod, out = self.rodar(self.reg, formato_ibge(self.reg))
        self.assertEqual(cod, 0, out)
        self.assertEqual(self.conferencia(), ["ok com IBGE (tabela 1737)"])

    def test_ibge_diverge(self):
        outra = [dict(r) for r in self.reg]
        outra[-1]["valor"] = "0.55"
        cod, out = self.rodar(self.reg, formato_ibge(outra))
        self.assertEqual(cod, 1, out)
        self.assertEqual(len(A.ler_serie(self.ods)), len(self.serie))

    def test_ibge_fora(self):
        cod, out = self.rodar(self.reg, None)
        self.assertEqual(cod, 0, out)
        self.assertEqual(self.conferencia(), ["sem conferencia (IBGE (tabela 1737) indisponivel)"])

    def test_bcb_fora_usa_reserva(self):
        cod, out = self.rodar(None, formato_ibge(self.reg))
        self.assertEqual(cod, 0, out)
        self.assertIn("Usada a reserva", out)
        self.assertEqual(len(A.ler_serie(self.ods)), len(self.serie) + 2)

    def test_bcb_atrasado_em_relacao_ao_ibge(self):
        cod, out = self.rodar(self.reg[:-2], formato_ibge(self.reg))
        self.assertEqual(cod, 0, out)
        self.assertIn("possivelmente desatualizada", out)
        self.assertIn("Nada a fazer", out)

    def test_serie_atrasada(self):
        ult = self.serie[-1][0]
        a, m = A.mes_menos(ult // 100, ult % 100, -6)
        A.hoje_utc = lambda: (a, m, 20)
        cod, out = self.rodar(self.reg, formato_ibge(self.reg))
        self.assertEqual(cod, 0, out)
        self.assertIn("Serie atrasada", out)

    def test_divergencia_em_mes_antigo_so_avisa(self):
        outra = [dict(r) for r in self.reg]
        outra[3]["valor"] = f"{float(outra[3]['valor']) + 0.01:.2f}"
        cod, out = self.rodar(self.reg, formato_ibge(outra))
        self.assertEqual(cod, 0, out)
        self.assertIn("divergem em meses ja gravados", out)


class TestConferenciaODS(Base):
    def test_verificar_ods_pega_expectativa_errada(self):
        f = self.json("normal.json", self.reg)
        self.assertEqual(self.cli("--fonte", f, "--sem-html")[0], 0)
        with zipfile.ZipFile(os.path.join(RAIZ, ODS)) as z:
            antes = A.contagens({n: z.read(n) for n in z.namelist()}, "Object 1/content.xml")
        s = A.ler_serie(self.ods)
        ultima = A.PRIMEIRA_LINHA + len(s) - 1
        self.assertEqual(A.verificar_ods(self.ods, antes, 2, tuple(s[-1]), ultima,
                                         "Object 1/content.xml"), [])
        self.assertTrue(A.verificar_ods(self.ods, antes, 3, tuple(s[-1]), ultima + 1,
                                        "Object 1/content.xml"))


class TestAuditoria(Base):
    def oficial(self, ajuste=None):
        alvo = A.serie_pagina(self.serie)
        serie = {str(c): f"{v:.2f}" for c, v in alvo}
        if ajuste:
            c, delta = ajuste
            serie[str(c)] = f"{float(serie[str(c)]) + delta:.2f}"
        return [{"resultados": [{"series": [{"serie": serie}]}]}]

    def auditar(self, resposta):
        def falso(url, tentativas=2, espera=10):
            self.assertIn("variaveis/2266", url)
            if resposta is None:
                raise A.FalhaDownload("HTTP Error 503")
            return resposta
        A.baixar_json = falso
        return self.em_processo(A.auditar, self.ods)

    def test_igual_ao_oficial(self):
        cod, out = self.auditar(self.oficial())
        self.assertEqual(cod, 0, out)
        self.assertIn("Auditoria ok", out)

    def test_divergencia(self):
        cod, out = self.auditar(self.oficial((200001, 0.05)))
        self.assertEqual(cod, 1, out)
        self.assertIn("janeiro 2000", out)

    def test_base_divergente(self):
        cod, out = self.auditar(self.oficial((A.MES_BASE, 0.10)))
        self.assertEqual(cod, 1, out)

    def test_ibge_fora_nao_falha(self):
        cod, out = self.auditar(None)
        self.assertEqual(cod, 0, out)
        self.assertIn("Auditoria nao realizada", out)


if __name__ == "__main__":
    unittest.main()
