#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
atualizar_ipca.py — atualiza a planilha "Desvalorização_Real_Gráfico.ods" com os
dados mais recentes do IPCA, preservando a formatacao, as formulas e o grafico.

Fonte dos dados
---------------
1) Banco Central do Brasil - SGS, serie 433 = "IPCA - variacao mensal (%)".
       https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados
2) Reserva, se o BCB falhar: IBGE, API de agregados, tabela 1737,
   variavel 63 = "IPCA - Variacao mensal" (o mesmo dado, na origem).
       https://servicodados.ibge.gov.br/api/v3/agregados/1737/...

Pede-se so a janela recente (24 meses antes do ultimo mes da planilha em
diante), com cabecalhos que proibem resposta de cache: em nov/2026 a API do
BCB passou a responder "304 Not Modified" a requisicoes sem cabecalhos.

Numero-indice e regra de arredondamento
---------------------------------------
O numero-indice da coluna B (base dez/1993 = 100, mesma da tabela 1737 do
SIDRA/IBGE) e reconstruido encadeando as variacoes mensais a partir do ultimo
valor ja existente na planilha:

    indice_novo = arredonda(indice_anterior * (1 + variacao/100), 2)

O arredondamento a cada mes e proposital: a planilha guarda o indice com 2
casas, como o SIDRA, e cada mes novo parte do valor gravado. O erro dominante
nao e esse (no maximo 0,005 em ~7.700, ou 1e-6 relativo), e sim a propria
variacao publicada com 2 casas (+-0,005 p.p., ou 5e-5 relativo por mes), que
nenhuma precisao interna recupera. Diferencas de +-0,01 em relacao ao indice
oficial do SIDRA sao esperadas e aceitas.

Checagens
---------
- Cada registro baixado precisa ter data e valor numerico, com variacao
  plausivel (|v| <= 10%), sem meses repetidos nem lacunas. Uma fonte que falhe
  nisso e descartada (passa-se a reserva); com --fonte local, o script para.
- Os meses novos precisam comecar no mes seguinte ao ultimo da planilha.
- Conferencia cruzada: a outra fonte (IBGE, se o BCB foi usado) tambem e
  consultada. Se ela trouxer, para algum mes novo, uma variacao diferente
  (as duas publicam o mesmo dado com 2 casas, entao qualquer diferenca conta),
  nada e gravado. Se a outra fonte estiver fora do ar ou ainda sem o mes,
  segue-se com aviso.
- Meses que a planilha ja tem sao comparados com a variacao baixada; se o
  BCB/IBGE tiver revisado algum, sai um aviso (a planilha nao e alterada).
- O .ods novo e gravado num temporario e conferido (XML bem-formado, serie
  relida, formulas e valores das colunas C e D das linhas novas, caixa de
  texto com o mes e o valor final, linhas, intervalos e pontos do grafico)
  antes de substituir o original. Se algo nao bater, o original fica intacto.
- No GitHub Actions, avisos e erros tambem saem como anotacoes
  (::warning:: / ::error::), que aparecem no resumo do run mesmo quando ele
  termina verde.
- A data e a fonte da coleta vao para as propriedades do documento
  (meta.xml: IPCA_coleta, IPCA_fonte) e dali para a pagina.

Alem da planilha, o script regrava uma pagina HTML interativa
("index.html") a partir do arquivo "modelo.html", que fica ao
lado do script. O modelo e um HTML comum: edite o visual nele a vontade; o
script so toca os marcadores /*__SERIE__*/ (a serie) e /*__META__*/ (data e
fonte da coleta). Sem o modelo.html por perto, o script apenas avisa e segue.

Uso
---
    python3 atualizar_ipca.py                      # atualiza planilha + HTML
    python3 atualizar_ipca.py -e entrada.ods -s saida.ods
    python3 atualizar_ipca.py --fonte dados.json   # usa um JSON local (offline)
    python3 atualizar_ipca.py --simular            # so mostra o que falta
    python3 atualizar_ipca.py --so-html            # so regera o HTML
    python3 atualizar_ipca.py --sem-html           # so mexe na planilha
    python3 atualizar_ipca.py --verificar          # confere planilha x index.html
    python3 atualizar_ipca.py --fonte a.json --conferir-com b.json
                                                   # testes: conferencia com JSON local

O script e idempotente: rodar duas vezes seguidas nao duplica nada.
"""

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from xml.sax.saxutils import escape

URL_BCB = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados"
URL_IBGE = ("https://servicodados.ibge.gov.br/api/v3/agregados/1737/"
            "periodos/{periodos}/variaveis/63?localidades=N1%5Ball%5D")
CABECALHOS = {
    "User-Agent": "Mozilla/5.0 (compatible; atualizar_ipca/1.2)",
    "Accept": "application/json",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
}
JANELA_MESES = 24      # quantos meses antes do ultimo da planilha pedir de novo
LIMITE_VARIACAO = 10.0  # |variacao mensal| acima disso (em %) e tida como erro
TOLERANCIA_REVISAO = 0.015  # p.p.; a variacao implicita na planilha fica a
                            # menos de 0,001 p.p. da publicada

MESES = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho",
         "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
# nomes exatamente como aparecem na planilha (com acento)
MESES_PT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

BASE = 915.93          # indice de julho/1994 (celula B3) = base do grafico
MODELO = "modelo.html"          # template da pagina (ao lado do script)
PAGINA = "index.html"  # pagina gerada (servida pelo GitHub Pages)
PRIMEIRA_LINHA = 3     # linha da planilha onde comeca a serie
MARCA_SERIE = "/*__SERIE__*/"
MARCA_META = "/*__META__*/"
PROP_COLETA, PROP_FONTE = "IPCA_coleta", "IPCA_fonte"   # propriedades no meta.xml

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "chart": "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
}
FONTE_BCB, FONTE_IBGE = "Banco Central (SGS 433)", "IBGE (tabela 1737)"


def q(nome):
    """'table:table-row' -> '{urn...table:1.0}table-row' (nome para o ElementTree)."""
    pre, local = nome.split(":")
    return "{%s}%s" % (NS[pre], local)


# --------------------------------------------------------------------------
# utilidades
# --------------------------------------------------------------------------
def num(x):
    """Formata um float como o LibreOffice grava em office:value."""
    return f"{x:.15g}"


def br(texto):
    return texto.replace(".", ",")


def rotulo(ano, mes):
    return f"{MESES_PT[mes - 1]} {ano}"


def rotulo_titulo(ano, mes):
    r = MESES_PT[mes - 1]
    return f"{r[0].upper()}{r[1:]} {ano}"


NO_ACTIONS = os.environ.get("GITHUB_ACTIONS") == "true"


def anotacao(tipo, msg):
    """Anotacao do GitHub Actions (aparece no resumo do run). Fora do Actions, nada."""
    if NO_ACTIONS:
        txt = msg.replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")
        print(f"::{tipo} title=Atualizar IPCA::{txt}", flush=True)


def aviso(msg):
    print(f"Aviso: {msg}", flush=True)
    anotacao("warning", msg)


def falhar(msg):
    """Encerra com codigo 1, deixando a causa visivel tambem no resumo do Actions."""
    anotacao("error", msg)
    sys.exit(msg)


def mes_seguinte(ano, mes):
    return (ano + 1, 1) if mes == 12 else (ano, mes + 1)


def mes_menos(ano, mes, k):
    """(ano, mes) recuado k meses."""
    t = ano * 12 + (mes - 1) - k
    return t // 12, t % 12 + 1


def lacunas(meses):
    """Recebe [(ano, mes), ...] ordenado e devolve as quebras de sequencia
    como [((ano, mes) esperado, (ano, mes) encontrado), ...]."""
    erros = []
    for anterior, atual in zip(meses, meses[1:]):
        if atual != mes_seguinte(*anterior):
            erros.append((mes_seguinte(*anterior), atual))
    return erros


class FalhaDownload(Exception):
    """Uma fonte nao respondeu com dados validos."""


def baixar_json(url, tentativas=2, espera=10):
    """Baixa e decodifica um JSON, tentando de novo em caso de falha."""
    req = urllib.request.Request(url, headers=CABECALHOS)
    ultimo_erro = None
    for n in range(1, tentativas + 1):
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as e:
            # urllib trata 304 (e qualquer nao-2xx) como HTTPError, sem corpo
            ultimo_erro = e
            print(f"  tentativa {n}/{tentativas} falhou: {e}")
            if n < tentativas:
                time.sleep(espera)
    raise FalhaDownload(str(ultimo_erro))


def ler_bcb(dados):
    """Formato SGS: [{"data": "01/09/2026", "valor": "0.48"}, ...]"""
    if not isinstance(dados, list):
        raise ValueError("a resposta nao e uma lista de registros")
    serie = []
    for item in dados:
        data = item.get("data") if isinstance(item, dict) else None
        valor = item.get("valor") if isinstance(item, dict) else None
        m = re.fullmatch(r"\d{2}/(\d{2})/(\d{4})", str(data or "").strip())
        if not m or not 1 <= int(m.group(1)) <= 12:
            raise ValueError(f"data invalida no registro {item!r}")
        try:
            v = float(valor)
        except (TypeError, ValueError):
            raise ValueError(f"valor nao numerico em {data}: {valor!r}") from None
        serie.append((int(m.group(2)), int(m.group(1)), v))
    return serie


def ler_ibge(dados):
    """Formato IBGE v3: [{"resultados": [{"series": [{"serie": {"202609": "0.48"}}]}]}]"""
    if not isinstance(dados, list) or not dados:
        raise ValueError("a resposta nao e uma lista com resultados")
    serie = []
    for periodo, valor in dados[0]["resultados"][0]["series"][0]["serie"].items():
        try:
            v = float(valor)
        except (TypeError, ValueError):
            continue        # "..." ou "-": mes ainda sem dado
        if not re.fullmatch(r"\d{4}(0[1-9]|1[0-2])", periodo):
            raise ValueError(f"periodo invalido: {periodo!r}")
        serie.append((int(periodo[:4]), int(periodo[4:]), v))
    return serie


def checar_serie(serie):
    """Levanta ValueError se a serie baixada tiver valor implausivel, mes
    repetido ou lacuna. Recebe a serie ja ordenada."""
    for a, me, v in serie:
        if v != v or abs(v) > LIMITE_VARIACAO:      # v != v pega NaN
            raise ValueError(f"variacao implausivel em {rotulo(a, me)}: {v}%")
    meses = [(a, me) for a, me, _ in serie]
    if len(set(meses)) != len(meses):
        raise ValueError("mes repetido na resposta")
    erros = lacunas(meses)
    if erros:
        esperado, achado = erros[0]
        raise ValueError(f"lacuna na serie: depois de {rotulo(*mes_menos(*esperado, 1))} "
                         f"veio {rotulo(*achado)}")


def ler_arquivo(caminho):
    """JSON local no formato do SGS, ja conferido. Para nele qualquer problema."""
    try:
        with open(caminho, encoding="utf-8") as fh:
            serie = sorted(ler_bcb(json.load(fh)))
        checar_serie(serie)
    except (ValueError, OSError) as e:
        falhar(f"Arquivo {caminho} invalido: {e}")
    return serie


def fontes_online(desde):
    """[(nome, url, leitor), ...] na ordem de preferencia: BCB, depois IBGE."""
    a, m = desde
    url_bcb = (f"{URL_BCB}?formato=json&dataInicial=01/{m:02d}/{a}"
               f"&dataFinal={time.strftime('%d/%m/%Y')}")
    return [(FONTE_BCB, url_bcb, ler_bcb),
            (FONTE_IBGE, URL_IBGE.format(periodos=f"-{JANELA_MESES + 12}"), ler_ibge)]


def tentar_fonte(nome, url, ler):
    """Baixa, le e confere uma fonte. Devolve a serie ordenada ou None."""
    print(f"Baixando de {nome}...", flush=True)
    try:
        serie = sorted(ler(baixar_json(url)))
    except FalhaDownload as e:
        aviso(f"{nome} nao respondeu: {e}")
        return None
    except (KeyError, IndexError, TypeError, AttributeError, ValueError) as e:
        aviso(f"{nome}: resposta em formato inesperado ({type(e).__name__}: {e})")
        return None
    if not serie:
        aviso(f"{nome}: resposta vazia")
        return None
    try:
        checar_serie(serie)
    except ValueError as e:
        aviso(f"{nome}: dados recusados ({e})")
        return None
    return serie


def baixar_serie(fonte, desde, conferir_com=None):
    """Devolve (serie, nome_da_fonte, serie_de_conferencia, nome_da_conferencia).

    serie: [(ano, mes, variacao_percentual), ...] ordenado, da fonte principal.
    A conferencia vem da outra fonte (None se indisponivel ou nao aplicavel).

    fonte: None (BCB com reserva no IBGE), URL no formato do SGS ou caminho
    de um JSON local no formato do SGS (estes dois desligam reserva e
    conferencia online). conferir_com: JSON local no formato do SGS usado
    como conferencia (para testes). desde: (ano, mes) inicial da consulta.
    """
    conf = (ler_arquivo(conferir_com), f"arquivo local ({os.path.basename(conferir_com)})") \
        if conferir_com else (None, None)

    if fonte:
        if fonte.startswith(("http://", "https://")):
            serie = tentar_fonte("URL informada", fonte, ler_bcb)
            if serie is None:
                falhar(f"Nao foi possivel obter a serie em {fonte}.")
            return serie, fonte, *conf
        return ler_arquivo(fonte), f"arquivo local ({os.path.basename(fonte)})", *conf

    fontes = fontes_online(desde)
    for i, (nome, url, ler) in enumerate(fontes):
        serie = tentar_fonte(nome, url, ler)
        if serie is None:
            continue
        if i > 0:
            aviso(f"Usada a reserva ({nome}): {fontes[0][0]} falhou. "
                  "Nao ha segunda fonte para a conferencia cruzada.")
            return serie, nome, *conf
        if conferir_com:
            return serie, nome, *conf
        # conferencia: a outra fonte, consultada sempre
        n2, u2, l2 = fontes[1]
        return serie, nome, tentar_fonte(n2, u2, l2), n2
    falhar("Nao foi possivel obter a serie do IPCA (BCB e IBGE falharam).")


def conferir_fontes(novos, conferencia, nome_conf):
    """Compara os meses novos com a segunda fonte. Divergencia encerra o script;
    falta da segunda fonte (ou do mes nela) so gera aviso."""
    if conferencia is None:
        aviso("Conferencia cruzada indisponivel; meses novos gravados com uma fonte so.")
        return
    outra = {(a, me): v for a, me, v in conferencia}
    divergentes, sem_dado = [], []
    for a, me, v in novos:
        if (a, me) not in outra:
            sem_dado.append(rotulo(a, me))
        elif round(abs(outra[(a, me)] - v), 6) > 0:
            divergentes.append(f"{rotulo(a, me)}: {v:+.2f}% x {outra[(a, me)]:+.2f}%")
    if divergentes:
        falhar(f"As fontes divergem ({nome_conf}): " + "; ".join(divergentes)
               + ". Nada foi gravado.")
    if sem_dado:
        aviso(f"{nome_conf} ainda sem " + ", ".join(sem_dado)
              + "; esses meses nao tiveram conferencia cruzada.")
    else:
        print(f"Conferencia cruzada ok com {nome_conf} ({len(novos)} mes(es) novo(s)).")


def checar_revisoes(serie_baixada, serie_planilha):
    """Compara a variacao baixada com a implicita na planilha nos meses que as
    duas tem. Devolve [(AAAAMM, variacao_planilha, variacao_baixada), ...]."""
    indice = dict(serie_planilha)
    divergentes = []
    for a, me, v in serie_baixada:
        cod = a * 100 + me
        pa, pm = mes_menos(a, me, 1)
        ant = pa * 100 + pm
        if cod in indice and ant in indice:
            implicita = (indice[cod] / indice[ant] - 1) * 100
            if abs(implicita - v) > TOLERANCIA_REVISAO:
                divergentes.append((cod, implicita, v))
    return divergentes


# --------------------------------------------------------------------------
# blocos XML
# --------------------------------------------------------------------------
LINHA_PLANILHA = (
    '<table:table-row table:style-name="ro1">'
    '<table:table-cell table:style-name="ce1" office:value-type="string" '
    'calcext:value-type="string"><text:p>{mes}</text:p></table:table-cell>'
    '<table:table-cell table:style-name="ce1" office:value-type="float" '
    'office:value="{idx}" calcext:value-type="float">'
    '<text:p>{idx_br}</text:p></table:table-cell>'
    '<table:table-cell table:formula="of:=[.B{lin}]/{base}" '
    'office:value-type="float" office:value="{c}" calcext:value-type="float">'
    '<text:p>{c_br}</text:p></table:table-cell>'
    '<table:table-cell table:formula="of:=(1/[.C{lin}])" '
    'office:value-type="percentage" office:value="{d}" '
    'calcext:value-type="percentage"><text:p>{d_br}</text:p></table:table-cell>'
    '<table:table-cell table:style-name="Default" '
    'table:number-columns-repeated="4"/></table:table-row>'
)

LINHA_GRAFICO = (
    '<table:table-row>'
    '<table:table-cell office:value-type="string">'
    '<text:p>{mes}</text:p></table:table-cell>'
    '<table:table-cell office:value-type="float" office:value="{d}">'
    '<text:p>{d}</text:p></table:table-cell>'
    '</table:table-row>'
)


# --------------------------------------------------------------------------
# leitura do ODS (ElementTree, so leitura)
# --------------------------------------------------------------------------
def trecho_planilha1(content):
    """Devolve (inicio, fim) do trecho do content.xml que pertence a Planilha1."""
    ini = content.index('table:name="Planilha1"')
    fim = content.find("<table:table ", ini)
    return ini, (fim if fim > 0 else len(content))


def folha_planilha1(raiz):
    for t in raiz.iter(q("table:table")):
        if t.get(q("table:name")) == "Planilha1":
            return t
    raise ValueError("aba Planilha1 nao encontrada")


def celulas(linha, quantas):
    """Primeiras `quantas` celulas da linha, expandindo number-columns-repeated."""
    saida = []
    for c in linha:
        if c.tag not in (q("table:table-cell"), q("table:covered-table-cell")):
            continue
        saida.extend([c] * int(c.get(q("table:number-columns-repeated"), "1")))
        if len(saida) >= quantas:
            break
    return saida[:quantas]


def texto_celula(celula):
    """Texto dos <text:p> da propria celula (ignora caixas de texto ancoradas)."""
    return "".join("".join(p.itertext()) for p in celula.findall("text:p", NS)).strip()


def serie_do_content(content_xml):
    """Le [[AAAAMM, indice], ...] da Planilha1 (colunas A e B) de um content.xml."""
    folha = folha_planilha1(ET.fromstring(content_xml))
    serie = []
    for linha in folha.iter(q("table:table-row")):
        cel = celulas(linha, 2)
        if len(cel) < 2:
            continue
        m = re.fullmatch(r"(\w+) (\d{4})", texto_celula(cel[0]))
        if not m or m.group(1) not in MESES_PT:
            continue
        valor = cel[1].get(q("office:value"))
        if valor is None:
            raise ValueError(f"{m.group(0)} sem indice na coluna B")
        serie.append([int(m.group(2)) * 100 + MESES_PT.index(m.group(1)) + 1,
                      round(float(valor), 2)])
    serie.sort()
    return serie


def ler_serie(caminho_ods):
    """Le a planilha e devolve [[AAAAMM, indice], ...]."""
    with zipfile.ZipFile(caminho_ods) as z:
        return serie_do_content(z.read("content.xml"))


def ler_meta(caminho_ods):
    """Devolve {"coleta": ..., "fonte": ...} das propriedades do documento
    (so as que existirem)."""
    with zipfile.ZipFile(caminho_ods) as z:
        if "meta.xml" not in z.namelist():
            return {}
        raiz = ET.fromstring(z.read("meta.xml"))
    props = {u.get(q("meta:name")): (u.text or "").strip()
             for u in raiz.iter(q("meta:user-defined"))}
    meta = {}
    if props.get(PROP_COLETA):
        meta["coleta"] = props[PROP_COLETA]
    if props.get(PROP_FONTE):
        meta["fonte"] = props[PROP_FONTE]
    return meta


def gravar_meta(meta_xml, coleta, fonte):
    """Grava/atualiza IPCA_coleta e IPCA_fonte no meta.xml (texto)."""
    for nome, valor in ((PROP_COLETA, coleta), (PROP_FONTE, fonte)):
        elemento = (f'<meta:user-defined meta:name="{nome}" '
                    f'meta:value-type="string">{escape(valor)}</meta:user-defined>')
        padrao = re.compile(r'<meta:user-defined meta:name="%s"[^>]*?'
                            r'(?:/>|>.*?</meta:user-defined>)' % re.escape(nome), re.S)
        if padrao.search(meta_xml):
            meta_xml = padrao.sub(lambda _: elemento, meta_xml, count=1)
        else:
            meta_xml = meta_xml.replace("</office:meta>", elemento + "</office:meta>", 1)
    return meta_xml


def contagens(conteudo, grafico_nome):
    """Numeros que o script precisa manter coerentes; usados antes e depois da
    gravacao para conferir o arquivo novo."""
    raiz = ET.fromstring(conteudo["content.xml"])      # tambem testa se e XML valido
    folha = folha_planilha1(raiz)
    total = sum(int(r.get(q("table:number-rows-repeated"), "1"))
                for r in folha.iter(q("table:table-row")))
    r = {"serie": serie_do_content(conteudo["content.xml"]), "linhas_folha": total,
         "cache": None, "pontos": None}
    if "meta.xml" in conteudo:
        ET.fromstring(conteudo["meta.xml"])
    if grafico_nome:
        g = ET.fromstring(conteudo[grafico_nome])
        linhas = g.find(".//" + q("table:table-rows"))
        r["cache"] = len(linhas.findall(q("table:table-row"))) if linhas is not None else None
        r["pontos"] = [sum(int(p.get(q("chart:repeated"), "1"))
                           for p in s.findall(q("chart:data-point")))
                       for s in g.iter(q("chart:series"))]
    return r


def perto(a, b):
    return abs(a - b) <= 1e-9 * max(1.0, abs(b))


def conferir_conteudo(content_xml, n, ultima_linha_nova):
    """Confere, no content.xml novo, as colunas C e D das n ultimas linhas da
    serie (formula e valor) e a caixa de texto com o mes e o valor final."""
    problemas = []
    folha = folha_planilha1(ET.fromstring(content_xml))
    dados = []
    for linha in folha.iter(q("table:table-row")):
        cel = celulas(linha, 4)
        if len(cel) == 4 and re.fullmatch(r"(\w+) (\d{4})", texto_celula(cel[0])):
            dados.append(cel)
    for k, cel in enumerate(dados[-n:] if n else []):
        lin = ultima_linha_nova - n + 1 + k
        nome = texto_celula(cel[0])
        try:
            idx = float(cel[1].get(q("office:value")))
            c = float(cel[2].get(q("office:value")))
            d = float(cel[3].get(q("office:value")))
        except (TypeError, ValueError):
            problemas.append(f"{nome} (linha {lin}): valor ausente nas colunas B a D")
            continue
        if cel[2].get(q("table:formula")) != f"of:=[.B{lin}]/{num(BASE)}":
            problemas.append(f"{nome} (linha {lin}): formula da coluna C nao aponta para B{lin}")
        if cel[3].get(q("table:formula")) != f"of:=(1/[.C{lin}])":
            problemas.append(f"{nome} (linha {lin}): formula da coluna D nao aponta para C{lin}")
        if not perto(c, idx / BASE):
            problemas.append(f"{nome}: coluna C = {c}, esperado {idx / BASE}")
        if not perto(d, BASE / idx):
            problemas.append(f"{nome}: coluna D = {d}, esperado {BASE / idx}")

    # caixa de texto: mes e valor de R$ 100 de jul/1994 no ultimo mes
    if n and dados:
        idx = float(dados[-1][1].get(q("office:value")))
        m = re.fullmatch(r"(\w+) (\d{4})", texto_celula(dados[-1][0]))
        esperado = (m.group(1)[0].upper() + m.group(1)[1:] + " " + m.group(2),
                    br(f"{BASE / idx * 100:.2f}"))
        caixa = next(folha.iter(q("draw:text-box")), None)
        spans = {}
        if caixa is not None:
            for sp in caixa.iter(q("text:span")):
                spans.setdefault(sp.get(q("text:style-name")), "".join(sp.itertext()))
        achado = (spans.get("T1"), spans.get("T3"))
        if achado != esperado:
            problemas.append(f"caixa de texto mostra {achado}; esperado {esperado}")
    return problemas


def verificar_ods(caminho, antes, n, ultimo, ultima_linha_nova, grafico_nome):
    """Confere o .ods recem-gravado. Devolve a lista de problemas (vazia = ok)."""
    problemas = []
    with zipfile.ZipFile(caminho) as z:
        infos = z.infolist()
        conteudo = {i.filename: z.read(i.filename) for i in infos}
        if z.testzip() is not None:
            problemas.append("pacote ZIP corrompido")
    if "mimetype" in conteudo and (infos[0].filename != "mimetype"
                                   or infos[0].compress_type != zipfile.ZIP_STORED):
        problemas.append("'mimetype' nao e a primeira entrada sem compressao")
    try:
        depois = contagens(conteudo, grafico_nome)
    except (ET.ParseError, ValueError) as e:
        return problemas + [f"XML invalido ou ilegivel: {e}"]

    s0, s1 = antes["serie"], depois["serie"]
    if len(s1) != len(s0) + n:
        problemas.append(f"a serie tem {len(s1)} meses; esperado {len(s0) + n}")
    if s1[:len(s0)] != s0:
        problemas.append("meses antigos mudaram")
    if not s1 or s1[-1] != list(ultimo):
        problemas.append(f"ultimo mes lido {s1[-1] if s1 else None}; esperado {list(ultimo)}")
    if lacunas([(c // 100, c % 100) for c, _ in s1]):
        problemas.append("a serie da planilha tem lacuna")
    if depois["linhas_folha"] != antes["linhas_folha"]:
        problemas.append(f"Planilha1 tem {depois['linhas_folha']} linhas; "
                         f"esperado {antes['linhas_folha']}")

    problemas += conferir_conteudo(conteudo["content.xml"], n, ultima_linha_nova)

    textos = [conteudo["content.xml"].decode("utf-8")]
    if grafico_nome:
        textos.append(conteudo[grafico_nome].decode("utf-8"))
        if antes["cache"] is not None and depois["cache"] != antes["cache"] + n:
            problemas.append(f"cache do grafico com {depois['cache']} linhas; "
                             f"esperado {antes['cache'] + n}")
        esperado = [p + n for p in antes["pontos"]]
        if depois["pontos"] != esperado:
            problemas.append(f"pontos do grafico {depois['pontos']}; esperado {esperado}")
    for col in ("A", "D"):
        alvo = f"Planilha1.{col}{PRIMEIRA_LINHA}:Planilha1.{col}{ultima_linha_nova}"
        if not all(alvo in t for t in textos):
            problemas.append(f"intervalo {alvo} ausente")
    return problemas


# --------------------------------------------------------------------------
# geracao da pagina HTML
# --------------------------------------------------------------------------
def montar_html(serie, meta, pasta_script):
    """Devolve o HTML final (modelo + serie + metadados) ou None se nao der."""
    modelo = os.path.join(pasta_script, MODELO)
    if not os.path.exists(modelo):
        aviso(f"{MODELO} nao encontrado em {pasta_script}; "
              "a pagina HTML nao foi gerada.")
        return None

    with open(modelo, encoding="utf-8") as fh:
        html = fh.read()

    dados = "[" + ",".join(f"[{c},{v:g}]" for c, v in serie) + "]"
    if MARCA_SERIE not in html:
        aviso(f"marcador {MARCA_SERIE} ausente em {MODELO}; "
              "a pagina HTML nao foi gerada.")
        return None

    # substitui o marcador e o array que o segue (seja [] ou uma serie antiga).
    # O array vai ate o primeiro "];" -- um "]" sozinho fecharia so o primeiro par.
    html, trocas = re.subn(re.escape(MARCA_SERIE) + r"\s*\[.*?\];",
                           lambda _: MARCA_SERIE + dados + ";", html, count=1, flags=re.S)
    if not trocas:
        aviso(f"o array apos {MARCA_SERIE} em {MODELO} nao termina com '];'; "
              "a pagina HTML nao foi gerada.")
        return None

    # metadados da coleta: objeto plano, vai ate o primeiro "};"
    if MARCA_META in html:
        js = json.dumps(meta, ensure_ascii=False, sort_keys=True).replace("</", "<\\/")
        html = re.sub(re.escape(MARCA_META) + r"\s*\{.*?\};",
                      lambda _: MARCA_META + js + ";", html, count=1, flags=re.S)
    return html


def gerar_html(serie, meta, pasta_script, destino=None):
    """Injeta a serie no modelo.html e grava a pagina final."""
    html = montar_html(serie, meta, pasta_script)
    if html is None:
        return None
    destino = destino or os.path.join(pasta_script, PAGINA)
    with open(destino, "w", encoding="utf-8") as fh:
        fh.write(html)
    print(f"Pagina gerada: {destino} ({len(serie)} meses, "
          f"ate {rotulo(serie[-1][0] // 100, serie[-1][0] % 100)})")
    return destino


def verificar(caminho_ods, pasta_script, destino=None):
    """Confere se o index.html corresponde a planilha e ao modelo atuais."""
    destino = destino or os.path.join(pasta_script, PAGINA)
    serie = ler_serie(caminho_ods)
    erros = []
    if not serie:
        erros.append("planilha sem serie")
    else:
        for esperado, achado in lacunas([(c // 100, c % 100) for c, _ in serie]):
            erros.append(f"lacuna na planilha: falta {rotulo(*esperado)} (veio {rotulo(*achado)})")
    esperado = montar_html(serie, ler_meta(caminho_ods), pasta_script) if serie else None
    if esperado is None:
        erros.append("nao foi possivel montar a pagina a partir do modelo")
    elif not os.path.exists(destino):
        erros.append(f"{destino} nao existe")
    else:
        with open(destino, encoding="utf-8") as fh:
            atual = fh.read()
        if atual != esperado:
            m = re.search(re.escape(MARCA_SERIE) + r"(\[.*?\]);", atual, re.S)
            serie_html = json.loads(m.group(1)) if m else None
            if serie_html != serie:
                n = len(serie_html) if serie_html else 0
                erros.append(f"a serie do {os.path.basename(destino)} ({n} meses) "
                             f"difere da planilha ({len(serie)} meses)")
            else:
                erros.append(f"{os.path.basename(destino)} difere do que o "
                             f"{MODELO} geraria (modelo ou metadados mudaram)")
    if erros:
        for e in erros:
            print(f"ERRO: {e}"); anotacao("error", e)
        falhar("Verificacao falhou: rode o script sem --sem-html/--so-html "
                 "para sincronizar.")
    print(f"Verificacao ok: planilha e {os.path.basename(destino)} com {len(serie)} "
          f"meses, ate {rotulo(serie[-1][0] // 100, serie[-1][0] % 100)}.")


# --------------------------------------------------------------------------
# atualizacao da planilha
# --------------------------------------------------------------------------
def atualizar(caminho_entrada, caminho_saida, fonte=None, simular=False,
              conferir_com=None):
    with zipfile.ZipFile(caminho_entrada) as z:
        arquivos = z.namelist()
        conteudo = {nome: z.read(nome) for nome in arquivos}

    content = conteudo["content.xml"].decode("utf-8")
    grafico_nome = next((n for n in arquivos if n.endswith("/content.xml")
                         and n.startswith("Object")), None)
    grafico = conteudo[grafico_nome].decode("utf-8") if grafico_nome else None

    # ---- descobre a ultima linha preenchida --------------------------------
    linhas = [int(m) for m in re.findall(r'of:=\[\.B(\d+)\]/', content)]
    if not linhas:
        falhar("Nao encontrei a serie na planilha (formula =B/base).")
    ultima_linha = max(linhas)

    # ultimo rotulo e ultimo indice
    m = re.search(
        r'<text:p>([a-zçã]+ \d{4})</text:p></table:table-cell>'
        r'<table:table-cell[^>]*office:value="([\d.]+)"[^>]*>'
        r'(?:(?!</table:table-row>).)*?of:=\[\.B%d\]' % ultima_linha,
        content, re.S)
    if not m:
        falhar("Nao consegui ler o ultimo mes da serie.")
    ultimo_rotulo, ultimo_indice = m.group(1), float(m.group(2))
    nome_mes, ano_txt = ultimo_rotulo.rsplit(" ", 1)
    ultimo_ano, ultimo_mes = int(ano_txt), MESES_PT.index(nome_mes) + 1

    print(f"Planilha atual: ate {ultimo_rotulo} "
          f"(linha {ultima_linha}, indice {ultimo_indice})")

    # ---- confere a planilha com o leitor XML -------------------------------
    try:
        antes = contagens(conteudo, grafico_nome)
    except (ET.ParseError, ValueError) as e:
        falhar(f"A planilha de entrada nao pode ser lida: {e}")
    serie_planilha = antes["serie"]
    if serie_planilha[-1] != [ultimo_ano * 100 + ultimo_mes, round(ultimo_indice, 2)]:
        falhar(f"Leituras divergentes do ultimo mes: {serie_planilha[-1]} (XML) x "
                 f"{ultimo_rotulo} {ultimo_indice} (formula da linha {ultima_linha}).")
    if len(serie_planilha) != ultima_linha - PRIMEIRA_LINHA + 1:
        falhar(f"A Planilha1 tem {len(serie_planilha)} meses, mas a ultima formula "
                 f"esta na linha {ultima_linha}.")
    furos = lacunas([(c // 100, c % 100) for c, _ in serie_planilha])
    if furos:
        falhar(f"A planilha tem lacuna: falta {rotulo(*furos[0][0])}.")

    # ---- baixa, confere e filtra o que falta -------------------------------
    serie, nome_fonte, conferencia, nome_conf = baixar_serie(
        fonte, mes_menos(ultimo_ano, ultimo_mes, JANELA_MESES), conferir_com)
    print(f"Recebidos {len(serie)} meses de {nome_fonte} "
          f"({rotulo(*serie[0][:2])} .. {rotulo(*serie[-1][:2])}).")

    for cod, na_planilha, baixada in checar_revisoes(serie, serie_planilha):
        aviso(f"{rotulo(cod // 100, cod % 100)} na planilha = {na_planilha:+.2f}%, "
              f"na fonte = {baixada:+.2f}% (possivel revisao; a planilha nao foi alterada).")

    faltando = [(a, me, v) for (a, me, v) in serie
                if (a, me) > (ultimo_ano, ultimo_mes)]
    if not faltando:
        print("Nada a fazer: a planilha ja esta no mes mais recente publicado.")
        return False
    if faltando[0][:2] != mes_seguinte(ultimo_ano, ultimo_mes):
        falhar(f"Lacuna: a planilha vai ate {ultimo_rotulo}, mas a fonte so "
                 f"traz a partir de {rotulo(*faltando[0][:2])}. Nada foi gravado.")

    print(f"Meses novos: {len(faltando)} "
          f"({rotulo(*faltando[0][:2])} .. {rotulo(*faltando[-1][:2])})")

    # conferencia cruzada: online, so quando a fonte principal foi o BCB;
    # com --fonte, so se --conferir-com foi dado
    if conferir_com or (not fonte and nome_fonte == FONTE_BCB):
        conferir_fontes(faltando, conferencia, nome_conf)

    if simular:
        idx = ultimo_indice
        for a, me, v in faltando:
            idx = round(idx * (1 + v / 100), 2)
            print(f"  {rotulo(a, me):>18}  {v:+.2f}%   indice {idx:.2f}   "
                  f"poder de compra {BASE / idx * 100:.2f}%")
        return False

    # ---- monta as novas linhas --------------------------------------------
    novas_planilha, novas_grafico = [], []
    idx = ultimo_indice
    linha = ultima_linha
    for a, me, v in faltando:
        idx = round(idx * (1 + v / 100), 2)
        linha += 1
        c = idx / BASE
        d = 1 / c
        novas_planilha.append(LINHA_PLANILHA.format(
            mes=rotulo(a, me),
            idx=num(idx), idx_br=br(f"{idx:.2f}"),
            lin=linha, base=num(BASE),
            c=num(c), c_br=br(num(c)),
            d=num(d), d_br=br(f"{d * 100:.2f}") + "%",
        ))
        novas_grafico.append(LINHA_GRAFICO.format(mes=rotulo(a, me), d=num(d)))
    n = len(faltando)
    nova_ultima_linha = linha
    ultimo_ano_n, ultimo_mes_n = faltando[-1][0], faltando[-1][1]

    # ---- insere as linhas na planilha -------------------------------------
    marca = 'of:=(1/[.C%d])' % ultima_linha
    pos = content.index(marca)
    fim = content.index("</table:table-row>", pos) + len("</table:table-row>")
    content = content[:fim] + "".join(novas_planilha) + content[fim:]

    # A partir daqui as substituicoes valem so dentro da Planilha1, para nao
    # depender da ordem das abas nem de estilos repetidos em outras folhas.
    p_ini, p_fim = trecho_planilha1(content)
    folha = content[p_ini:p_fim]

    # mantem o total de 1.048.576 linhas da folha
    def desconta(mo):
        return 'table:number-rows-repeated="%d"' % max(1, int(mo.group(1)) - n)
    folha = re.sub(r'table:number-rows-repeated="(\d{5,})"', desconta,
                   folha, count=1)

    # ---- atualiza os intervalos usados pelo grafico ------------------------
    def trocar_intervalos(texto):
        for col in ("A", "D"):
            texto = texto.replace(
                f"Planilha1.{col}{PRIMEIRA_LINHA}:Planilha1.{col}{ultima_linha}",
                f"Planilha1.{col}{PRIMEIRA_LINHA}:Planilha1.{col}{nova_ultima_linha}")
        return texto

    # ---- caixa de texto com o valor final ---------------------------------
    # A primeira caixa de texto da Planilha1 e a do mes atual; a segunda (fixa)
    # e a de 1994. As duas trocas ficam presas ao mesmo <draw:text-box>.
    valor_final = BASE / idx * 100          # R$ 100 de jul/1994 valem hoje...
    caixa = re.search(r"<draw:text-box>.*?</draw:text-box>", folha, re.S)
    if caixa:
        bloco = re.sub(
            r'(<text:span text:style-name="T1">)[^<]*(</text:span>)',
            lambda mo: mo.group(1) + rotulo_titulo(ultimo_ano_n, ultimo_mes_n) + mo.group(2),
            caixa.group(0), count=1)
        bloco = re.sub(
            r'(<text:span text:style-name="T3">)[^<]*(</text:span>)',
            lambda mo: mo.group(1) + br(f"{valor_final:.2f}") + mo.group(2),
            bloco, count=1)
        folha = folha[:caixa.start()] + bloco + folha[caixa.end():]
    else:
        aviso("caixa de texto do valor final nao encontrada na Planilha1.")

    content = trocar_intervalos(content[:p_ini] + folha + content[p_fim:])

    conteudo["content.xml"] = content.encode("utf-8")

    # ---- grafico: intervalos, cache local e contagem de pontos ------------
    if grafico:
        grafico = trocar_intervalos(grafico)
        grafico = grafico.replace(
            "</table:table-rows>", "".join(novas_grafico) + "</table:table-rows>")
        pontos = re.findall(r'<chart:data-point chart:repeated="(\d+)"/>', grafico)
        if pontos:
            ultimo = pontos[-1]
            alvo = '<chart:data-point chart:repeated="%s"/>' % ultimo
            i = grafico.rindex(alvo)
            grafico = (grafico[:i]
                       + '<chart:data-point chart:repeated="%d"/>' % (int(ultimo) + n)
                       + grafico[i + len(alvo):])
        conteudo[grafico_nome] = grafico.encode("utf-8")

    # ---- proveniencia: data e fonte da coleta nas propriedades -------------
    if "meta.xml" in conteudo:
        conteudo["meta.xml"] = gravar_meta(
            conteudo["meta.xml"].decode("utf-8"),
            time.strftime("%Y-%m-%d"), nome_fonte).encode("utf-8")

    # ---- grava num temporario, confere e so entao substitui ---------------
    fd, tmp = tempfile.mkstemp(suffix=".ods")
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            # "mimetype" precisa ser o primeiro item e sem compressao
            if "mimetype" in conteudo:
                z.writestr(zipfile.ZipInfo("mimetype"), conteudo["mimetype"],
                           compress_type=zipfile.ZIP_STORED)
            for nome in arquivos:
                if nome == "mimetype":
                    continue
                z.writestr(nome, conteudo[nome])
        problemas = verificar_ods(tmp, antes, n,
                                  (ultimo_ano_n * 100 + ultimo_mes_n, round(idx, 2)),
                                  nova_ultima_linha, grafico_nome)
        if problemas:
            for p in problemas:
                print(f"ERRO: {p}"); anotacao("error", p)
            falhar("O arquivo novo nao passou na conferencia; "
                     f"{caminho_saida} nao foi alterado.")
        shutil.move(tmp, caminho_saida)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)

    print(f"Planilha atualizada ate {rotulo(ultimo_ano_n, ultimo_mes_n)}: "
          f"indice {idx:.2f} | R$ 100 de julho/1994 valem "
          f"R$ {valor_final:.2f}".replace(".", ","))
    print(f"Arquivo gravado e conferido: {caminho_saida}")
    return True


def main():
    p = argparse.ArgumentParser(description="Atualiza a planilha do IPCA.")
    p.add_argument("-e", "--entrada",
                   default="Desvalorização_Real_Gráfico.ods")
    p.add_argument("-s", "--saida", default=None,
                   help="padrao: grava por cima do arquivo de entrada")
    p.add_argument("--fonte", default=None,
                   help="URL no formato do SGS ou JSON local (desliga a reserva do IBGE)")
    p.add_argument("--simular", action="store_true",
                   help="mostra o que seria acrescentado, sem gravar")
    p.add_argument("--so-html", action="store_true",
                   help="nao mexe na planilha; so regera a pagina HTML")
    p.add_argument("--sem-html", action="store_true",
                   help="nao gera a pagina HTML")
    p.add_argument("--conferir-com", default=None, metavar="JSON",
                   help="JSON local (formato do SGS) usado na conferencia cruzada; para testes")
    p.add_argument("--verificar", action="store_true",
                   help="so confere se o index.html corresponde a planilha (sai com 1 se nao)")
    p.add_argument("--html", default=None,
                   help=f"caminho da pagina gerada (padrao: {PAGINA})")
    a = p.parse_args()

    if not os.path.exists(a.entrada):
        falhar(f"Arquivo nao encontrado: {a.entrada}")

    pasta = os.path.dirname(os.path.abspath(__file__))
    saida = a.saida or a.entrada

    if a.verificar:
        verificar(a.entrada, pasta, a.html)
        return

    if not a.so_html:
        atualizar(a.entrada, saida, a.fonte, a.simular, a.conferir_com)
        if a.simular:
            return
    else:
        saida = a.entrada

    if not a.sem_html:
        gerar_html(ler_serie(saida), ler_meta(saida), pasta, a.html)


if __name__ == "__main__":
    main()
