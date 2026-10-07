#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
atualizar_ipca.py — atualiza a planilha "Desvalorização_Real_Gráfico.ods" com os
dados mais recentes do IPCA, preservando a formatacao, as formulas e o grafico.

Fonte dos dados
---------------
Banco Central do Brasil - SGS (Sistema Gerenciador de Series Temporais),
serie 433 = "IPCA - variacao mensal (%)" (dado produzido pelo IBGE).
    https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados?formato=json

O numero-indice da coluna B (base dez/1993 = 100, mesma da tabela 1737 do
SIDRA/IBGE) e reconstruido encadeando as variacoes mensais a partir do ultimo
valor ja existente na planilha:

    indice_novo = arredonda(indice_anterior * (1 + variacao/100), 2)

Alem da planilha, o script regrava uma pagina HTML interativa
("Desvalorizacao_Real.html") a partir do arquivo "modelo.html", que fica ao
lado do script. O modelo e um HTML comum: edite o visual nele a vontade; o
unico ponto que o script toca e o marcador /*__SERIE__*/, onde a serie e
injetada. Sem o modelo.html por perto, o script apenas avisa e segue.

Uso
---
    python3 atualizar_ipca.py                      # atualiza planilha + HTML
    python3 atualizar_ipca.py -e entrada.ods -s saida.ods
    python3 atualizar_ipca.py --fonte dados.json   # usa um JSON local (offline)
    python3 atualizar_ipca.py --simular            # so mostra o que falta
    python3 atualizar_ipca.py --so-html            # so regera o HTML
    python3 atualizar_ipca.py --sem-html           # so mexe na planilha

O script e idempotente: rodar duas vezes seguidas nao duplica nada.
"""

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import urllib.request
import zipfile

URL_BCB = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados?formato=json"

MESES = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho",
         "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
# nomes exatamente como aparecem na planilha (com acento)
MESES_PT = ["janeiro", "fevereiro", "mar\u00e7o", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

BASE = 915.93          # indice de julho/1994 (celula B3) = base do grafico
MODELO = "modelo.html"          # template da pagina (ao lado do script)
PAGINA = "Desvalorizacao_Real.html"  # pagina gerada
PRIMEIRA_LINHA = 3     # linha da planilha onde comeca a serie


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


def baixar_serie(fonte):
    """Devolve [(ano, mes, variacao_percentual), ...] ordenado."""
    if fonte and not fonte.startswith("http"):
        with open(fonte, encoding="utf-8") as fh:
            dados = json.load(fh)
    else:
        req = urllib.request.Request(
            fonte or URL_BCB,
            headers={"User-Agent": "atualizar_ipca/1.0"},
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            dados = json.loads(resp.read().decode("utf-8"))

    serie = []
    for item in dados:
        dia, mes, ano = item["data"].split("/")
        serie.append((int(ano), int(mes), float(item["valor"])))
    serie.sort()
    return serie


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
# leitura da serie e geracao da pagina HTML
# --------------------------------------------------------------------------
def ler_serie(caminho_ods):
    """Le a planilha e devolve [[AAAAMM, indice], ...] direto do XML do ODS."""
    with zipfile.ZipFile(caminho_ods) as z:
        content = z.read("content.xml").decode("utf-8")

    # corta na Planilha1 para nao pegar outras abas
    ini = content.index('table:name="Planilha1"')
    fim = content.find("<table:table ", ini)
    trecho = content[ini:fim if fim > 0 else len(content)]

    padrao = re.compile(
        r"<text:p>([a-z\u00e7\u00e3]+) (\d{4})</text:p></table:table-cell>"
        r'<table:table-cell[^>]*office:value="([\d.]+)"'
    )
    serie = []
    for nome, ano, indice in padrao.findall(trecho):
        if nome not in MESES_PT:
            continue
        serie.append([int(ano) * 100 + MESES_PT.index(nome) + 1,
                      round(float(indice), 2)])
    serie.sort()
    return serie


def gerar_html(serie, pasta_script, destino=None):
    """Injeta a serie no modelo.html e grava a pagina final."""
    modelo = os.path.join(pasta_script, MODELO)
    if not os.path.exists(modelo):
        print(f"Aviso: {MODELO} nao encontrado em {pasta_script}; "
              "a pagina HTML nao foi gerada.")
        return None

    with open(modelo, encoding="utf-8") as fh:
        html = fh.read()

    dados = "[" + ",".join(f"[{c},{v:g}]" for c, v in serie) + "]"
    marcador = "/*__SERIE__*/"
    if marcador not in html:
        print(f"Aviso: marcador {marcador} ausente em {MODELO}; "
              "a pagina HTML nao foi gerada.")
        return None

    # substitui o marcador e o array que o segue (seja [] ou uma serie antiga)
    html = re.sub(re.escape(marcador) + r"\s*\[.*?\]",
                  marcador + dados, html, count=1, flags=re.S)

    destino = destino or os.path.join(pasta_script, PAGINA)
    with open(destino, "w", encoding="utf-8") as fh:
        fh.write(html)
    print(f"Pagina gerada: {destino} ({len(serie)} meses, "
          f"ate {rotulo(serie[-1][0] // 100, serie[-1][0] % 100)})")
    return destino


def atualizar(caminho_entrada, caminho_saida, fonte=None, simular=False):
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
        sys.exit("Nao encontrei a serie na planilha (formula =B/base).")
    ultima_linha = max(linhas)

    # ultimo rotulo e ultimo indice
    m = re.search(
        r'<text:p>([a-z\u00e7\u00e3]+ \d{4})</text:p></table:table-cell>'
        r'<table:table-cell[^>]*office:value="([\d.]+)"[^>]*>'
        r'(?:(?!</table:table-row>).)*?of:=\[\.B%d\]' % ultima_linha,
        content, re.S)
    if not m:
        sys.exit("Nao consegui ler o ultimo mes da serie.")
    ultimo_rotulo, ultimo_indice = m.group(1), float(m.group(2))
    nome_mes, ano_txt = ultimo_rotulo.rsplit(" ", 1)
    ultimo_ano, ultimo_mes = int(ano_txt), MESES_PT.index(nome_mes) + 1

    print(f"Planilha atual: ate {ultimo_rotulo} "
          f"(linha {ultima_linha}, indice {ultimo_indice})")

    # ---- baixa e filtra o que falta ---------------------------------------
    serie = baixar_serie(fonte)
    faltando = [(a, me, v) for (a, me, v) in serie
                if (a, me) > (ultimo_ano, ultimo_mes)]
    if not faltando:
        print("Nada a fazer: a planilha ja esta no mes mais recente publicado.")
        return False

    print(f"Meses novos: {len(faltando)} "
          f"({rotulo(*faltando[0][:2])} .. {rotulo(*faltando[-1][:2])})")

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

    # mantem o total de 1.048.576 linhas da folha
    def desconta(mo):
        return 'table:number-rows-repeated="%d"' % max(1, int(mo.group(1)) - n)
    content = re.sub(r'table:number-rows-repeated="(\d{5,})"', desconta,
                     content, count=1)

    # ---- atualiza os intervalos usados pelo grafico ------------------------
    def trocar_intervalos(texto):
        for col in ("A", "D"):
            texto = texto.replace(
                f"Planilha1.{col}{PRIMEIRA_LINHA}:Planilha1.{col}{ultima_linha}",
                f"Planilha1.{col}{PRIMEIRA_LINHA}:Planilha1.{col}{nova_ultima_linha}")
        return texto

    content = trocar_intervalos(content)

    # ---- caixa de texto com o valor final ---------------------------------
    valor_final = BASE / idx * 100          # R$ 100 de jul/1994 valem hoje...
    content = re.sub(
        r'(<draw:text-box>.*?<text:span text:style-name="T1">)[^<]*(</text:span>)',
        lambda mo: mo.group(1) + rotulo_titulo(ultimo_ano_n, ultimo_mes_n) + mo.group(2),
        content, count=1, flags=re.S)
    content = re.sub(
        r'(<text:span text:style-name="T3">)[^<]*(</text:span>)',
        lambda mo: mo.group(1) + br(f"{valor_final:.2f}") + mo.group(2),
        content, count=1, flags=re.S)

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

    # ---- regrava o pacote ODS ---------------------------------------------
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".ods")
    tmp.close()
    with zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as z:
        # "mimetype" precisa ser o primeiro item e sem compressao
        if "mimetype" in conteudo:
            z.writestr(zipfile.ZipInfo("mimetype"), conteudo["mimetype"],
                       compress_type=zipfile.ZIP_STORED)
        for nome in arquivos:
            if nome == "mimetype":
                continue
            z.writestr(nome, conteudo[nome])
    shutil.move(tmp.name, caminho_saida)

    print(f"Planilha atualizada ate {rotulo(ultimo_ano_n, ultimo_mes_n)}: "
          f"indice {idx:.2f} | R$ 100 de julho/1994 valem "
          f"R$ {valor_final:.2f}".replace(".", ","))
    print(f"Arquivo gravado: {caminho_saida}")
    return True


def main():
    p = argparse.ArgumentParser(description="Atualiza a planilha do IPCA.")
    p.add_argument("-e", "--entrada",
                   default="Desvalorização_Real_Gráfico.ods")
    p.add_argument("-s", "--saida", default=None,
                   help="padrao: grava por cima do arquivo de entrada")
    p.add_argument("--fonte", default=None,
                   help="URL alternativa ou caminho de um JSON local")
    p.add_argument("--simular", action="store_true",
                   help="mostra o que seria acrescentado, sem gravar")
    p.add_argument("--so-html", action="store_true",
                   help="nao mexe na planilha; so regera a pagina HTML")
    p.add_argument("--sem-html", action="store_true",
                   help="nao gera a pagina HTML")
    p.add_argument("--html", default=None,
                   help=f"caminho da pagina gerada (padrao: {PAGINA})")
    a = p.parse_args()

    if not os.path.exists(a.entrada):
        sys.exit(f"Arquivo nao encontrado: {a.entrada}")

    pasta = os.path.dirname(os.path.abspath(__file__))
    saida = a.saida or a.entrada

    if not a.so_html:
        atualizar(a.entrada, saida, a.fonte, a.simular)
        if a.simular:
            return
    else:
        saida = a.entrada

    if not a.sem_html:
        gerar_html(ler_serie(saida), pasta, a.html)


if __name__ == "__main__":
    main()
