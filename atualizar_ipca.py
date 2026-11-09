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

O numero-indice da coluna B (base dez/1993 = 100, mesma da tabela 1737 do
SIDRA/IBGE) e reconstruido encadeando as variacoes mensais a partir do ultimo
valor ja existente na planilha:

    indice_novo = arredonda(indice_anterior * (1 + variacao/100), 2)

Alem da planilha, o script regrava uma pagina HTML interativa
("index.html") a partir do arquivo "modelo.html", que fica ao
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
import time
import urllib.error
import urllib.request
import zipfile

URL_BCB = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados"
URL_IBGE = ("https://servicodados.ibge.gov.br/api/v3/agregados/1737/"
            "periodos/{periodos}/variaveis/63?localidades=N1%5Ball%5D")
CABECALHOS = {
    "User-Agent": "Mozilla/5.0 (compatible; atualizar_ipca/1.1)",
    "Accept": "application/json",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
}
JANELA_MESES = 24      # quantos meses antes do ultimo da planilha pedir de novo

MESES = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho",
         "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
# nomes exatamente como aparecem na planilha (com acento)
MESES_PT = ["janeiro", "fevereiro", "mar\u00e7o", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

BASE = 915.93          # indice de julho/1994 (celula B3) = base do grafico
MODELO = "modelo.html"          # template da pagina (ao lado do script)
PAGINA = "index.html"  # pagina gerada (servida pelo GitHub Pages)
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


def mes_menos(ano, mes, k):
    """(ano, mes) recuado k meses."""
    t = ano * 12 + (mes - 1) - k
    return t // 12, t % 12 + 1


def ler_bcb(dados):
    """Formato SGS: [{"data": "01/09/2026", "valor": "0.48"}, ...]"""
    serie = []
    for item in dados:
        _, mes, ano = item["data"].split("/")
        serie.append((int(ano), int(mes), float(item["valor"])))
    return serie


def ler_ibge(dados):
    """Formato IBGE v3: [{"resultados": [{"series": [{"serie": {"202609": "0.48"}}]}]}]"""
    serie = []
    for periodo, valor in dados[0]["resultados"][0]["series"][0]["serie"].items():
        try:
            v = float(valor)
        except (TypeError, ValueError):
            continue        # "..." ou "-": mes ainda sem dado
        serie.append((int(periodo[:4]), int(periodo[4:]), v))
    return serie


def baixar_serie(fonte, desde):
    """Devolve [(ano, mes, variacao_percentual), ...] ordenado.

    fonte: None (BCB com reserva no IBGE), URL no formato do SGS ou caminho
    de um JSON local no formato do SGS. desde: (ano, mes) a partir do qual
    os dados sao pedidos.
    """
    if fonte and not fonte.startswith(("http://", "https://")):
        with open(fonte, encoding="utf-8") as fh:
            serie = ler_bcb(json.load(fh))
        return sorted(serie)

    a, m = desde
    url_bcb = fonte or (f"{URL_BCB}?formato=json&dataInicial=01/{m:02d}/{a}"
                        f"&dataFinal={time.strftime('%d/%m/%Y')}")
    tentativas = [("Banco Central (SGS 433)", url_bcb, ler_bcb)]
    if not fonte:
        tentativas.append(("IBGE (tabela 1737)",
                           URL_IBGE.format(periodos=f"-{JANELA_MESES + 12}"), ler_ibge))

    for nome, url, ler in tentativas:
        print(f"Baixando de {nome}...")
        try:
            serie = ler(baixar_json(url))
        except FalhaDownload:
            continue
        except (KeyError, IndexError, TypeError, ValueError) as e:
            print(f"  resposta em formato inesperado: {e!r}")
            continue
        if serie:
            return sorted(serie)
        print("  resposta vazia.")
    sys.exit("Nao foi possivel obter a serie do IPCA (BCB e IBGE falharam).")


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
def trecho_planilha1(content):
    """Devolve (inicio, fim) do trecho do content.xml que pertence a Planilha1."""
    ini = content.index('table:name="Planilha1"')
    fim = content.find("<table:table ", ini)
    return ini, (fim if fim > 0 else len(content))


def ler_serie(caminho_ods):
    """Le a planilha e devolve [[AAAAMM, indice], ...] direto do XML do ODS."""
    with zipfile.ZipFile(caminho_ods) as z:
        content = z.read("content.xml").decode("utf-8")

    ini, fim = trecho_planilha1(content)
    trecho = content[ini:fim]

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

    # substitui o marcador e o array que o segue (seja [] ou uma serie antiga).
    # O array vai ate o primeiro "];" -- um "]" sozinho fecharia so o primeiro par.
    html, trocas = re.subn(re.escape(marcador) + r"\s*\[.*?\];",
                           lambda _: marcador + dados + ";", html, count=1, flags=re.S)
    if not trocas:
        print(f"Aviso: o array apos {marcador} em {MODELO} nao termina com '];'; "
              "a pagina HTML nao foi gerada.")
        return None

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
    serie = baixar_serie(fonte, mes_menos(ultimo_ano, ultimo_mes, JANELA_MESES))
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
        print("Aviso: caixa de texto do valor final nao encontrada na Planilha1.")

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
                   help="URL no formato do SGS ou JSON local (desliga a reserva do IBGE)")
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
