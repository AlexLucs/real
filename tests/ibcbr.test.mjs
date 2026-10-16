// Testa o bloco <script id="calc"> do ibcbr.html contra tests/fixture.json.
// Uso: node tests/ibcbr.test.mjs   (regere a fixture com tests/gerar_fixture.py)
import {readFileSync} from "node:fs";
import vm from "node:vm";
import assert from "node:assert/strict";

const raiz = new URL("../", import.meta.url);
const html = readFileSync(new URL("ibcbr.html", raiz), "utf8");
const codigo = html.match(/<script id="calc">([\s\S]*?)<\/script>/)[1];
const ctx = vm.createContext({Intl});
vm.runInContext(codigo + ";this.api={COLUNAS,converterSerie,calcularColuna,formatar,chaveMes,unirIntervalos,cobre,janelaNecessaria,faixaAoRedor};", ctx);
const {COLUNAS, converterSerie, calcularColuna, formatar, chaveMes, unirIntervalos, cobre, janelaNecessaria, faixaAoRedor} = ctx.api;
const simples = x => JSON.parse(JSON.stringify(x)); // arrays do vm têm outro protótipo
const fx = JSON.parse(readFileSync(new URL("tests/fixture.json", raiz), "utf8"));

let n = 0; const t = (nome, fn) => { fn(); n++; };

t("converterSerie descarta vazio, nulo, não numérico e não positivo", () => {
  const r = converterSerie([{data:"01/01/2020",valor:"97.61698"},{data:"01/02/2020",valor:""},
    {data:"01/03/2020",valor:null},{data:"01/04/2020",valor:" "},{data:"01/05/2020",valor:"abc"},
    {data:"01/06/2020",valor:"0"},{data:"x",valor:"1"},null]);
  assert.deepEqual({...r}, {"2020-01":97.61698});
});
t("chaveMes normaliza meses fora de 1..12", () => {
  assert.equal(chaveMes(2026,0),"2025-12"); assert.equal(chaveMes(2026,-23),"2024-01"); assert.equal(chaveMes(2026,13),"2027-01");
});
t("unirIntervalos junta sobrepostos e adjacentes, mantém separados os demais", () => {
  assert.deepEqual(simples(unirIntervalos([["2024-08","2026-07"],["2020-01","2024-07"],["2010-01","2012-12"]])),
    [["2010-01","2012-12"],["2020-01","2026-07"]]);
});
t("cobre exige a janela inteira dentro de um intervalo", () => {
  const iv = [["2010-01","2012-12"],["2020-01","2026-07"]];
  assert.equal(cobre(iv,"2024-08","2026-07"), true);
  assert.equal(cobre(iv,"2012-06","2014-05"), false);
  assert.equal(cobre([],"2024-08","2026-07"), false);
});
t("janelaNecessaria vai de m-23 a m", () => {
  assert.deepEqual(simples(janelaNecessaria(2026,7)), ["2024-08","2026-07"]);
});
t("faixaAoRedor cobre os 5 anos ao redor e respeita os limites", () => {
  assert.deepEqual(simples(faixaAoRedor(2015,"2026-07")), ["2011-01","2017-12"]);
  assert.deepEqual(simples(faixaAoRedor(2005,"2026-07")), ["2003-01","2007-12"]);
  assert.deepEqual(simples(faixaAoRedor(2025,"2026-07")), ["2021-01","2026-07"]);
  // todo mês de y-2 a y+2 fica coberto
  const f = faixaAoRedor(2015,"2026-07");
  for(let a = 2013; a <= 2017; a++) for(let m = 1; m <= 12; m++)
    assert.ok(cobre([f], ...janelaNecessaria(a,m)), `${m}/${a}`);
});
t("formatar: meio para longe do zero, -0 e não finitos", () => {
  assert.equal(formatar(-0.25),"-0,3"); assert.equal(formatar(0.25),"0,3"); assert.equal(formatar(-0.04),"0,0");
  assert.equal(formatar(null),"–"); assert.equal(formatar(Infinity),"–");
});
for(const caso of fx.casos){
  t(`tabela ${caso.mes}/${caso.ano}`, () => {
    COLUNAS.forEach((c, i) => {
      const r = calcularColuna(fx.series[c.codSA], fx.series[c.codNSA], caso.ano, caso.mes);
      const obtido = Object.fromEntries(Object.entries(r).map(([k,v]) => [k, formatar(v)]));
      assert.deepEqual(obtido, caso.colunas[i], `${c.nome} ${caso.mes}/${caso.ano}`);
    });
  });
}
t("série faltando vira null, não número", () => {
  const r = calcularColuna(undefined, fx.series[24363], 2026, 7);
  assert.equal(r.varMes, null); assert.notEqual(r.varMes12, null);
});
console.log(`ok: ${n} testes (${fx.origem})`);
