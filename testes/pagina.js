// Gera pela página (timbrador.html, no Chromium sem janela) o mesmo documento de uma especificação JSON.
// Uso: node testes/pagina.js PAGINA_COM_MARCA.html PAGINA_SEM_MARCA.html MARCA_PORTATIL.json ESPEC.json PASTA_SAIDA
// Confere também que a página sem marca só mostra o bloco da marca e que ela aceita o marca.json carregado.
const { chromium } = require("playwright");
const fs = require("fs");
(async () => {
  const [pagMarca, pagVazia, marca, espec, saida] = process.argv.slice(2);
  const nav = await chromium.launch();
  const erros = [];
  const abrir = async url => { const p = await nav.newPage(); p.on("pageerror", e => erros.push(e.message)); await p.goto("file://" + url); return p; };
  const vazia = await abrir(pagVazia);
  if (!(await vazia.$eval("#formulario", e => e.hidden))) erros.push("página sem marca mostrou o formulário");
  await vazia.setInputFiles("#marca-arquivo", marca);
  await vazia.waitForLoadState("load"); await vazia.waitForTimeout(300);
  if (await vazia.$eval("#formulario", e => e.hidden)) erros.push("página não aceitou o marca.json carregado");
  const p = await abrir(pagMarca);
  const r = await p.evaluate(async spec => {
    const conteudo = spec.corpo || []; delete spec.corpo;
    const v = validar(spec);
    if (v.erros.length) return { erro: v.erros.map(e => e.msg).join(" ") };
    const u = new Uint8Array(await gerarDocx(spec, conteudo).arrayBuffer());
    let bin = ""; for (const x of u) bin += String.fromCharCode(x);
    return { nome: nomeArquivo(spec), b64: btoa(bin) };
  }, JSON.parse(fs.readFileSync(espec, "utf8")));
  await nav.close();
  if (r.erro) erros.push("validar: " + r.erro);
  else fs.writeFileSync(`${saida}/${r.nome}`, Buffer.from(r.b64, "base64"));
  if (erros.length) { console.error(erros.join("\n")); process.exit(1); }
  console.log(`${saida}/${r.nome}`);
})();
