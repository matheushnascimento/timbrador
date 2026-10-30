// Relatório de condução pela página (timbrador.html, no Chromium sem janela), sobre arquivos do gerador.
// Uso: node testes/pagina_rel.js PAGINA_COM_MARCA.html REL_v1.1.docx REL_ENCERRADO.docx LISTA.json PASTA_SAIDA
// Confere que a página lê o quadro e o SHA-256 das rodadas como o gerador, recusa rodada no relatório encerrado
// e em linha alterada, e grava em PASTA_SAIDA/rel.docx a v1.2 feita a partir da v1.1 com uma rodada nova.
const { chromium } = require("playwright");
const fs = require("fs");
(async () => {
  const [pagina, v11, encerrado, lista, saida] = process.argv.slice(2);
  const nav = await chromium.launch();
  const p = await nav.newPage();
  const erros = [];
  p.on("pageerror", e => erros.push(e.message));
  await p.goto("file://" + pagina);
  const b64 = f => fs.readFileSync(f).toString("base64");
  const r = await p.evaluate(async ([v11, fim, lista]) => {
    const buf = s => Uint8Array.from(atob(s), c => c.charCodeAt(0)).buffer;
    const falhas = [], copia = o => JSON.parse(JSON.stringify(o));
    const reg = lista.documentos.find(d => d.codigo === "ACME-REL-ECO-001");
    const e = await lerDocx(buf(fim));
    if (conferirQuadroRodadas(e.blocos, e.ficha.rodadas).length) falhas.push("quadro do relatório do gerador não confere com a ficha");
    if ((await digitaisRodadas(e.ficha.rodadas)).join() !== reg.rodadas_sha256.join()) falhas.push("SHA-256 das rodadas difere do gerador");
    const nova = { n: 3, tipo: "rodada", data: "21/10/2026", feito: "x", avanco: "sim", compromisso: { oque: "a", quem: "b", ate: "c" } };
    let s = { ...copia(e.ficha), rodadas: e.ficha.rodadas.concat([nova]) }; delete s.codigo;
    if (!(await regrasRel(s, lista, e.ficha)).erros.some(m => m.includes("foi encerrado"))) falhas.push("aceitou rodada no relatório encerrado");
    // a partir da v1.1, com a lista como estava na v1.1
    const a = await lerDocx(buf(v11)), l11 = copia(lista), r11 = l11.documentos.find(d => d.codigo === "ACME-REL-ECO-001");
    r11.rodadas_sha256 = r11.rodadas_sha256.slice(0, 2); r11.versao = "1.1"; delete r11.encerrado_por;
    const r2 = { n: 2, tipo: "rodada", data: "15/10/2026", feito: "Frete ainda não ativado.", avanco: "nao", travou: "contrato", depende: "Fulano",
      compromisso: { oque: "Ativar frete", quem: "Logística", ate: "22/10/2026" } };
    const alterada = copia(a.ficha); alterada.rodadas[1].feito = "60 anúncios cadastrados.";
    s = { ...copia(alterada), rodadas: alterada.rodadas.concat([r2]) }; delete s.codigo;
    if (!(await regrasRel(s, l11, alterada)).erros.some(m => m.includes("não confere com a registrada"))) falhas.push("aceitou rodada publicada alterada");
    s = { ...copia(a.ficha), rodadas: a.ficha.rodadas.concat([r2]), estado: "OFICIAL", data_aprovacao: r2.data, data_emissao: r2.data }; delete s.codigo;
    s.historico = historicoAtualizado(a.ficha, versaoRodada(a.ficha.versao), resumoRodada(r2), s.aprovado_por, r2.data);
    const rr = await regrasRel(s, l11, a.ficha);
    if (rr.erros.length) return { falhas: falhas.concat(rr.erros) };
    if (!rr.avisos.length && s.versao !== "1.2") falhas.push("versão da rodada");
    const v = validar(s);
    if (v.erros.length) return { falhas: falhas.concat(v.erros.map(x => x.msg)) };
    const u = new Uint8Array(await gerarDocx(s, corpoRel(a.blocos, s)).arrayBuffer());
    let bin = ""; for (const x of u) bin += String.fromCharCode(x);
    return { falhas, b64: btoa(bin), versao: s.versao, codigo: s.codigo };
  }, [b64(v11), b64(encerrado), JSON.parse(fs.readFileSync(lista, "utf8"))]);
  await nav.close();
  erros.push(...(r.falhas || []));
  if (r.b64) fs.writeFileSync(`${saida}/rel.docx`, Buffer.from(r.b64, "base64"));
  if (r.b64 && (r.versao !== "1.2" || r.codigo !== "ACME-REL-ECO-001")) erros.push(`gerou ${r.codigo} v${r.versao}`);
  if (erros.length) { console.error(erros.join("\n")); process.exit(1); }
  console.log(`${saida}/rel.docx`);
})();
