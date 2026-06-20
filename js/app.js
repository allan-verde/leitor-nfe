/* app.js — cola tudo: upload/drag&drop, leitura de .xml e .zip (JSZip),
   processamento (parse.js), preview na tela e download do .xlsx (excel.js). */
(() => {
  "use strict";
  const $ = s => document.querySelector(s);
  const drop = $("#drop"), fileInput = $("#file");
  let resultado = null; // guarda o último processamento p/ exportar

  const fmtMoeda = v => (v == null ? '' : Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
  const fmtPct = v => (v == null ? '' : (Number(v) * 100).toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + '%');
  const esc = s => (s == null ? '' : String(s)).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const setStatus = (html, cls = "") => { const s = $("#status"); s.className = cls; s.innerHTML = html; };

  // ---- eventos de upload ----
  $("#pick").onclick = e => { e.stopPropagation(); fileInput.click(); };
  drop.onclick = () => fileInput.click();
  fileInput.onchange = () => { if (fileInput.files.length) tratar(fileInput.files); };
  ["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("hover"); }));
  ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("hover"); }));
  drop.addEventListener("drop", e => { if (e.dataTransfer.files.length) tratar(e.dataTransfer.files); });

  // ---- lê os arquivos (xml direto ou dentro de zip) ----
  async function coletarXmls(fileList) {
    const arquivos = [];
    for (const f of fileList) {
      const nome = f.name.toLowerCase();
      if (nome.endsWith(".zip")) {
        const zip = await JSZip.loadAsync(await f.arrayBuffer());
        const entradas = Object.values(zip.files).filter(e => !e.dir && e.name.toLowerCase().endsWith(".xml"));
        for (const e of entradas) arquivos.push({ nome: e.name, conteudo: await e.async("string") });
      } else if (nome.endsWith(".xml")) {
        arquivos.push({ nome: f.name, conteudo: await f.text() });
      }
    }
    return arquivos;
  }

  async function tratar(fileList) {
    setStatus('<span class="spinner"></span> Lendo arquivo(s)…');
    $("#resultado").style.display = "none";
    try {
      const arquivos = await coletarXmls(fileList);
      if (!arquivos.length) { setStatus("Nenhum XML encontrado nos arquivos selecionados.", "erro"); return; }
      const res = NFE.processar(arquivos);
      if (!res.ok) { setStatus("Erro: " + esc(res.erro), "erro"); return; }
      resultado = res;
      render(res.preview);
      setStatus("Leitura concluída.", "ok");
    } catch (err) {
      setStatus("Erro: " + esc(err.message || err), "erro");
      console.error(err);
    }
  }

  // ---- render do preview ----
  const badge = n => ` <span class="badge">${n}</span>`;
  function render(p) {
    const id = p.identificacao;
    $("#ident").innerHTML = [
      ["Empresa", id.empresa], ["CNPJ", id.cnpj || '—'], ["IE", id.ie || '—'], ["UF", id.uf || '—'],
      ["Regime", id.regime || '—'], ["Período", id.periodo || '—'],
      ["Entradas / Saídas", id.n_entradas + " / " + id.n_saidas],
      ["Modelo 55 / 65", id.n_mod55 + " / " + id.n_mod65],
    ].map(([k, v]) => `<div><div class="k">${esc(k)}</div><div class="v">${esc(v)}</div></div>`).join("");

    $("#meta").textContent = `${p.arquivos_lidos} XML lido(s)` +
      (p.arquivos_com_erro ? ` · ${p.arquivos_com_erro} com erro` : "") +
      ` · ${id.data_ini || ""} a ${id.data_fim || ""}`;

    $("#avisoIndef").innerHTML = id.n_indefinidas
      ? `<div class="aviso"><b>${id.n_indefinidas}</b> nota(s) não puderam ser classificadas como entrada nem saída (empresa não consta como emitente nem destinatário). Confira a aba <b>Indefinidas</b>.</div>`
      : "";

    const tabs = [
      ["resumo", "Resumo", panelResumo(p)],
      ["entradas", "Entradas" + badge(p.entradas.length), panelNotas(p.entradas, "Fornecedor")],
      ["saidas", "Saídas" + badge(p.saidas.length), panelNotas(p.saidas, "Cliente")],
      ["class", "Classificação entradas", panelClass(p.classificacao_entradas)],
    ];
    if (id.n_indefinidas) tabs.push(["indef", "Indefinidas" + badge(p.indefinidas.length), panelIndef(p.indefinidas)]);

    $("#tabs").innerHTML = tabs.map((t, i) => `<div class="tab${i === 0 ? ' active' : ''}" data-t="${t[0]}">${t[1]}</div>`).join("");
    $("#panels").innerHTML = tabs.map((t, i) => `<div class="panel${i === 0 ? ' active' : ''}" id="panel-${t[0]}">${t[2]}</div>`).join("");
    document.querySelectorAll(".tab").forEach(tab => tab.onclick = () => {
      document.querySelectorAll(".tab").forEach(x => x.classList.remove("active"));
      document.querySelectorAll(".panel").forEach(x => x.classList.remove("active"));
      tab.classList.add("active"); $("#panel-" + tab.dataset.t).classList.add("active");
    });

    $("#resultado").style.display = "block";
  }

  function tabelaResumo(titulo, bloco) {
    const linhas = bloco.linhas.map(l => `<tr>
        <td>${esc(l.categoria)}</td><td class="num">${l.qtd}</td>
        <td class="num">${fmtMoeda(l.vProd)}</td><td class="num">${fmtMoeda(l.vNF)}</td></tr>`).join("");
    return `<table><thead><tr><th>${esc(titulo)}</th><th>Qtd notas</th><th>Valor produtos</th><th>Valor total (vNF)</th></tr></thead>
      <tbody>${linhas || '<tr><td colspan="4" class="muted">Sem notas.</td></tr>'}
      <tr class="total"><td>TOTAL</td><td class="num">${bloco.total.qtd}</td>
        <td class="num">${fmtMoeda(bloco.total.vProd)}</td><td class="num">${fmtMoeda(bloco.total.vNF)}</td></tr></tbody></table>`;
  }
  function panelResumo(p) {
    return `<div class="scroll" style="padding:0">
      <div style="padding:12px 14px"><b style="color:var(--navy)">COMPRAS (ENTRADAS)</b></div>${tabelaResumo("Natureza (categoria)", p.resumo_entradas)}
      <div style="padding:16px 14px 12px"><b style="color:var(--navy)">VENDAS (SAÍDAS)</b></div>${tabelaResumo("Natureza (categoria)", p.resumo_saidas)}
    </div>`;
  }
  function panelNotas(notas, rotulo) {
    if (!notas.length) return `<div class="muted">Nenhuma nota.</div>`;
    const body = notas.map(n => `<tr>
      <td class="ctr">${esc(n.nNF)}</td><td class="ctr">${esc(n.serie)}</td><td class="ctr">${esc(n.data)}</td>
      <td class="ctr">${esc(n.mod)}</td><td>${esc(n.contraparte)}</td><td class="ctr">${esc(n.uf)}</td>
      <td>${esc(n.cfop)}</td><td>${esc(n.categoria)}</td>
      <td class="num">${fmtMoeda(n.vProd)}</td><td class="num">${fmtMoeda(n.vNF)}</td>
      <td class="num">${fmtMoeda(n.vICMS)}</td><td class="num">${fmtMoeda(n.vIPI)}</td></tr>`).join("");
    const tot = k => notas.reduce((s, n) => s + (n[k] || 0), 0);
    return `<div class="scroll"><table><thead><tr>
      <th>NF</th><th>Série</th><th>Data</th><th>Mod</th><th>${esc(rotulo)}</th><th>UF</th>
      <th>CFOP</th><th>Categoria</th><th>Vlr Produtos</th><th>Vlr NF</th><th>ICMS</th><th>IPI</th></tr></thead>
      <tbody>${body}<tr class="total"><td colspan="8">TOTAL (${notas.length})</td>
        <td class="num">${fmtMoeda(tot('vProd'))}</td><td class="num">${fmtMoeda(tot('vNF'))}</td>
        <td class="num">${fmtMoeda(tot('vICMS'))}</td><td class="num">${fmtMoeda(tot('vIPI'))}</td></tr></tbody></table></div>`;
  }
  function panelClass(linhas) {
    if (!linhas.length) return `<div class="muted">Sem itens de entrada para classificar.</div>`;
    const cor = c => c.startsWith('Compra') ? 'cls-green' : (c.startsWith('Bonific') ? 'cls-yellow' : (c.startsWith('Uso') ? 'cls-red' : ''));
    const body = linhas.map(l => `<tr>
      <td class="${cor(l.classificacao)}">${esc(l.classificacao)}</td>
      <td class="num">${l.qtd}</td><td class="num">${fmtMoeda(l.vProd)}</td><td class="num">${fmtPct(l.pct)}</td></tr>`).join("");
    return `<div class="scroll"><table><thead><tr><th>Classificação</th><th>Qtd itens</th><th>Valor produtos</th><th>% do total</th></tr></thead>
      <tbody>${body}</tbody></table></div>
      <div class="muted">A destinação de itens em bonificação (revenda × uso/consumo) deve ser confirmada pela contabilidade — o XML não declara a finalidade.</div>`;
  }
  function panelIndef(notas) {
    const body = notas.map(n => `<tr><td class="ctr">${esc(n.nNF)}</td><td>${esc(n.emit)}</td>
       <td>${esc(n.dest)}</td><td class="ctr">${esc(n.tpNF)}</td><td>${esc(n.chave)}</td></tr>`).join("");
    return `<div class="scroll"><table><thead><tr><th>NF</th><th>Emitente</th><th>Destinatário</th><th>tpNF</th><th>Chave</th></tr></thead>
      <tbody>${body}</tbody></table></div>`;
  }

  // ---- exportar ----
  $("#exportar").onclick = async () => {
    if (!resultado) return;
    const btn = $("#exportar"); const txt = btn.textContent;
    btn.disabled = true; btn.textContent = "Gerando…";
    try {
      const { entradas, saidas, empresa, periodo } = resultado;
      const blob = await NFEExcel.gerarBlob(entradas, saidas, empresa, periodo.per, periodo.ini, periodo.fim);
      const nomeEmp = (empresa.nome || "Empresa").split(/\s+/).slice(0, 3).join("_");
      const nome = ("Relatorio_" + nomeEmp + (periodo.per ? "_" + periodo.per : "") + ".xlsx").replace(/\//g, "-");
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = nome; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch (err) {
      setStatus("Erro ao gerar a planilha: " + esc(err.message || err), "erro");
      console.error(err);
    } finally {
      btn.disabled = false; btn.textContent = txt;
    }
  };
  $("#novo").onclick = () => { resultado = null; $("#resultado").style.display = "none"; fileInput.value = ""; setStatus(""); window.scrollTo(0, 0); };
})();
