/* excel.js — gera o relatório .xlsx de 7 abas com ExcelJS, no navegador.
   Porte fiel de gerar() (openpyxl) de nfe_core.py: mesmo estilo NAVY/LIGHT,
   fórmulas SUMIF/COUNTIF/SUM, mesclagens, freeze panes e formatos. */

const NFEExcel = (() => {
  "use strict";

  // ARGB (ExcelJS usa FF + RRGGBB)
  const NAVY = 'FF1F3864', LIGHT = 'FFDDEBF7', GREY = 'FF808080', WHITE = 'FFFFFFFF', FONT = 'Arial';
  const CUR = '#,##0.00;(#,##0.00);"-"', PCT = '0.0%';

  const fillSolid = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
  const BORDER = { top: thin, left: thin, bottom: thin, right: thin };
  const C = { horizontal: 'center', vertical: 'middle' };
  const R = { horizontal: 'right', vertical: 'middle' };
  const L = { horizontal: 'left', vertical: 'middle' };

  const { categoriaCfop, classificaItem, fmtDoc, CRT, ORDEM_CAT, ORDEM_CLASS } = NFE;
  const iNum = v => parseInt(v, 10) || 0;

  function addSheet(wb, name, view) {
    return wb.addWorksheet(name, { views: [Object.assign({ showGridLines: false }, view || {})] });
  }
  function setw(ws, widths) { widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; }); }
  function hb(ws, row, n, start = 1) {
    for (let c = start; c < start + n; c++) {
      const x = ws.getCell(row, c);
      x.font = { name: FONT, color: { argb: WHITE }, bold: true, size: 10 };
      x.fill = fillSolid(NAVY); x.alignment = C; x.border = BORDER;
    }
  }
  const F = (formula) => ({ formula });

  function gerar(entradas, saidas, empresa, per, ini, fim) {
    const wb = new ExcelJS.Workbook();
    const n_ent = entradas.length, n_sai = saidas.length, tot_docs = n_ent + n_sai;
    const crt_lbl = empresa.crt ? (CRT[empresa.crt] || '') : '';
    const sortByDt = lst => [...lst].sort((a, b) => (a.dhEmi || '').localeCompare(b.dhEmi || ''));

    // ===== Cria as abas na ordem final de exibição =====
    const ws = addSheet(wb, 'Resumo');
    const wd = addSheet(wb, 'Detalhe NF', { state: 'frozen', xSplit: 2, ySplit: 3 });
    const wc = addSheet(wb, 'Consolidado NF', { state: 'frozen', xSplit: 0, ySplit: 3 });
    const wcl = addSheet(wb, 'Classificação Entradas');
    const we = addSheet(wb, 'Entradas', { state: 'frozen', xSplit: 2, ySplit: 3 });
    const ws_ = addSheet(wb, 'Saídas', { state: 'frozen', xSplit: 2, ySplit: 3 });
    const wi = addSheet(wb, 'Itens', { state: 'frozen', xSplit: 3, ySplit: 3 });

    // ===== ENTRADAS =====
    we.getCell(1, 1).value = `NOTAS DE ENTRADA — ${per}`;
    we.getCell(1, 1).font = { name: FONT, bold: true, size: 12, color: { argb: NAVY } };
    const cc = ["NF","Série","Data","Chave de acesso","Fornecedor","UF","CFOP","Categoria","Vlr Produtos","Vlr NF","ICMS","IPI"];
    cc.forEach((h, j) => we.getCell(3, j + 1).value = h);
    hb(we, 3, cc.length);
    let rw = 4;
    for (const r of sortByDt(entradas)) {
      const cat = categoriaCfop(r.cfop_dom, 'entrada');
      [iNum(r.nNF), iNum(r.serie), (r.dhEmi||'').slice(0,10), r.chave, r.emit_nome, r.emit_uf,
       r.cfops.join(','), cat, r.vProd, r.vNF, r.vICMS, r.vIPI].forEach((v, j) => we.getCell(rw, j + 1).value = v);
      rw++;
    }
    const e_last = rw - 1;
    we.getCell(rw, 8).value = "TOTAL"; we.getCell(rw, 8).font = { name: FONT, bold: true };
    [9,10,11,12].forEach(col => we.getCell(rw, col).value = F(`SUM(${cl(col)}4:${cl(col)}${e_last})`));
    for (let c = 1; c <= cc.length; c++) { we.getCell(rw, c).fill = fillSolid(LIGHT); we.getCell(rw, c).font = { name: FONT, bold: true }; }
    estiloLinhas(we, 4, rw, [9,10,11,12], [1,2,3,6,7,8], cc.length);
    setw(we, [8,6,11,38,34,5,14,26,13,13,12,12]);

    // ===== SAÍDAS =====
    ws_.getCell(1, 1).value = `NOTAS DE SAÍDA — ${per}`;
    ws_.getCell(1, 1).font = { name: FONT, bold: true, size: 12, color: { argb: NAVY } };
    const cs = ["NF","Série","Data","Chave de acesso","Cliente","UF","CFOP","Categoria","Vlr Produtos","Vlr NF","ICMS","IPI"];
    cs.forEach((h, j) => ws_.getCell(3, j + 1).value = h);
    hb(ws_, 3, cs.length);
    rw = 4;
    for (const r of sortByDt(saidas)) {
      const cat = categoriaCfop(r.cfop_dom, 'saida');
      [iNum(r.nNF), iNum(r.serie), (r.dhEmi||'').slice(0,10), r.chave, r.dest_nome, r.dest_uf,
       r.cfops.join(','), cat, r.vProd, r.vNF, r.vICMS, r.vIPI].forEach((v, j) => ws_.getCell(rw, j + 1).value = v);
      rw++;
    }
    const s_last = rw - 1;
    ws_.getCell(rw, 8).value = "TOTAL"; ws_.getCell(rw, 8).font = { name: FONT, bold: true };
    [9,10,11,12].forEach(col => ws_.getCell(rw, col).value = F(`SUM(${cl(col)}4:${cl(col)}${s_last})`));
    for (let c = 1; c <= cs.length; c++) { ws_.getCell(rw, c).fill = fillSolid(LIGHT); ws_.getCell(rw, c).font = { name: FONT, bold: true }; }
    estiloLinhas(ws_, 4, rw, [9,10,11,12], [1,2,3,6,7,8], cs.length);
    setw(ws_, [8,6,11,38,34,5,14,26,13,13,12,12]);

    // ===== RESUMO =====
    ws.getCell('A1').value = "RELATÓRIO DE COMPRAS E VENDAS";
    ws.getCell('A1').font = { name: FONT, bold: true, size: 14, color: { argb: NAVY } };
    let sub = `${empresa.nome} — CNPJ ${fmtDoc(empresa.cnpj)}`; if (crt_lbl) sub += ` — ${crt_lbl}`;
    ws.getCell('A2').value = sub; ws.getCell('A2').font = { name: FONT, size: 9, color: { argb: GREY } };
    ws.getCell('A3').value = `Período: ${ini} a ${fim}  (${per})  —  ${tot_docs} NF-e (${n_ent} entradas + ${n_sai} saídas)`;
    ws.getCell('A3').font = { name: FONT, size: 9, color: { argb: GREY } };
    const heads = ["Natureza (categoria)","Qtd notas","Valor produtos","Valor total (vNF)"];

    const presentesCat = (notas, lado) => {
      const setPresentes = new Set(notas.map(r => categoriaCfop(r.cfop_dom, lado)));
      const base = ORDEM_CAT.filter(c => setPresentes.has(c));
      const extra = [...setPresentes].filter(c => !base.includes(c)).sort();
      return base.concat(extra);
    };
    const cats_ent = presentesCat(entradas, 'entrada');
    const cats_sai = presentesCat(saidas, 'saida');

    let r0 = 5;
    ws.getCell(r0, 1).value = "COMPRAS (ENTRADAS)"; ws.getCell(r0, 1).font = { name: FONT, bold: true, size: 11, color: { argb: NAVY } };
    heads.forEach((h, j) => ws.getCell(r0 + 1, j + 1).value = h); hb(ws, r0 + 1, 4);
    let rr = r0 + 2; const f_e = rr;
    for (const cat of cats_ent) {
      ws.getCell(rr, 1).value = cat;
      ws.getCell(rr, 2).value = F(`COUNTIF(Entradas!$H$4:$H$${e_last},A${rr})`);
      ws.getCell(rr, 3).value = F(`SUMIF(Entradas!$H$4:$H$${e_last},A${rr},Entradas!$I$4:$I$${e_last})`);
      ws.getCell(rr, 4).value = F(`SUMIF(Entradas!$H$4:$H$${e_last},A${rr},Entradas!$J$4:$J$${e_last})`);
      rr++;
    }
    const l_e = rr - 1;
    ws.getCell(rr, 1).value = "TOTAL ENTRADAS";
    [[2,'B'],[3,'C'],[4,'D']].forEach(([col, Lr]) => ws.getCell(rr, col).value = F(`SUM(${Lr}${f_e}:${Lr}${l_e})`));
    const e_sum_tot = rr;
    for (let c = 1; c <= 4; c++) { ws.getCell(rr, c).fill = fillSolid(LIGHT); ws.getCell(rr, c).font = { name: FONT, bold: true }; }

    const r1 = rr + 3;
    ws.getCell(r1, 1).value = "VENDAS (SAÍDAS)"; ws.getCell(r1, 1).font = { name: FONT, bold: true, size: 11, color: { argb: NAVY } };
    heads.forEach((h, j) => ws.getCell(r1 + 1, j + 1).value = h); hb(ws, r1 + 1, 4);
    rr = r1 + 2; const f_s = rr;
    for (const cat of cats_sai) {
      ws.getCell(rr, 1).value = cat;
      ws.getCell(rr, 2).value = F(`COUNTIF(Saídas!$H$4:$H$${s_last},A${rr})`);
      ws.getCell(rr, 3).value = F(`SUMIF(Saídas!$H$4:$H$${s_last},A${rr},Saídas!$I$4:$I$${s_last})`);
      ws.getCell(rr, 4).value = F(`SUMIF(Saídas!$H$4:$H$${s_last},A${rr},Saídas!$J$4:$J$${s_last})`);
      rr++;
    }
    const l_s = rr - 1;
    ws.getCell(rr, 1).value = "TOTAL SAÍDAS";
    [[2,'B'],[3,'C'],[4,'D']].forEach(([col, Lr]) => ws.getCell(rr, col).value = F(`SUM(${Lr}${f_s}:${Lr}${l_s})`));
    const s_sum_tot = rr;
    for (let c = 1; c <= 4; c++) { ws.getCell(rr, c).fill = fillSolid(LIGHT); ws.getCell(rr, c).font = { name: FONT, bold: true }; }

    const r2 = rr + 3;
    ws.getCell(r2, 1).value = "INDICADORES"; ws.getCell(r2, 1).font = { name: FONT, bold: true, size: 11, color: { argb: NAVY } };
    const inds = [["Total de entradas (vNF)", `D${e_sum_tot}`], ["Total de saídas (vNF)", `D${s_sum_tot}`],
      ["Total de produtos comprados (R$)", `C${e_sum_tot}`], ["Total de produtos vendidos (R$)", `C${s_sum_tot}`],
      ["ICMS destacado nas saídas (R$)", `SUM(Saídas!K4:K${s_last})`], ["IPI destacado nas entradas (R$)", `SUM(Entradas!L4:L${e_last})`]];
    inds.forEach(([lbl, val], i) => {
      const rw2 = r2 + 1 + i; ws.getCell(rw2, 1).value = lbl; ws.getCell(rw2, 1).font = { name: FONT, size: 10 };
      const x = ws.getCell(rw2, 2); x.value = F(val); x.numFmt = CUR; x.alignment = R; x.font = { name: FONT, size: 10 };
    });
    for (const [a, b] of [[r0 + 2, e_sum_tot], [r1 + 2, s_sum_tot]]) {
      for (let rw2 = a; rw2 <= b; rw2++) {
        ws.getCell(rw2, 2).alignment = C;
        for (const cc2 of [3, 4]) { ws.getCell(rw2, cc2).numFmt = CUR; ws.getCell(rw2, cc2).alignment = R; }
        for (let cc2 = 1; cc2 <= 4; cc2++) ws.getCell(rw2, cc2).border = BORDER;
        ws.getCell(rw2, 1).alignment = L;
      }
    }
    setw(ws, [34, 12, 18, 18]);

    // ===== DETALHE NF =====
    wd.getCell(1, 1).value = "DETALHAMENTO COMPLETO POR NOTA FISCAL";
    wd.getCell(1, 1).font = { name: FONT, bold: true, size: 12, color: { argb: NAVY } };
    const cols = ["Origem","NF","Série","Chave de acesso","Protocolo","Emissão","Natureza da operação","Emitente",
      "CNPJ/IE emit.","Destinatário / Cliente","CPF/CNPJ dest.","UF","Município","Vlr Produtos","Desconto",
      "Frete","Seguro","Outros","BC ICMS","ICMS","IPI","PIS","COFINS","ICMS ST","VALOR NF","Pagamento",
      "Modalidade frete","Transportadora"];
    cols.forEach((h, j) => wd.getCell(3, j + 1).value = h); hb(wd, 3, cols.length);
    rw = 4;
    const rowNf = (origem, r) => {
      const pag = r.pag.length ? r.pag.map(p => `${p[0]} R$ ${fmt2(p[1])}`).join("; ") : "";
      const vals = [origem, iNum(r.nNF), iNum(r.serie), r.chave, r.protocolo, (r.dhEmi||'').slice(0,10), r.natOp,
        r.emit_nome, `${fmtDoc(r.emit_cnpj)} / ${r.emit_ie || '-'}`, r.dest_nome, fmtDoc(r.dest_doc),
        r.dest_uf, r.dest_mun || '', r.vProd, r.vDesc, r.vFrete, r.vSeg, r.vOutro, r.vBC,
        r.vICMS, r.vIPI, r.vPIS, r.vCOFINS, r.vST, r.vNF, pag, r.modFrete || '', r.transp_nome || ''];
      vals.forEach((v, j) => wd.getCell(rw, j + 1).value = v);
    };
    for (const r of sortByDt(entradas)) { rowNf("Entrada", r); rw++; }
    for (const r of sortByDt(saidas)) { rowNf("Saída", r); rw++; }
    const d_last = rw - 1;
    wd.getCell(rw, 13).value = "TOTAL"; wd.getCell(rw, 13).font = { name: FONT, bold: true };
    for (let col = 14; col <= 25; col++) wd.getCell(rw, col).value = F(`SUM(${cl(col)}4:${cl(col)}${d_last})`);
    for (let c = 1; c <= cols.length; c++) { wd.getCell(rw, c).fill = fillSolid(LIGHT); wd.getCell(rw, c).font = { name: FONT, bold: true }; }
    estiloLinhas(wd, 4, rw, range(14, 25), [1,2,3,5,6,12], cols.length);
    setw(wd, [9,7,6,38,16,11,30,26,22,28,18,5,16,13,11,10,9,10,13,12,11,10,11,11,14,22,22,20]);

    // ===== CONSOLIDADO NF =====
    wc.getCell(1, 1).value = "CONSOLIDADO POR NOTA FISCAL — PRODUTOS";
    wc.getCell(1, 1).font = { name: FONT, bold: true, size: 12, color: { argb: NAVY } };
    wc.getCell(2, 1).value = "Cada nota apresenta cabeçalho, produtos e subtotal.";
    wc.getCell(2, 1).font = { name: FONT, size: 9, color: { argb: GREY } };
    const pc = ["Produto","NCM","CFOP","CST/CSOSN","Un","Qtde","Vlr Unit.","Vlr Produto","ICMS","IPI","PIS","COFINS"];
    setw(wc, [40,11,7,10,5,8,13,14,12,11,10,10]); const NCOL = pc.length;
    rw = 4;
    const grava = (origem, r) => {
      const rotulo = origem === 'Entrada' ? 'Fornecedor' : 'Cliente';
      const nome = (origem === 'Entrada' ? r.emit_nome : r.dest_nome) || (origem === 'Saída' ? 'Consumidor final' : '');
      const uf = (origem === 'Entrada' ? r.emit_uf : r.dest_uf) || '';
      const tit = `${origem.toUpperCase()}  •  NF ${iNum(r.nNF)}/${iNum(r.serie)}  •  ${(r.dhEmi||'').slice(0,10)}  •  `
        + `${rotulo}: ${nome} (${uf})  •  Nat.: ${r.natOp}  •  Valor NF: R$ ${fmt2(r.vNF)}`;
      wc.getCell(rw, 1).value = tit; wc.getCell(rw, 1).font = { name: FONT, bold: true, size: 10, color: { argb: WHITE } };
      for (let c = 1; c <= NCOL; c++) { wc.getCell(rw, c).fill = fillSolid(NAVY); wc.getCell(rw, c).border = BORDER; }
      wc.mergeCells(rw, 1, rw, NCOL); wc.getCell(rw, 1).alignment = L; rw++;
      const s = `Chave: ${r.chave}   |   Protocolo: ${r.protocolo}   |   Frete: ${r.modFrete || '-'}   |   Pagto: `
        + (r.pag.length ? r.pag.map(p => `${p[0]} R$ ${fmt2(p[1])}`).join("; ") : '-');
      wc.getCell(rw, 1).value = s; wc.getCell(rw, 1).font = { name: FONT, size: 8, italic: true, color: { argb: 'FFD9E1F2' } };
      for (let c = 1; c <= NCOL; c++) wc.getCell(rw, c).fill = fillSolid(NAVY);
      wc.mergeCells(rw, 1, rw, NCOL); wc.getCell(rw, 1).alignment = L; rw++;
      pc.forEach((h, j) => wc.getCell(rw, j + 1).value = h); hb(wc, rw, NCOL); rw++;
      const first = rw;
      for (const it of r.itens) {
        [it.xProd, it.NCM, it.CFOP, it.CST_CSOSN, it.uCom, it.qCom, it.vUnCom,
         F(`F${rw}*G${rw}`), it.vICMS, it.vIPI, it.vPIS, it.vCOFINS].forEach((v, j) => wc.getCell(rw, j + 1).value = v);
        wc.getCell(rw, 6).numFmt = '#,##0'; wc.getCell(rw, 6).alignment = R;
        for (const col of [7,8,9,10,11,12]) { wc.getCell(rw, col).numFmt = CUR; wc.getCell(rw, col).alignment = R; }
        for (const col of [2,3,4,5]) wc.getCell(rw, col).alignment = C;
        for (let c = 1; c <= NCOL; c++) { wc.getCell(rw, c).border = BORDER; wc.getCell(rw, c).font = { name: FONT, size: 9 }; }
        rw++;
      }
      const last = rw - 1;
      wc.getCell(rw, 1).value = `Subtotal NF ${iNum(r.nNF)}`; wc.getCell(rw, 1).font = { name: FONT, bold: true, size: 9 };
      wc.getCell(rw, 6).value = F(`SUM(F${first}:F${last})`);
      [[8,'H'],[9,'I'],[10,'J'],[11,'K'],[12,'L']].forEach(([col, Lr]) => wc.getCell(rw, col).value = F(`SUM(${Lr}${first}:${Lr}${last})`));
      wc.getCell(rw, 6).numFmt = '#,##0'; wc.getCell(rw, 6).alignment = R;
      for (const col of [8,9,10,11,12]) { wc.getCell(rw, col).numFmt = CUR; wc.getCell(rw, col).alignment = R; }
      for (let c = 1; c <= NCOL; c++) { wc.getCell(rw, c).fill = fillSolid(LIGHT); wc.getCell(rw, c).border = BORDER; wc.getCell(rw, c).font = { name: FONT, bold: true, size: 9 }; }
      rw += 2;
    };
    for (const r of sortByDt(entradas)) grava("Entrada", r);
    for (const r of sortByDt(saidas)) grava("Saída", r);

    // ===== CLASSIFICAÇÃO ENTRADAS =====
    wcl.getCell(1, 1).value = "CLASSIFICAÇÃO FISCAL DAS ENTRADAS";
    wcl.getCell(1, 1).font = { name: FONT, bold: true, size: 12, color: { argb: NAVY } };
    wcl.getCell(2, 1).value = "Classificação por item (uma nota pode conter finalidades distintas). Critério: CFOP + descrição.";
    wcl.getCell(2, 1).font = { name: FONT, size: 9, color: { argb: GREY } };
    wcl.getCell(4, 1).value = "RESUMO POR CLASSIFICAÇÃO";
    wcl.getCell(4, 1).font = { name: FONT, bold: true, size: 11, color: { argb: NAVY } };
    ["Classificação","Qtd itens","Valor produtos","% do total"].forEach((h, j) => wcl.getCell(5, j + 1).value = h);
    hb(wcl, 5, 4);
    const agg = new Map();
    for (const r of entradas) for (const it of r.itens) {
      const k = classificaItem(it); const a = agg.get(k) || [0, 0]; a[0]++; a[1] += it.vProd; agg.set(k, a);
    }
    const presentesCl = ORDEM_CLASS.filter(k => agg.has(k)).concat([...agg.keys()].filter(k => !ORDEM_CLASS.includes(k)));
    rr = 6; const f_sum = rr;
    for (const k of presentesCl) { wcl.getCell(rr, 1).value = k; wcl.getCell(rr, 2).value = agg.get(k)[0]; wcl.getCell(rr, 3).value = round2(agg.get(k)[1]); rr++; }
    const l_sum = rr - 1;
    wcl.getCell(rr, 1).value = "TOTAL ENTRADAS"; wcl.getCell(rr, 1).font = { name: FONT, bold: true };
    wcl.getCell(rr, 2).value = F(`SUM(B${f_sum}:B${l_sum})`); wcl.getCell(rr, 3).value = F(`SUM(C${f_sum}:C${l_sum})`);
    const tr = rr;
    for (let i = f_sum; i < rr; i++) { wcl.getCell(i, 4).value = F(`C${i}/$C$${tr}`); wcl.getCell(i, 4).numFmt = PCT; }
    wcl.getCell(rr, 4).value = F(`C${rr}/$C$${tr}`); wcl.getCell(rr, 4).numFmt = PCT;
    for (let i = f_sum; i <= rr; i++) {
      wcl.getCell(i, 2).alignment = C; wcl.getCell(i, 3).numFmt = CUR; wcl.getCell(i, 3).alignment = R; wcl.getCell(i, 4).alignment = R;
      for (let c = 1; c <= 4; c++) { wcl.getCell(i, c).border = BORDER; if (i === rr) { wcl.getCell(i, c).fill = fillSolid(LIGHT); wcl.getCell(i, c).font = { name: FONT, bold: true }; } }
    }
    const base = rr + 3;
    wcl.getCell(base - 1, 1).value = "DETALHE ITEM A ITEM"; wcl.getCell(base - 1, 1).font = { name: FONT, bold: true, size: 11, color: { argb: NAVY } };
    const dc = ["NF","Data","Produto","NCM","CFOP","Natureza","Qtde","Vlr Produto","Classificação"];
    dc.forEach((h, j) => wcl.getCell(base, j + 1).value = h); hb(wcl, base, dc.length);
    rw = base + 1; const fi = rw;
    const fills = { 'Compra para comercialização':'FFE2EFDA','Compra para industrialização':'FFE2EFDA',
      'Bonificação para revenda':'FFFFF2CC','Uso e consumo / brinde':'FFFCE4D6','Uso e consumo':'FFFCE4D6',
      'Ativo imobilizado':'FFDDD9C4','Demonstração (não é compra)':'FFD9E1F2' };
    for (const r of sortByDt(entradas)) for (const it of r.itens) {
      const cls = classificaItem(it);
      [iNum(r.nNF), (r.dhEmi||'').slice(0,10), it.xProd, it.NCM, it.CFOP, r.natOp, it.qCom, it.vProd, cls].forEach((v, j) => wcl.getCell(rw, j + 1).value = v);
      wcl.getCell(rw, 7).numFmt = '#,##0'; wcl.getCell(rw, 7).alignment = R;
      wcl.getCell(rw, 8).numFmt = CUR; wcl.getCell(rw, 8).alignment = R;
      for (const col of [1,2,4,5]) wcl.getCell(rw, col).alignment = C;
      for (let c = 1; c <= dc.length; c++) { wcl.getCell(rw, c).border = BORDER; wcl.getCell(rw, c).font = { name: FONT, size: 9 }; }
      if (fills[cls]) wcl.getCell(rw, 9).fill = fillSolid(fills[cls]);
      rw++;
    }
    const li = rw - 1;
    wcl.getCell(rw, 7).value = "TOTAL"; wcl.getCell(rw, 7).font = { name: FONT, bold: true };
    wcl.getCell(rw, 8).value = F(`SUM(H${fi}:H${li})`);
    for (let c = 1; c <= dc.length; c++) { wcl.getCell(rw, c).fill = fillSolid(LIGHT); wcl.getCell(rw, c).font = { name: FONT, bold: true }; wcl.getCell(rw, c).border = BORDER; }
    wcl.getCell(rw, 8).numFmt = CUR; wcl.getCell(rw, 8).alignment = R;
    wcl.getCell(rw + 2, 1).value = "Observação: a destinação final dos itens em bonificação (revenda x uso/consumo) deve ser confirmada pela contabilidade,";
    wcl.getCell(rw + 2, 1).font = { name: FONT, size: 8, italic: true, color: { argb: GREY } };
    wcl.getCell(rw + 3, 1).value = "pois o XML não declara a finalidade de uso pelo adquirente.";
    wcl.getCell(rw + 3, 1).font = { name: FONT, size: 8, italic: true, color: { argb: GREY } };
    setw(wcl, [8,11,40,11,7,30,8,14,28]);
    wcl.views = [{ showGridLines: false, state: 'frozen', xSplit: 0, ySplit: base }];

    // ===== ITENS =====
    wi.getCell(1, 1).value = "ITENS DETALHADOS COM IMPOSTOS";
    wi.getCell(1, 1).font = { name: FONT, bold: true, size: 12, color: { argb: NAVY } };
    const ic = ["Origem","NF","Item","Produto","NCM","CFOP","CST/CSOSN","Orig","Un","Qtde","Vlr Unit.","Vlr Produto",
      "Desconto","BC ICMS","% ICMS","Vlr ICMS","CST IPI","% IPI","Vlr IPI","CST PIS","Vlr PIS","CST COF","Vlr COFINS"];
    ic.forEach((h, j) => wi.getCell(3, j + 1).value = h); hb(wi, 3, ic.length);
    rw = 4;
    for (const [origem, lst] of [["Entrada", entradas], ["Saída", saidas]]) {
      for (const r of sortByDt(lst)) for (const it of r.itens) {
        [origem, iNum(r.nNF), iNum(it.nItem), it.xProd, it.NCM, it.CFOP, it.CST_CSOSN, it.orig, it.uCom, it.qCom, it.vUnCom,
         F(`J${rw}*K${rw}`), it.vDesc, it.vBC_ICMS, it.pICMS ? it.pICMS / 100 : 0, it.vICMS, it.IPI_CST,
         it.pIPI ? it.pIPI / 100 : 0, it.vIPI, it.PIS_CST, it.vPIS, it.COFINS_CST, it.vCOFINS].forEach((v, j) => wi.getCell(rw, j + 1).value = v);
        rw++;
      }
    }
    const i_last = rw - 1;
    wi.getCell(rw, 11).value = "TOTAL"; wi.getCell(rw, 11).font = { name: FONT, bold: true };
    for (const col of [12,13,14,16,19,21,23]) wi.getCell(rw, col).value = F(`SUM(${cl(col)}4:${cl(col)}${i_last})`);
    for (let c = 1; c <= ic.length; c++) { wi.getCell(rw, c).fill = fillSolid(LIGHT); wi.getCell(rw, c).font = { name: FONT, bold: true }; }
    for (let r2 = 4; r2 <= rw; r2++) {
      wi.getCell(r2, 10).numFmt = '#,##0'; wi.getCell(r2, 10).alignment = R;
      for (const col of [11,12,13,14,16,19,21,23]) { wi.getCell(r2, col).numFmt = CUR; wi.getCell(r2, col).alignment = R; }
      for (const col of [15,18]) { wi.getCell(r2, col).numFmt = '0.00%'; wi.getCell(r2, col).alignment = R; }
      for (const col of [1,2,3,5,6,7,8,9,17,20,22]) wi.getCell(r2, col).alignment = C;
      for (let c = 1; c <= ic.length; c++) { wi.getCell(r2, c).border = BORDER; if (r2 !== rw) wi.getCell(r2, c).font = { name: FONT, size: 9 }; }
    }
    setw(wi, [9,7,6,32,11,7,10,6,5,8,12,13,10,12,9,11,8,8,11,8,10,8,11]);

    return wb;
  }

  // estiliza linhas de dados (formatos de moeda em colsCur, centro em colsCtr) + total na última linha
  function estiloLinhas(ws, first, totalRow, colsCur, colsCtr, ncols) {
    for (let r = first; r <= totalRow; r++) {
      for (const col of colsCur) { ws.getCell(r, col).numFmt = CUR; ws.getCell(r, col).alignment = R; }
      for (const col of colsCtr) ws.getCell(r, col).alignment = C;
      for (let c = 1; c <= ncols; c++) { ws.getCell(r, c).border = BORDER; if (r !== totalRow) ws.getCell(r, c).font = { name: FONT, size: 9 }; }
    }
  }

  const cl = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const round2 = v => Math.round((v + Number.EPSILON) * 100) / 100;
  const fmt2 = v => Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const range = (a, b) => { const o = []; for (let i = a; i <= b; i++) o.push(i); return o; };

  async function gerarBlob(entradas, saidas, empresa, per, ini, fim) {
    const wb = gerar(entradas, saidas, empresa, per, ini, fim);
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  return { gerar, gerarBlob };
})();
