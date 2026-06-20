/* parse.js — leitura de NF-e/NFC-e e regras de negócio, 100% no navegador.
   Porte fiel de nfe_core.py (parse, classificação, detecção de empresa,
   separação entrada/saída, período e preview). Nenhum dado sai da máquina. */

const NFE = (() => {
  "use strict";

  // ---------- tabelas de domínio ----------
  const PAG = {'01':'Dinheiro','02':'Cheque','03':'Cartão crédito','04':'Cartão débito','05':'Crédito loja',
    '10':'Vale alimentação','11':'Vale refeição','12':'Vale presente','13':'Vale combustível',
    '15':'Boleto','16':'Depósito','17':'PIX','18':'Transferência','19':'Carteira digital',
    '90':'Sem pagamento','99':'Outros'};
  const MODF = {'0':'Conta emitente (CIF)','1':'Conta destinatário (FOB)','2':'Conta terceiros',
    '3':'Próprio remetente','4':'Próprio destinatário','9':'Sem frete'};
  const CRT = {'1':'Simples Nacional','2':'Simples Nacional (excesso sublimite)','3':'Regime Normal','4':'MEI'};
  const MESES = ['','Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

  const SUF_ENT = {'101':'Compra para comercialização','102':'Compra para comercialização',
    '103':'Compra para comercialização','113':'Compra para comercialização','117':'Compra para comercialização',
    '551':'Ativo imobilizado','552':'Ativo imobilizado','556':'Uso e consumo','557':'Uso e consumo',
    '407':'Uso e consumo','910':'Bonificação / brinde','911':'Amostra grátis','912':'Demonstração',
    '913':'Retorno de demonstração','915':'Remessa p/ conserto','916':'Retorno de conserto',
    '949':'Outras entradas'};
  const SUF_SAI = {'101':'Venda (produção própria)','102':'Venda a consumidor / revenda',
    '103':'Venda a consumidor / revenda','405':'Venda (subst. tributária)',
    '910':'Bonificação / brinde','911':'Amostra grátis','912':'Remessa demonstração',
    '913':'Retorno de demonstração','915':'Remessa p/ conserto','916':'Retorno de conserto',
    '949':'Outras saídas'};
  const ORDEM_CAT = ['Compra para comercialização','Compra p/ industrialização','Venda a consumidor / revenda',
    'Venda (produção própria)','Bonificação / brinde','Amostra grátis','Uso e consumo',
    'Ativo imobilizado','Demonstração','Remessa demonstração','Retorno de demonstração',
    'Venda (subst. tributária)','Outras entradas','Outras saídas'];
  const ORDEM_CLASS = ['Compra para comercialização','Compra para industrialização','Bonificação para revenda',
    'Uso e consumo / brinde','Uso e consumo','Ativo imobilizado',
    'Demonstração (não é compra)','Outras'];

  function fmtDoc(s) {
    if (!s) return '';
    if (s.length === 14) return `${s.slice(0,2)}.${s.slice(2,5)}.${s.slice(5,8)}/${s.slice(8,12)}-${s.slice(12)}`;
    if (s.length === 11) return `${s.slice(0,3)}.${s.slice(3,6)}.${s.slice(6,9)}-${s.slice(9)}`;
    return s;
  }

  function categoriaCfop(cfop, lado) {
    const suf = (cfop || '').slice(-3);
    const tbl = lado === 'entrada' ? SUF_ENT : SUF_SAI;
    return tbl[suf] || `CFOP ${cfop} — outras`;
  }

  // ---------- helpers de leitura DOM (ignora namespace via localName) ----------
  function localFind(node, name) {
    if (!node) return null;
    for (const c of node.children) if (c.localName === name) return c;
    return null;
  }
  function localFindAll(node, name) {
    const out = [];
    if (!node) return out;
    for (const c of node.children) if (c.localName === name) out.push(c);
    return out;
  }
  // caminho tipo "a/b/c" partindo de node
  function pathFind(node, path) {
    let cur = node;
    for (const part of path.split('/')) {
      cur = localFind(cur, part);
      if (!cur) return null;
    }
    return cur;
  }
  function txt(node, path) {
    const n = path ? pathFind(node, path) : node;
    return n ? n.textContent : null;
  }
  // busca em profundidade pelo primeiro elemento com dado localName
  function deepFind(root, name) {
    const all = root.getElementsByTagName('*');
    for (const el of all) if (el.localName === name) return el;
    return null;
  }
  const num = v => parseFloat(v || 0) || 0;

  // ---------- parse de um XML (string) ----------
  function parse(xmlString) {
    const doc = new DOMParser().parseFromString(xmlString, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) return null;
    const inf = deepFind(doc, 'infNFe');
    if (!inf) return null;

    const ide = localFind(inf, 'ide');
    const emit = localFind(inf, 'emit');
    const dest = localFind(inf, 'dest');
    const tot = pathFind(inf, 'total/ICMSTot');
    const protInf = (() => { const p = deepFind(doc, 'protNFe'); return p ? localFind(p, 'infProt') : null; })();
    const transp = localFind(inf, 'transp');

    const rec = {
      chave: (inf.getAttribute('Id') || '').replace('NFe', ''),
      protocolo: txt(protInf, 'nProt'),
      cStat: txt(protInf, 'cStat'),
      mod: txt(ide, 'mod'),
      nNF: txt(ide, 'nNF'), serie: txt(ide, 'serie'), natOp: txt(ide, 'natOp'),
      dhEmi: (txt(ide, 'dhEmi') || txt(ide, 'dEmi') || ''),
      tpNF: txt(ide, 'tpNF'),
      emit_cnpj: (txt(emit, 'CNPJ') || txt(emit, 'CPF')),
      emit_nome: txt(emit, 'xNome'),
      emit_uf: txt(emit, 'enderEmit/UF'),
      emit_ie: txt(emit, 'IE'), emit_crt: txt(emit, 'CRT'),
      dest_doc: (txt(dest, 'CNPJ') || txt(dest, 'CPF')),
      dest_nome: txt(dest, 'xNome'),
      dest_uf: txt(dest, 'enderDest/UF'),
      dest_mun: txt(dest, 'enderDest/xMun'),
    };
    for (const tag of ['vProd','vNF','vDesc','vFrete','vSeg','vOutro','vBC','vICMS','vIPI','vPIS','vCOFINS','vST'])
      rec[tag] = num(txt(tot, tag));

    rec.pag = [];
    const pagNode = localFind(inf, 'pag');
    if (pagNode) for (const dp of localFindAll(pagNode, 'detPag')) {
      const t = txt(dp, 'tPag');
      rec.pag.push([PAG[t] || t, num(txt(dp, 'vPag'))]);
    }
    rec.modFrete = transp ? (MODF[txt(transp, 'modFrete')] || txt(transp, 'modFrete')) : null;
    rec.transp_nome = transp ? txt(transp, 'transporta/xNome') : null;

    const itens = [];
    for (const det of localFindAll(inf, 'det')) {
      const prod = localFind(det, 'prod');
      const imp = localFind(det, 'imposto');
      const icms = imp ? localFind(imp, 'ICMS') : null;
      const icn = (icms && icms.children.length) ? icms.children[0] : null;
      const ipi = imp ? localFind(imp, 'IPI') : null;
      const ipt = ipi ? localFind(ipi, 'IPITrib') : null;
      const pis = imp ? localFind(imp, 'PIS') : null;
      const pn = (pis && pis.children.length) ? pis.children[0] : null;
      const cof = imp ? localFind(imp, 'COFINS') : null;
      const cn = (cof && cof.children.length) ? cof.children[0] : null;
      itens.push({
        nItem: det.getAttribute('nItem'),
        cProd: txt(prod, 'cProd'), xProd: txt(prod, 'xProd'),
        NCM: txt(prod, 'NCM'), CFOP: txt(prod, 'CFOP'), uCom: txt(prod, 'uCom'),
        qCom: num(txt(prod, 'qCom')), vUnCom: num(txt(prod, 'vUnCom')),
        vProd: num(txt(prod, 'vProd')), vDesc: num(txt(prod, 'vDesc')),
        orig: icn ? txt(icn, 'orig') : null,
        CST_CSOSN: icn ? (txt(icn, 'CST') || txt(icn, 'CSOSN')) : null,
        vBC_ICMS: icn ? num(txt(icn, 'vBC')) : 0,
        pICMS: icn ? num(txt(icn, 'pICMS')) : 0,
        vICMS: icn ? num(txt(icn, 'vICMS')) : 0,
        IPI_CST: ipt ? txt(ipt, 'CST') : null,
        pIPI: ipt ? num(txt(ipt, 'pIPI')) : 0,
        vIPI: ipt ? num(txt(ipt, 'vIPI')) : 0,
        PIS_CST: pn ? txt(pn, 'CST') : null,
        vPIS: pn ? num(txt(pn, 'vPIS')) : 0,
        COFINS_CST: cn ? txt(cn, 'CST') : null,
        vCOFINS: cn ? num(txt(cn, 'vCOFINS')) : 0,
      });
    }
    rec.itens = itens;
    rec.cfops = [...new Set(itens.map(i => i.CFOP))].sort();
    rec.cfop_dom = itens.length ? itens.reduce((a, b) => (b.vProd > a.vProd ? b : a)).CFOP : '';
    return rec;
  }

  // ---------- classificação por item ----------
  function classificaItem(it) {
    const cf = it.CFOP || ''; const suf = cf.slice(-3); const nome = (it.xProd || '').toUpperCase();
    if (suf === '912') return 'Demonstração (não é compra)';
    if (suf === '910' || suf === '911') {
      if (['NOT FOR SALE','NAO PARA VENDA','AMOSTRA','BRINDE'].some(x => nome.includes(x)))
        return 'Uso e consumo / brinde';
      return 'Bonificação para revenda';
    }
    if (suf === '551' || suf === '552') return 'Ativo imobilizado';
    if (['556','557','407'].includes(suf)) return 'Uso e consumo';
    if (['101','102','103','113','117'].includes(suf)) return 'Compra para comercialização';
    return 'Outras';
  }

  // ---------- detecção da empresa num lote misto ----------
  function detectarEmpresaMisto(notas) {
    const cnt = new Map();
    const bump = d => { if (d) cnt.set(d, (cnt.get(d) || 0) + 1); };
    for (const r of notas) { bump(r.emit_cnpj); bump(r.dest_doc); }
    if (!cnt.size) return { cnpj:'', nome:'', uf:null, ie:null, crt:null };
    let doc = null, max = -1;
    for (const [k, v] of cnt) if (v > max) { max = v; doc = k; }
    const ident = { cnpj: doc, nome:'', uf:null, ie:null, crt:null };
    const comoEmit = notas.find(r => r.emit_cnpj === doc);
    if (comoEmit) Object.assign(ident, { nome: comoEmit.emit_nome, uf: comoEmit.emit_uf, ie: comoEmit.emit_ie, crt: comoEmit.emit_crt });
    else {
      const comoDest = notas.find(r => r.dest_doc === doc);
      if (comoDest) Object.assign(ident, { nome: comoDest.dest_nome, uf: comoDest.dest_uf });
    }
    return ident;
  }

  // ---------- separação de direção ----------
  function separarDirecao(notas, empresaCnpj) {
    const entradas = [], saidas = [], indefinidas = [];
    for (const r of notas) {
      if (empresaCnpj && r.emit_cnpj === empresaCnpj) saidas.push(r);
      else if (empresaCnpj && r.dest_doc === empresaCnpj) entradas.push(r);
      else if (r.tpNF === '0') entradas.push(r);
      else if (r.tpNF === '1') saidas.push(r);
      else indefinidas.push(r);
    }
    return { entradas, saidas, indefinidas };
  }

  function periodo(entradas, saidas) {
    const datas = [...entradas, ...saidas].filter(r => r.dhEmi).map(r => r.dhEmi.slice(0, 10));
    if (!datas.length) return { per:'', ini:'', fim:'' };
    datas.sort();
    const ini = datas[0], fim = datas[datas.length - 1];
    const cnt = new Map();
    for (const d of datas) { const k = d.slice(0, 7); cnt.set(k, (cnt.get(k) || 0) + 1); }
    let topMes = null, max = -1;
    for (const [k, v] of cnt) if (v > max) { max = v; topMes = k; }
    const [ano, mes] = topMes.split('-');
    return { per: `${MESES[parseInt(mes, 10)]}/${ano}`, ini, fim };
  }

  // ---------- preview ----------
  function resumoCategorias(notas, lado) {
    const agg = new Map(); // cat -> [qtd, vProd, vNF]
    for (const r of notas) {
      const cat = categoriaCfop(r.cfop_dom, lado);
      const a = agg.get(cat) || [0, 0, 0];
      a[0]++; a[1] += r.vProd; a[2] += r.vNF; agg.set(cat, a);
    }
    const presentes = ORDEM_CAT.filter(c => agg.has(c)).concat([...agg.keys()].filter(c => !ORDEM_CAT.includes(c)));
    const linhas = presentes.map(c => ({ categoria: c, qtd: agg.get(c)[0], vProd: round2(agg.get(c)[1]), vNF: round2(agg.get(c)[2]) }));
    let q = 0, vp = 0, vn = 0;
    for (const a of agg.values()) { q += a[0]; vp += a[1]; vn += a[2]; }
    return { linhas, total: { qtd: q, vProd: round2(vp), vNF: round2(vn) } };
  }
  function linhasNotas(notas, lado) {
    return [...notas].sort((a, b) => (a.dhEmi || '').localeCompare(b.dhEmi || '')).map(r => ({
      nNF: r.nNF, serie: r.serie, data: (r.dhEmi || '').slice(0, 10), chave: r.chave, mod: r.mod,
      contraparte: (lado === 'entrada' ? r.emit_nome : r.dest_nome) || '',
      uf: (lado === 'entrada' ? r.emit_uf : r.dest_uf) || '',
      cfop: r.cfops.join(','), categoria: categoriaCfop(r.cfop_dom, lado),
      vProd: round2(r.vProd), vNF: round2(r.vNF), vICMS: round2(r.vICMS), vIPI: round2(r.vIPI),
    }));
  }
  function resumoClassificacao(entradas) {
    const agg = new Map();
    for (const r of entradas) for (const it of r.itens) {
      const k = classificaItem(it); const a = agg.get(k) || [0, 0]; a[0]++; a[1] += it.vProd; agg.set(k, a);
    }
    const presentes = ORDEM_CLASS.filter(k => agg.has(k)).concat([...agg.keys()].filter(k => !ORDEM_CLASS.includes(k)));
    let total = 0; for (const a of agg.values()) total += a[1];
    return presentes.map(k => ({ classificacao: k, qtd: agg.get(k)[0], vProd: round2(agg.get(k)[1]), pct: total ? agg.get(k)[1] / total : 0 }));
  }
  function contaMod(notas, m) { return notas.filter(r => r.mod === m).length; }

  function montarPreview(entradas, saidas, empresa, p, indefinidas) {
    indefinidas = indefinidas || [];
    const todas = entradas.concat(saidas);
    return {
      identificacao: {
        empresa: empresa.nome || '(não identificada)',
        cnpj: fmtDoc(empresa.cnpj), ie: empresa.ie || '', uf: empresa.uf || '',
        regime: empresa.crt ? (CRT[empresa.crt] || '') : '',
        periodo: p.per, data_ini: p.ini, data_fim: p.fim,
        n_entradas: entradas.length, n_saidas: saidas.length, n_indefinidas: indefinidas.length,
        n_mod55: contaMod(todas, '55'), n_mod65: contaMod(todas, '65'),
      },
      resumo_entradas: resumoCategorias(entradas, 'entrada'),
      resumo_saidas: resumoCategorias(saidas, 'saida'),
      classificacao_entradas: resumoClassificacao(entradas),
      entradas: linhasNotas(entradas, 'entrada'),
      saidas: linhasNotas(saidas, 'saida'),
      indefinidas: indefinidas.map(r => ({ nNF: r.nNF, chave: r.chave, emit: r.emit_nome, dest: r.dest_nome, tpNF: r.tpNF })),
    };
  }

  const round2 = v => Math.round((v + Number.EPSILON) * 100) / 100;

  // monta o resultado final a partir das notas já parseadas
  function _finalizar(notas, totalArquivos, erros) {
    if (!notas.length) return { ok: false, erro: 'Os XMLs enviados não puderam ser lidos como NF-e/NFC-e.', lidos: totalArquivos, erros };
    const empresa = detectarEmpresaMisto(notas);
    const { entradas, saidas, indefinidas } = separarDirecao(notas, empresa.cnpj);
    const p = periodo(entradas, saidas);
    const preview = montarPreview(entradas, saidas, empresa, p, indefinidas);
    preview.arquivos_lidos = totalArquivos; preview.arquivos_com_erro = erros;
    return { ok: true, empresa, entradas, saidas, indefinidas, periodo: p, preview };
  }

  // orquestração síncrona (usada nos testes em Node)
  function processar(arquivos) {
    const notas = []; let erros = 0;
    for (const a of arquivos) {
      try { const r = parse(a.conteudo); if (r) notas.push(r); else erros++; }
      catch (e) { erros++; }
    }
    return _finalizar(notas, arquivos.length, erros);
  }

  // orquestração assíncrona em blocos: não trava a UI mesmo com muitos XMLs.
  // onProgress(processados, total) é chamado a cada bloco.
  async function processarAsync(arquivos, onProgress) {
    const notas = []; let erros = 0;
    const CHUNK = 150;
    for (let i = 0; i < arquivos.length; i++) {
      try { const r = parse(arquivos[i].conteudo); if (r) notas.push(r); else erros++; }
      catch (e) { erros++; }
      if ((i + 1) % CHUNK === 0) {
        if (onProgress) onProgress(i + 1, arquivos.length);
        await new Promise(res => setTimeout(res, 0)); // cede o controle p/ o navegador repintar
      }
    }
    if (onProgress) onProgress(arquivos.length, arquivos.length);
    return _finalizar(notas, arquivos.length, erros);
  }

  return { parse, processar, processarAsync, classificaItem, categoriaCfop, detectarEmpresaMisto,
    separarDirecao, periodo, montarPreview, fmtDoc,
    PAG, MODF, CRT, MESES, SUF_ENT, SUF_SAI, ORDEM_CAT, ORDEM_CLASS };
})();
