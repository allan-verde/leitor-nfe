/* worker.js — gera o .xlsx numa thread separada, para a UI não travar com
   volume alto de notas. Recebe os dados já parseados (objetos simples) e
   devolve o ArrayBuffer do arquivo. Não usa DOM (o parse fica na página). */
importScripts('../vendor/exceljs.min.js'); // define ExcelJS
importScripts('parse.js');                  // define NFE (helpers de domínio; parse() não é usado aqui)
importScripts('excel.js');                  // define NFEExcel (usa ExcelJS + NFE)

self.onmessage = async (e) => {
  try {
    const { entradas, saidas, empresa, periodo } = e.data;
    const wb = NFEExcel.gerar(entradas, saidas, empresa, periodo.per, periodo.ini, periodo.fim);
    const out = await wb.xlsx.writeBuffer();
    // writeBuffer pode devolver Uint8Array/Buffer; transferimos o ArrayBuffer subjacente
    const ab = (out && out.buffer instanceof ArrayBuffer)
      ? out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength)
      : out;
    self.postMessage({ ok: true, buf: ab }, [ab]); // transfere (sem cópia)
  } catch (err) {
    self.postMessage({ ok: false, erro: String((err && err.message) || err) });
  }
};
