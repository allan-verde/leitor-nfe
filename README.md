# Leitor de NF-e — versão web (estática, 100% no navegador)

Lê XMLs de NF-e (modelo 55) e NFC-e (modelo 65), mostra os dados básicos e
exporta o relatório Excel padronizado de 7 abas — **tudo dentro do navegador
do usuário**. Nenhum dado é enviado para servidor algum: a página não faz
nenhuma chamada de rede (as bibliotecas ExcelJS e JSZip estão embutidas em
`vendor/`). Por isso é segura para dados fiscais sigilosos e pode ser hospedada
no GitHub Pages.

## Estrutura (esta pasta é o site inteiro)

```
index.html          página
css/style.css        estilo (identidade NAVY/LIGHT/Arial do relatório)
js/parse.js          leitura do XML + regras (porte de nfe_core.py)
js/excel.js          geração das 7 abas com ExcelJS (porte de gerar())
js/app.js            upload/drag&drop, leitura de .zip, preview e download
vendor/              ExcelJS e JSZip embutidos (sem CDN)
.nojekyll            evita processamento Jekyll no GitHub Pages
```

## Testar localmente

Abrir direto pelo `file://` funciona, mas servir por HTTP é mais fiel:

```bash
# dentro desta pasta (webapp/)
python -m http.server 8000
# abrir http://127.0.0.1:8000
```

## Publicar no GitHub Pages

1. Crie um repositório no GitHub (ex.: `leitor-nfe`).
2. Coloque **o conteúdo desta pasta `webapp/`** na **raiz** do repositório
   (o `index.html` deve ficar na raiz), faça commit e push:
   ```bash
   git init
   git add .
   git commit -m "Leitor de NF-e (web estático)"
   git branch -M main
   git remote add origin https://github.com/SEU_USUARIO/leitor-nfe.git
   git push -u origin main
   ```
3. No GitHub: **Settings → Pages → Build and deployment → Source: Deploy from a
   branch → Branch: `main` / `/ (root)`** → Save.
4. Aguarde ~1 min. O site fica em `https://SEU_USUARIO.github.io/leitor-nfe/`.

> Repositório público (plano grátis): só o **código** fica visível — não há
> segredo nem dado fiscal nele. Os XMLs do usuário nunca saem da máquina dele.

## Compatibilidade

Navegadores modernos (Chrome, Edge, Firefox). Usa File API, DOMParser, Blob e
`URL.createObjectURL` — todos amplamente suportados.
