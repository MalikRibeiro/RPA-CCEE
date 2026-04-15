// ==UserScript==
// @name         CCEE Extrator de Relatórios (100% automatico)
// @namespace    http://tampermonkey.net/
// @version      1.0.5
// @description  Extração automática CCEE com correção de nomes, abas únicas e pausa imediata
// @author       Malik (Ajustado por Manus)
// @match        https://operacao.ccee.org.br/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_download
// @run-at       document-start
// ==/UserScript==

(() => {
  'use strict';

// ════════════════════════════════════════════════════════════════
  //  BLOCO 1 — PDF AUTO-DOWNLOAD (roda na aba saw.dll)
  // ════════════════════════════════════════════════════════════════
  if (
    window.location.href.includes('saw.dll') &&
    window.location.href.includes('downloadExportedFile')
  ) {
    if (window.__cceeDownloader) return;
    window.__cceeDownloader = true;

    // TRUQUE NINJA: Para imediatamente o carregamento nativo da aba!
    // Isso impede que o navegador "queime" o token de uso único da CCEE.
    window.stop();

    const params = new URLSearchParams(window.location.search);
    const itemName = decodeURIComponent(params.get('ItemName') || 'relatorio');

    let empresa = '';
    let cfg = {};
    try { empresa = GM_getValue('ccee_empresa_atual', ''); } catch {}
    try { cfg = JSON.parse(GM_getValue('ccee_config', '{}')); } catch {}

    const mes       = cfg.mes       || 'xxx';
    const ano       = cfg.ano       || 'xx';
    const relatorio = cfg.relatorio || itemName.split(' ')[0];

    // Formata nome
    const formatarNome = (s) => s.replace(/[\s\\/:*?"<>|]/g, '_').replace(/_+/g, '_');
    const nomeArquivo = empresa
      ? `${formatarNome(empresa)}_${formatarNome(relatorio)}_${mes}_${ano}.pdf`
      : `${formatarNome(itemName)}_${mes}_${ano}.pdf`;

    // Baixa o arquivo silenciosamente
    GM_download({
      url: window.location.href,
      name: nomeArquivo,
      saveAs: false,
      onload: () => {
        GM_setValue('ccee_download_ok', nomeArquivo);
        GM_setValue('ccee_download_ts', String(Date.now()));
        setTimeout(() => window.close(), 1500); // Fecha se der certo
      },
      onerror: (err) => {
        const motivo = err.error || err.details || 'Desconhecido';
        console.error('[CCEE AutoDL] Erro GM_download:', err);
        GM_setValue('ccee_download_ok', `__ERRO__:${motivo}`);
        GM_setValue('ccee_download_ts', String(Date.now()));
        setTimeout(() => window.close(), 1500);
      }
    });
    return;
  }

  if (window !== window.top) return;
  if (window.__cceeExtrator) return;
  window.__cceeExtrator = true;

  // ════════════════════════════════════════════════════════════════
  //  HELPERS UTILITÁRIOS
  // ════════════════════════════════════════════════════════════════
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  // Função de espera que respeita PAUSA e STOPPED instantaneamente
  async function waitFor(fn, timeout = 10000, interval = 500) {
    const t0 = Date.now();
    while (Date.now() - t0 < timeout) {
      if (estado === 'stopped') return null;
      while (estado === 'paused') {
        if (estado === 'stopped') return null;
        await sleep(200);
      }
      try {
        const r = await fn();
        if (r) return r;
      } catch {}
      await sleep(interval);
    }
    return null;
  }

  function xpathFirst(xpath, doc) {
    try {
      const r = (doc || document).evaluate(
        xpath, doc || document, null,
        XPathResult.FIRST_ORDERED_NODE_TYPE, null
      );
      return r.singleNodeValue || null;
    } catch { return null; }
  }

  function findInFrames(searchFn, doc, depth) {
    doc   = doc   || document;
    depth = depth || 0;
    if (depth > 5) return null;
    const result = searchFn(doc);
    if (result) return result;
    try {
      const iframes = doc.querySelectorAll('iframe');
      for (const iframe of iframes) {
        try {
          const child = iframe.contentDocument || iframe.contentWindow?.document;
          if (!child) continue;
          const found = findInFrames(searchFn, child, depth + 1);
          if (found) return found;
        } catch {}
      }
    } catch {}
    return null;
  }

  function findXPath(xpath) {
    return findInFrames(doc => xpathFirst(xpath, doc));
  }

  let lastClickTime = 0;
  function clickOracleBI(el) {
    if (!el) return false;
    
    // Removemos a trava de 2.5s daqui porque ela estava bloqueando a seleção da empresa.
    // O duplo clique no PDF já está protegido pelo window.__pdfLock lá embaixo!
    try {
      el.scrollIntoView({ block: 'center', inline: 'nearest' });
      // Restaurando os eventos nativos que o Oracle BI EXIGE para funcionar
      el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true, cancelable: true }));
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
      el.click(); 
      return true;
    } catch (e) {
      return false;
    }
  }

  function validarConteudo() {
    const xpaths = [
      "//table[contains(@class, 'TableView')]",
      "//div[contains(@class, 'ViewContainer')]",
      "//*[contains(@class, 'DashboardContent')]",
      "//td[contains(@class, 'TableHeaderCell')]",
      "//*[contains(@id, 'saw_')]",
      "//div[contains(@class, 'ResultsContainer')]"
    ];
    const erroXpaths = [
      "//div[contains(text(), 'Nenhum resultado')]",
      "//div[contains(text(), 'Sem dados')]",
      "//*[contains(@class, 'ErrorMessage')]"
    ];
    for (const xpath of erroXpaths) {
      if (findXPath(xpath)) return false;
    }
    for (const xpath of xpaths) {
      if (findXPath(xpath)) return true;
    }
    return false;
  }

  function dispararSubmenuImprimir(el) {
    if (!el) return;
    const ev = new MouseEvent('mouseover', { bubbles: true, cancelable: true });
    try { el.dispatchEvent(ev); } catch {}
    try {
      if (typeof el.onmouseover === 'function') {
        el.onmouseover.call(el, ev);
      }
    } catch {}
    try {
      const attr = el.getAttribute('onmouseover');
      if (attr && attr.includes('saw.dashboard.DisplayLayouts')) {
        new Function('event', attr).call(el, ev);
      }
    } catch {}
  }

  // ════════════════════════════════════════════════════════════════
  //  ESTADO DA AUTOMAÇÃO
  // ════════════════════════════════════════════════════════════════
  let estado       = 'idle';
  let filaPendente = [];
  let indiceAtual  = 0;
  let resultados   = { sucesso: 0, falha: 0 };

  // ════════════════════════════════════════════════════════════════
  //  ENGINE — processa uma empresa por vez
  // ════════════════════════════════════════════════════════════════
  async function processarEmpresa(empresa) {
    if (estado === 'stopped') return false;

    uiLog(`▶ <b>${empresa}</b>`);
    uiStatus(`Processando: ${empresa}`);

    GM_setValue('ccee_download_ok', '');
    GM_setValue('ccee_empresa_atual', empresa);

    // ── PASSO 1: Selecionar empresa ──────────────────
    uiLog(`  [1/7] Selecionando empresa...`);
    const xpEmpresa = `//div[contains(@class, 'listBoxPanelOptionBasic') and @title='${empresa}']`;
    const elEmpresa = await waitFor(() => findXPath(xpEmpresa), 8000, 400);
    if (!elEmpresa || estado === 'stopped') return false;
    clickOracleBI(elEmpresa);
    await sleep(2000);

    // ── PASSO 2: Clicar em Aplicar ──────────────────────────────
    uiLog(`  [2/7] Clicando em Aplicar...`);
    const elAplicar = await waitFor(() => {
      return document.getElementById('gobtn') ||
             findXPath("//input[@value='Aplicar']") ||
             findXPath("//button[normalize-space(text())='Aplicar']");
    }, 6000, 400);
    if (!elAplicar || estado === 'stopped') return false;
    clickOracleBI(elAplicar);

    // ── PASSO 3: Aguardar carregamento ───────────────────
    uiLog(`  [3/7] Aguardando dados...`);
    await sleep(6000);
    const conteudoOK = await waitFor(validarConteudo, 35000, 1000);
    if (!conteudoOK || estado === 'stopped') {
      uiLog(`  ✗ Conteúdo não carregou`, 'err');
      return false;
    }
    uiLog(`  [3/7] Conteúdo validado ✓`);
    await sleep(2500);

    // ── PASSO 4: Abrir engrenagem ─────────
    uiLog(`  [4/7] Abrindo opções...`);
    const xpGear = "//img[contains(@id, 'dashboardpageoptions') or contains(@src, 'popupmenu')]";
    const elGear = await waitFor(() => findXPath(xpGear), 8000, 500);
    if (!elGear || estado === 'stopped') return false;
    clickOracleBI(elGear);
    await sleep(3000);

    // ── PASSO 5: Localizar Imprimir ──────────────────────
    uiLog(`  [5/7] Localizando Imprimir...`);
    const elImprimir = await waitFor(() => findXPath("//*[@id='idPagePrint']"), 6000, 400);
    if (!elImprimir || estado === 'stopped') return false;

    // ── PASSO 6: Disparar PDF ──────────────────────
    uiLog(`  [6/7] Disparando PDF...`);
    dispararSubmenuImprimir(elImprimir);
    await sleep(3000);

    const xpathsPDF = [
      "//td[contains(@class, 'MenuItemTextCell') and contains(text(), 'PDF')]",
      "//td[contains(text(), 'Página Atual como PDF')]",
      "//a[contains(@class, 'MenuItem') and contains(., 'PDF')]",
      "//td[@class='MenuItemTextCell' and normalize-space(text())='PDF']"
    ];

    let elPDF = null;
    for (const xpath of xpathsPDF) {
      elPDF = await waitFor(() => {
        const el = findXPath(xpath);
        if (!el) return null;
        const rect = el.getBoundingClientRect();
        return (rect.width > 0 && rect.height > 0) ? el : null;
      }, 3500, 500);
      if (elPDF || estado === 'stopped') break;
    }

    if (!elPDF && estado !== 'stopped') {
      uiLog(`  ⚠ Tentando clique direto no Imprimir...`, 'warn');
      clickOracleBI(elImprimir);
      await sleep(3000);
      for (const xpath of xpathsPDF) {
        elPDF = await waitFor(() => findXPath(xpath), 3000, 500);
        if (elPDF || estado === 'stopped') break;
      }
    }

    if (!elPDF || estado === 'stopped') {
      uiLog(`  ✗ Opção PDF não encontrada`, 'err');
      return false;
    }

    // TRAVA DE DISPARO ÚNICO REFORÇADA
    if (window.__pdfLock) return false;
    window.__pdfLock = true;

    // TRAVA: Clica apenas UMA vez e aguarda o sinal de download
    clickOracleBI(elPDF);
    uiLog(`  [6/7] PDF clicado ✓`);

    // ── PASSO 7: Aguardar download ─────────────────
    uiLog(`  [7/7] Aguardando download...`);
    const tsAntes = Date.now();

    const resultado = await waitFor(() => {
      const ts = parseInt(GM_getValue('ccee_download_ts', '0') || '0');
      if (ts > tsAntes) {
        return GM_getValue('ccee_download_ok', '') || '__SEM_NOME__';
      }
      return null;
    }, 50000, 1000);

    // CORREÇÃO: LIBERA A TRAVA PARA A PRÓXIMA EMPRESA
    window.__pdfLock = false;

    if (!resultado || estado === 'stopped') {
      uiLog(`  ⚠ Timeout ou parado.`, 'warn');
      return false; // Mudei para false para contabilizar como falha
    }
    
    // Verifica se a string que voltou começa com __ERRO__
    if (resultado.startsWith('__ERRO__')) {
      // Pega o texto que vem depois dos dois pontos
      const motivoErro = resultado.split(':')[1] || 'Token bloqueado';
      uiLog(`  ✗ Falha no Download: <b>${motivoErro}</b>`, 'err');
      return false; // Mudei para false para contabilizar como falha
    }

    uiLog(`  ✓ Salvo: <i style="color:#888">${resultado}</i>`, 'ok');
    await sleep(3000);
    return true;
  }

  async function rodarAutomacao() {
    resultados = { sucesso: 0, falha: 0 };
    const empresas = [...filaPendente];

    for (let i = indiceAtual; i < empresas.length; i++) {
      // Verifica pausa/parada antes de cada empresa
      while (estado === 'paused') {
        uiStatus('Pausado — aguardando retomada...');
        await sleep(500);
      }
      if (estado === 'stopped') break;

      indiceAtual = i;
      uiProgress(i, empresas.length);

      const empresa = empresas[i];
      let ok = false;
      try {
        ok = await processarEmpresa(empresa);
      } catch (e) {
        uiLog(`  ✗ Erro: ${e.message}`, 'err');
      }

      if (estado === 'stopped') break;

      if (ok) resultados.sucesso++;
      else    resultados.falha++;

      uiProgress(i + 1, empresas.length);
      await sleep(2500);
    }

    const finalStatus = estado === 'stopped' ? 'Interrompido' : 'Concluído';
    estado = 'idle';
    setBtnState('idle');
    uiLog(`<br>═══ EXTRAÇÃO FINALIZADA (${finalStatus}) ═══`);
    uiLog(`✓ Sucesso: <b>${resultados.sucesso}</b>  &nbsp; ✗ Falhas: <b>${resultados.falha}</b>`, 'ok');
    uiStatus(finalStatus);
  }

  // ════════════════════════════════════════════════════════════════
  //  DADOS E PERSISTÊNCIA
  // ════════════════════════════════════════════════════════════════
  const EMPRESAS_DEFAULT = [
  "ADESI CL 514", "AMERICAS CL 514", "AOI YAMA CL 514", "APOLO TUBULARS", "APUCARANINHA",
  "ARBHORES COMPENSADOS CL 514", "AURORA FASGO CL", "AURORA MATRIZ", "AURORA XAXIM", "BAIA VERDE",
  "C VALE COOP AGROINDUSTRIAL", "C VALE ENERGIA", "CERMISSOES", "CERT", "CERTAJA",
  "CERTHIL DISTRIBUICAO", "CGH DONA MARIA PIANA", "CGH FAXINAL DOS GUEDES", "CGH MARUMBI ESP",
  "CHOPIM 1 ESP", "CIFA CL 514", "CONTESTADO", "CONTINENTAL COM", "COOPERALFA", "COOPERLUZ",
  "COOPERLUZ DIST", "CORONEL ARAUJO", "CRERAL DIST", "CVALE", "DETOFOL", "ELECTRA ENERGIA DIGITAL",
  "ELECTRA ENERGY", "EMBAUBA", "GUARICANA I5", "ISOELECTRIC L", "ITAGUACU", "MARCEGAGLIA",
  "MELISSA ESP PI", "MET CARTEC CL 514", "PALMAS", "PASSO FERRAZ", "PCH - ZECA GOLIN", "PCH BURITI",
  "PCH CAMBARA", "PCH FERRADURA", "PCH PESQUEIRO", "PCH PONTE BRANCA", "PCH QUEBRA DENTES",
  "PCH RIO INDIOS", "PCH SALTO DO GUASSUPI", "PCH TAMBAU", "PEQUI I5", "PITANGUI ESP PI",
  "PLASTBEL", "PLASTICOS PR", "RINCAO DOS ALBINOS", "RINCAO SAO MIGUEL", "SALTO DO TIMBO",
  "SALTO DO VAU I5", "SAO JORGE ESPECIAL", "SJHOTEIS", "SUCUPIRA", "THERMOPLAST CL 514",
  "TIMBER CL 514", "USINAS DO PRATA", "UTE GOIANESIA", "VIDROLAR CL 514", "VITORINO",
  "ZAELI CL 514", "CHAMINE", "CAVERNOSO I", "CAVERNOSO II", "BAIA VERDE", "PLASTICOS PR"
];

  const MESES = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];

  function loadConfig() {
    try { return JSON.parse(GM_getValue('ccee_config', '{}')); } catch { return {}; }
  }
  function saveConfig(c) {
    GM_setValue('ccee_config', JSON.stringify(c));
  }
  function loadEmpresas() {
    const s = GM_getValue('ccee_empresas', '');
    if (!s) return [...EMPRESAS_DEFAULT];
    try { return JSON.parse(s); } catch { return [...EMPRESAS_DEFAULT]; }
  }
  function saveEmpresas(l) {
    GM_setValue('ccee_empresas', JSON.stringify(l));
  }

  // ════════════════════════════════════════════════════════════════
  //  UI — HTML + CSS
  // ════════════════════════════════════════════════════════════════
  const CSS = `
  <style id="ccee_style">
  #ccee_wrap * { box-sizing:border-box; font-family:'JetBrains Mono','Fira Mono',monospace; }
  #ccee_toggle {
    position:fixed; bottom:16px; right:16px; width:48px; height:48px;
    border-radius:50%; background:#08080f; border:2px solid #00c896;
    color:#00c896; font-size:20px; cursor:pointer; z-index:2147483646;
    display:flex; align-items:center; justify-content:center;
    box-shadow:0 0 22px #00c89650; user-select:none; transition:transform .15s;
  }
  #ccee_toggle:hover { transform:scale(1.12); }
  #ccee_badge {
    position:absolute; top:-5px; right:-5px; background:#e03060;
    color:#fff; font-size:9px; border-radius:99px; padding:2px 5px;
    font-family:monospace; pointer-events:none; min-width:18px; text-align:center;
  }
  #ccee_panel {
    position:fixed; bottom:74px; right:16px; width:420px; max-width:96vw;
    height:500px; background:#07070e; border:1.5px solid #181828;
    border-radius:14px; z-index:2147483645; display:none; flex-direction:column;
    box-shadow:0 12px 60px #000e; overflow:hidden;
  }
  #ccee_panel.open { display:flex; }
  #ccee_hdr {
    padding:10px 14px; background:#0b0b18; border-bottom:1px solid #181828;
    display:flex; align-items:center; flex-shrink:0;
  }
  #ccee_hdr_title { color:#00c896; font-size:12px; letter-spacing:1.5px; font-weight:700; flex:1; }
  #ccee_hdr_ver   { color:#2a2a3e; font-size:10px; }
  #ccee_tabs {
    display:flex; background:#090912; border-bottom:1px solid #181828; flex-shrink:0;
  }
  .ccee_tab {
    padding:8px 16px; font-size:11px; color:#444; cursor:pointer;
    border-bottom:2px solid transparent; transition:all .15s; user-select:none;
  }
  .ccee_tab:hover { color:#888; }
  .ccee_tab.active { color:#00c896; border-bottom-color:#00c896; }
  .ccee_pane { display:none; flex:1; overflow-y:auto; padding:14px; }
  .ccee_pane.active { display:flex; flex-direction:column; gap:10px; }
  .ccee_pane::-webkit-scrollbar { width:3px; }
  .ccee_pane::-webkit-scrollbar-thumb { background:#2a2a3e; border-radius:3px; }
  .ccee_row   { display:flex; gap:10px; }
  .ccee_field { display:flex; flex-direction:column; flex:1; gap:5px; }
  .ccee_label { font-size:9px; color:#00c896; letter-spacing:1px; text-transform:uppercase; }
  .ccee_input {
    background:#0e0e1c; border:1px solid #252535; color:#ddd;
    padding:7px 10px; border-radius:6px; font-size:12px;
    font-family:inherit; outline:none; transition:border .15s;
  }
  .ccee_input:focus { border-color:#00c89660; }
  .ccee_hint {
    font-size:10px; color:#3a3a5a; line-height:1.6; padding:8px 10px;
    background:#0b0b18; border-radius:6px; border:1px solid #181828;
  }
  .ccee_hint b { color:#555; }
  .ccee_hint span { color:#00c896; }
  #ccee_emp_search {
    background:#0e0e1c; border:1px solid #252535; color:#ccc;
    padding:7px 10px; border-radius:6px; font-size:11px;
    font-family:inherit; outline:none; width:100%;
  }
  .ccee_emp_actions { display:flex; gap:6px; flex-wrap:wrap; }
  #ccee_emp_list {
    display:flex; flex-direction:column; gap:2px;
    flex:1; overflow-y:auto; min-height:0;
  }
  .ccee_emp_item {
    display:flex; align-items:center; gap:8px; padding:5px 8px;
    border-radius:4px; cursor:pointer; font-size:11px; color:#666;
    transition:background .1s;
  }
  .ccee_emp_item:hover { background:#0f0f20; }
  .ccee_emp_item.checked { color:#ccc; }
  .ccee_emp_item input[type=checkbox] { accent-color:#00c896; cursor:pointer; flex-shrink:0; }
  #ccee_pane_log { gap:6px; }
  #ccee_log_wrap {
    flex:1; overflow-y:auto; background:#060610; border:1px solid #181828;
    border-radius:6px; padding:10px;
  }
  #ccee_log { font-size:10.5px; color:#556; line-height:1.8; }
  #ccee_log .ok   { color:#00c896; }
  #ccee_log .err  { color:#e05060; }
  #ccee_log .warn { color:#f0b040; }
  #ccee_progress {
    padding:8px 14px; background:#090912; border-top:1px solid #181828;
    flex-shrink:0;
  }
  #ccee_status_txt { font-size:10px; color:#666; margin-bottom:5px; }
  #ccee_bar_bg { background:#161626; border-radius:99px; height:5px; overflow:hidden; }
  #ccee_bar_fill {
    background:linear-gradient(90deg,#00c896,#00a0c8);
    height:5px; border-radius:99px; width:0%; transition:width .4s ease;
  }
  #ccee_footer {
    padding:8px 14px; background:#090912; border-top:1px solid #181828;
    display:flex; gap:7px; flex-shrink:0;
  }
  .ccee_btn {
    flex:1; padding:8px; border-radius:7px; font-size:11px;
    font-family:inherit; cursor:pointer; border:1px solid;
    font-weight:700; letter-spacing:.5px; transition:all .15s;
  }
  .ccee_btn.start { background:#00c89618; border-color:#00c896; color:#00c896; }
  .ccee_btn.pause { background:#f0b04018; border-color:#f0b040; color:#f0b040; }
  .ccee_btn.stop  { background:#e0506018; border-color:#e05060; color:#e05060; }
  .ccee_btn:hover:not(:disabled) { filter:brightness(1.4); }
  .ccee_btn:disabled { opacity:.25; cursor:not-allowed; filter:none; }
  .ccee_sm {
    background:#0e0e1c; border:1px solid #252535; color:#666;
    padding:4px 10px; border-radius:5px; font-size:10px; cursor:pointer;
    font-family:inherit; transition:all .15s;
  }
  .ccee_sm:hover { color:#ddd; background:#181828; }
  </style>`;

  function buildPanelHTML(cfg) {
    const mesOpts = MESES.map(m =>
      `<option value="${m}"${cfg.mes === m ? ' selected' : ''}>${m}</option>`
    ).join('');

    return `
    ${CSS}
    <div id="ccee_toggle" title="CCEE Extrator (Alt+Shift+C)">
      ⚡<span id="ccee_badge">0</span>
    </div>

    <div id="ccee_panel">
      <div id="ccee_hdr">
        <span id="ccee_hdr_title">⚡ CCEE EXTRATOR</span>
        <span id="ccee_hdr_ver">v1.0.5</span>
      </div>

      <div id="ccee_tabs">
        <div class="ccee_tab active" data-tab="config">⚙ Config</div>
        <div class="ccee_tab" data-tab="empresas">🏢 Empresas</div>
        <div class="ccee_tab" data-tab="log">📋 Log</div>
      </div>

      <div class="ccee_pane active" id="ccee_pane_config">
        <div class="ccee_row">
          <div class="ccee_field">
            <div class="ccee_label">Nome do Relatório</div>
            <input class="ccee_input" id="ccee_rel" placeholder="ex: RCAP002" value="${cfg.relatorio || ''}">
          </div>
          <div class="ccee_field" style="max-width:90px">
            <div class="ccee_label">Mês</div>
            <select class="ccee_input" id="ccee_mes">${mesOpts}</select>
          </div>
          <div class="ccee_field" style="max-width:72px">
            <div class="ccee_label">Ano</div>
            <input class="ccee_input" id="ccee_ano" placeholder="26" value="${cfg.ano || ''}">
          </div>
        </div>
        <div class="ccee_hint">
          <b>Nome do arquivo gerado:</b><br>
          <span>{Empresa}_{Relatório}_{Mês}_{Ano}.pdf</span>
        </div>
        <div class="ccee_hint" style="margin-top:4px">
          <b>Como usar:</b><br>
          1. Preencha relatório, mês e ano acima<br>
          2. Vá em <b>🏢 Empresas</b> e marque quais extrair<br>
          3. Na CCEE, <b>navegue até o relatório correto</b><br>
          4. Volte aqui e clique em <span>▶ INICIAR</span>
        </div>
      </div>

      <div class="ccee_pane" id="ccee_pane_empresas">
        <input id="ccee_emp_search" placeholder="🔍 buscar empresa...">
        <div class="ccee_emp_actions">
          <button class="ccee_sm" id="ccee_selall">✓ Todas</button>
          <button class="ccee_sm" id="ccee_selnone">✗ Nenhuma</button>
          <button class="ccee_sm" id="ccee_addone">+ Adicionar</button>
          <button class="ccee_sm" id="ccee_resetlist" style="margin-left:auto">↺ Reset</button>
        </div>
        <div id="ccee_emp_list"></div>
      </div>

      <div class="ccee_pane" id="ccee_pane_log">
        <div style="display:flex;justify-content:flex-end">
          <button class="ccee_sm" id="ccee_clearlog">limpar</button>
        </div>
        <div id="ccee_log_wrap"><div id="ccee_log"></div></div>
      </div>

      <div id="ccee_progress">
        <div id="ccee_status_txt">Aguardando início...</div>
        <div id="ccee_bar_bg"><div id="ccee_bar_fill"></div></div>
      </div>

      <div id="ccee_footer">
        <button class="ccee_btn start" id="ccee_start">▶ INICIAR</button>
        <button class="ccee_btn pause" id="ccee_pause" disabled>⏸ PAUSAR</button>
        <button class="ccee_btn stop"  id="ccee_stop"  disabled>⏹ PARAR</button>
      </div>
    </div>`;
  }

  let todasEmpresas = [];

  function renderEmpresas(lista, filtro) {
    todasEmpresas = lista;
    const container = document.getElementById('ccee_emp_list');
    if (!container) return;
    const f = (filtro || '').toLowerCase();
    const visíveis = lista.filter(e => !f || e.toLowerCase().includes(f));
    container.innerHTML = visíveis.map(e => `
      <label class="ccee_emp_item checked">
        <input type="checkbox" checked data-emp="${e}">
        <span>${e}</span>
      </label>
    `).join('');
    atualizarBadge();
  }

  function getEmpresasSelecionadas() {
    return [...document.querySelectorAll('#ccee_emp_list input[type=checkbox]:checked')]
      .map(cb => cb.dataset.emp);
  }

  function atualizarBadge() {
    const b = document.getElementById('ccee_badge');
    const n = getEmpresasSelecionadas().length;
    if (b) b.textContent = n;
  }

  function uiLog(msg, cls) {
    const el = document.getElementById('ccee_log');
    if (!el) return;
    const line = document.createElement('div');
    if (cls) line.className = cls;
    line.innerHTML = msg;
    el.appendChild(line);
    const wrap = document.getElementById('ccee_log_wrap');
    if (wrap) wrap.scrollTop = wrap.scrollHeight;
    if (cls === 'err') switchTab('log');
  }

  function uiStatus(msg) {
    const el = document.getElementById('ccee_status_txt');
    if (el) el.textContent = msg;
  }

  function uiProgress(done, total) {
    const fill = document.getElementById('ccee_bar_fill');
    const txt  = document.getElementById('ccee_status_txt');
    if (fill) fill.style.width = total > 0 ? `${Math.round(done / total * 100)}%` : '0%';
    if (txt)  txt.textContent  = `${done} / ${total} empresas`;
  }

  function switchTab(name) {
    document.querySelectorAll('.ccee_tab')
      .forEach(t => t.classList.toggle('active', t.dataset.tab === name));
    document.querySelectorAll('.ccee_pane')
      .forEach(p => p.classList.toggle('active', p.id === `ccee_pane_${name}`));
  }

  function setBtnState(s) {
    const btnStart = document.getElementById('ccee_start');
    const btnPause = document.getElementById('ccee_pause');
    const btnStop  = document.getElementById('ccee_stop');
    if (!btnStart) return;

    if (s === 'idle') {
      btnStart.disabled = false;
      btnPause.disabled = true;
      btnStop.disabled  = true;
      btnPause.textContent = '⏸ PAUSAR';
    } else if (s === 'running') {
      btnStart.disabled = true;
      btnPause.disabled = false;
      btnStop.disabled  = false;
      btnPause.textContent = '⏸ PAUSAR';
    } else if (s === 'paused') {
      btnStart.disabled = true;
      btnPause.disabled = false;
      btnStop.disabled  = false;
      btnPause.textContent = '▶ RETOMAR';
    }
  }

  function bindEvents() {
    document.getElementById('ccee_toggle').onclick = () => {
      document.getElementById('ccee_panel').classList.toggle('open');
    };
    document.querySelectorAll('.ccee_tab').forEach(tab => {
      tab.onclick = () => switchTab(tab.dataset.tab);
    });
    document.getElementById('ccee_emp_search').oninput = e => {
      renderEmpresas(todasEmpresas, e.target.value);
    };
    document.getElementById('ccee_selall').onclick = () => {
      document.querySelectorAll('#ccee_emp_list input[type=checkbox]').forEach(cb => {
        cb.checked = true;
        cb.closest('.ccee_emp_item').classList.add('checked');
      });
      atualizarBadge();
    };
    document.getElementById('ccee_selnone').onclick = () => {
      document.querySelectorAll('#ccee_emp_list input[type=checkbox]').forEach(cb => {
        cb.checked = false;
        cb.closest('.ccee_emp_item').classList.remove('checked');
      });
      atualizarBadge();
    };
    document.getElementById('ccee_addone').onclick = () => {
      const nome = prompt('Nome exato da empresa:');
      if (!nome?.trim()) return;
      todasEmpresas.push(nome.trim().toUpperCase());
      saveEmpresas(todasEmpresas);
      renderEmpresas(todasEmpresas);
    };
    document.getElementById('ccee_resetlist').onclick = () => {
      if (!confirm('Resetar lista?')) return;
      todasEmpresas = [...EMPRESAS_DEFAULT];
      saveEmpresas(todasEmpresas);
      renderEmpresas(todasEmpresas);
    };
    document.getElementById('ccee_emp_list').addEventListener('change', e => {
      if (e.target.type === 'checkbox') {
        e.target.closest('.ccee_emp_item').classList.toggle('checked', e.target.checked);
        atualizarBadge();
      }
    });
    document.getElementById('ccee_clearlog').onclick = () => {
      const el = document.getElementById('ccee_log');
      if (el) el.innerHTML = '';
    };

    document.getElementById('ccee_start').onclick = async () => {
      const relatorio = document.getElementById('ccee_rel').value.trim();
      const mes       = document.getElementById('ccee_mes').value;
      const ano       = document.getElementById('ccee_ano').value.trim();

      if (!relatorio) { alert('Informe o relatório'); return; }
      if (!ano)        { alert('Informe o ano'); return; }

      const cfg = { relatorio, mes, ano };
      saveConfig(cfg);

      filaPendente = getEmpresasSelecionadas();
      if (!filaPendente.length) { alert('Selecione uma empresa'); return; }

      indiceAtual = 0;
      estado      = 'running';
      setBtnState('running');
      switchTab('log');

      uiLog(`═══ INICIANDO: <b>${relatorio}</b> | ${mes}/${ano} ═══`);
      uiLog(`${filaPendente.length} empresa(s) na fila`);
      uiLog(`<br>`);

      rodarAutomacao();
    };

    document.getElementById('ccee_pause').onclick = () => {
      if (estado === 'running') {
        estado = 'paused';
        setBtnState('paused');
        uiLog('⏸ Pausado.', 'warn');
      } else if (estado === 'paused') {
        estado = 'running';
        setBtnState('running');
        uiLog('▶ Retomado.', 'ok');
      }
    };

    document.getElementById('ccee_stop').onclick = () => {
      estado = 'stopped';
      setBtnState('idle');
      uiLog('⛔ Parado pelo usuário.', 'warn');
    };

    document.addEventListener('keydown', e => {
      if (e.altKey && e.shiftKey && e.code === 'KeyC') {
        document.getElementById('ccee_panel').classList.toggle('open');
      }
    });
  }

  function init() {
    if (!document.body) { setTimeout(init, 80); return; }
    const cfg      = loadConfig();
    const empresas = loadEmpresas();
    const wrap = document.createElement('div');
    wrap.id        = 'ccee_wrap';
    wrap.innerHTML = buildPanelHTML(cfg);
    document.body.appendChild(wrap);
    renderEmpresas(empresas);
    bindEvents();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
