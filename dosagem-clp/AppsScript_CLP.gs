/**
 * PERTIGA CALCULADORA DE DOSAGEM — Apps Script Unificado
 * Planilha: https://docs.google.com/spreadsheets/d/1H5I8PG1EcAW7T-1-Fa6fmc-NqxMMFFOFxEurWpwMlsM/edit?usp=sharing
 *
 * COMO USAR:
 * 1. Abra a planilha → Extensões → Apps Script
 * 2. Apague tudo no editor e cole este código
 * 3. Salve (Ctrl+S)
 * 4. Clique em Executar → setup  (cria todas as abas)
 * 5. Implantar → Nova implantação
 *      Tipo: App da Web
 *      Executar como: Eu (seu e-mail)
 *      Quem tem acesso: Qualquer pessoa
 * 6. Copie a URL gerada e cole nas variáveis
 *      DOSAGEM_WEBAPP_URL  e  CUSTO_MCC_WEBAPP_URL
 *    no arquivo index.html (pwa-clp/dosagem-clp/index.html)
 *
 * ABAS CRIADAS:
 *   Traços            ← salvar_traco / listar_tracos
 *   Experimental      ← salvar_experimental / listar_experimentais
 *   Cartas_Traco      ← salvar_carta_traco / listar_cartas_traco
 *   MCC_Produtos      ← salvar_custo_mcc / deletar_custo_mcc / listar_custo_mcc
 *   Granulometria_Miudo   ← salvar_ensaio_granulometria (tipo=miudo)
 *   Granulometria_Graudo  ← salvar_ensaio_granulometria (tipo=graudo)
 *   Massa_Especifica  ← salvar_ensaio_me / listar_ensaios_me
 *   Fornecedores      ← salvar_fornecedor / listar_fornecedores
 */

var SPREADSHEET_ID = "1H5I8PG1EcAW7T-1-Fa6fmc-NqxMMFFOFxEurWpwMlsM";

/* ══════════════════════════════════════════════════════════════════╗
   CABEÇALHOS DE CADA ABA
╚══════════════════════════════════════════════════════════════════ */

var HEADERS = {
  tracos: [
    "ID","Data e Hora","Nome do Traço","FCK (MPa)",
    "Cimento (kg)","Adição (kg)","Areia Natural (kg)","Areia Industrial (kg)",
    "Brita 0 (kg)","Brita ½ (kg)","Água (kg)","Aditivo SP (kg)",
    "Ar Incorporado (%)","Flow (mm)","a/c","a/agl",
    "Consumo Cimento (kg/m³)","Massa Esp. Concreto (kg/m³)","Volume Total (L)",
    "Teor Argamassa Massa (%)","Teor Argamassa Volume (%)",
    "Cliente","Obra","Engenheiro","Observações"
  ],
  experimental: [
    "ID","Data e Hora","Nome do Traço","FCK (MPa)",
    "Cimento (kg)","Adição (kg)","Areia Natural (kg)","Areia Industrial (kg)",
    "Brita 0 (kg)","Brita ½ (kg)","Água (kg)","Aditivo SP (kg)",
    "Ar Incorporado (%)","Volume Betonada (L)","Nº Betonadas",
    "Umid. Areia Nat. (%)","Umid. Areia Ind. (%)","Umid. Brita 0 (%)","Umid. Brita ½ (%)",
    "Tipo Consistência","Consistência (mm)","Temp. Ambiente (°C)","Temp. Concreto (°C)",
    "a/c","a/agl","Consumo Cimento (kg/m³)",
    "R24h CP1 (MPa)","R24h CP2 (MPa)","R3d CP1 (MPa)","R3d CP2 (MPa)",
    "R7d CP1 (MPa)","R7d CP2 (MPa)","R28d CP1 (MPa)","R28d CP2 (MPa)",
    "Consist. 10min (mm)","Consist. 20min (mm)","Consist. 30min (mm)","Observações"
  ],
  cartas_traco: [
    "ID","Data e Hora","Cliente","Obra","Local","Aplicação",
    "Engenheiro","ART / RRT","Data Carta","Obs",
    "Nome Traço","FCK (MPa)","Flow (mm)","Ar (%)",
    "Cimento (kg)","Adição (kg)","Areia Natural (kg)","Areia Industrial (kg)",
    "Brita 0 (kg)","Brita ½ (kg)","Água (kg)","Aditivo SP (kg)"
  ],
  mcc_produtos: [
    "ID","Data Cadastro","Nome Produto","Tipo Material","Chave MCC",
    "Fornecedor","Unidade (MCC)","Versão","Unidade Medida",
    "Preço (R$/unid.)","ICMS (%)","Deduz ICMS",
    "Frete (R$/unid.)","Unid. Frete","ICMS Frete (%)","Deduz ICMS Frete",
    "Custo Líq. R$/kg","Ativo","IPI (%)","Deduz IPI","Observações"
  ],
  gran_miudo: [
    "ID","Data e Hora","Material","Fornecedor","Local","Operador",
    "#4 (%)","#8 (%)","#16 (%)","#30 (%)","#50 (%)","#100 (%)","Fundo (%)",
    "Módulo Finura","Dimensão Máxima (mm)","Observações"
  ],
  gran_graudo: [
    "ID","Data e Hora","Material","Fornecedor","Local","Operador",
    "1½\" (%)","1\" (%)","¾\" (%)","½\" (%)","3/8\" (%)","#4 (%)","#8 (%)","Fundo (%)",
    "Módulo Finura","Dimensão Máxima (mm)","Observações"
  ],
  massa_especifica: [
    "ID","Data e Hora","Material","Tipo","Fornecedor","Local","Operador",
    "Método","M1 (g)","M2 (g)","M3 (g)","V (mL)","ME (kg/dm³)","Observações"
  ],
  fornecedores: [
    "ID","Data Cadastro","Nome / Razão Social","CNPJ","Contato",
    "Telefone","E-mail","Cidade","UF","Material Principal","Observações"
  ]
};

/* ══════════════════════════════════════════════════════════════════╗
   HELPER: abre planilha e cria aba se necessário
╚══════════════════════════════════════════════════════════════════ */

function getOrCreateSheet(ss, name, headers, headerColor) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.appendRow(headers);
    var r = sh.getRange(1, 1, 1, headers.length);
    r.setFontWeight("bold");
    r.setBackground(headerColor || "#1a5276");
    r.setFontColor("#FFFFFF");
    r.setHorizontalAlignment("center");
    r.setWrap(false);
    sh.setFrozenRows(1);
    for (var i = 1; i <= headers.length; i++) sh.setColumnWidth(i, 140);
    // Primeira coluna (ID) mais larga
    sh.setColumnWidth(1, 200);
    sh.setColumnWidth(2, 160);
  }
  return sh;
}

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/* ══════════════════════════════════════════════════════════════════╗
   doPost — ROTEAMENTO
╚══════════════════════════════════════════════════════════════════ */

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = getSS();

    /* ── salvar_traco ─────────────────────────────────────────── */
    if (d.action === "salvar_traco") {
      var sh = getOrCreateSheet(ss, "Traços", HEADERS.tracos, "#1a5276");
      var agora = new Date();
      var id = d.id || ("TRC-" + agora.getFullYear()
        + String(agora.getMonth()+1).padStart(2,"0")
        + String(agora.getDate()).padStart(2,"0")
        + "-" + Math.floor(Math.random()*90000+10000));
      sh.appendRow([
        id,
        d.dataHora          || agora.toLocaleString("pt-BR"),
        d.nomeTraco         || "",
        d.fck               || "",
        d.cimento           || "",
        d.adicao            || "",
        d.areiaNatural      || "",
        d.areiaIndustrial   || "",
        d.brita0            || "",
        d.brita1            || "",
        d.agua              || "",
        d.aditivoSP         || "",
        d.arIncorporado     || "",
        d.flow              || "",
        d.ac                || "",
        d.aAgl              || "",
        d.consumoCimento    || "",
        d.massaEspConcreto  || "",
        d.volumeTotal       || "",
        d.teorArgMassa      || "",
        d.teorArgVolume     || "",
        d.cliente           || "",
        d.obra              || "",
        d.engenheiro        || "",
        d.observacoes       || ""
      ]);
      return _ok("Traço salvo com sucesso!", id);
    }

    /* ── salvar_experimental ─────────────────────────────────── */
    if (d.action === "salvar_experimental") {
      var sh = getOrCreateSheet(ss, "Experimental", HEADERS.experimental, "#283593");
      var agora = new Date();
      var id = d.id || ("EXP-" + agora.getFullYear()
        + String(agora.getMonth()+1).padStart(2,"0")
        + String(agora.getDate()).padStart(2,"0")
        + "-" + Math.floor(Math.random()*90000+10000));
      sh.appendRow([
        id,
        d.dataHora          || agora.toLocaleString("pt-BR"),
        d.nomeTraco         || "",
        d.fck               || "",
        d.cimento           || "",
        d.adicao            || "",
        d.areiaNatural      || "",
        d.areiaIndustrial   || "",
        d.brita0            || "",
        d.brita1            || "",
        d.agua              || "",
        d.aditivoSP         || "",
        d.arIncorporado     || "",
        d.volumeBetonada    || "",
        d.numBetonadas      || "",
        d.umAN              || "",
        d.umAI              || "",
        d.umB0              || "",
        d.umB1              || "",
        d.tipoConsistencia  || "",
        d.consistencia      || "",
        d.tempAmbiente      || "",
        d.tempConcreto      || "",
        d.ac                || "",
        d.aAgl              || "",
        d.consumoCimento    || "",
        d.r24h_1            || "",
        d.r24h_2            || "",
        d.r3d_1             || "",
        d.r3d_2             || "",
        d.r7d_1             || "",
        d.r7d_2             || "",
        d.r28d_1            || "",
        d.r28d_2            || "",
        d.consist10min      || "",
        d.consist20min      || "",
        d.consist30min      || "",
        d.observacoes       || ""
      ]);
      return _ok("Traço experimental salvo!", id);
    }

    /* ── salvar_carta_traco ──────────────────────────────────── */
    if (d.action === "salvar_carta_traco") {
      var sh = getOrCreateSheet(ss, "Cartas_Traco", HEADERS.cartas_traco, "#6a1b9a");
      var agora = new Date();
      var id = d.id || ("CT-" + Date.now());
      sh.appendRow([
        id,
        d.dataHora      || agora.toLocaleString("pt-BR"),
        d.cliente       || "",
        d.obra          || "",
        d.local         || "",
        d.aplicacao     || "",
        d.engenheiro    || "",
        d.art           || "",
        d.data          || "",
        d.obs           || "",
        d.tracoNome     || "",
        d.fck           || "",
        d.flow          || "",
        d.ar            || "",
        d.cim           || "",
        d.adic          || "",
        d.an            || "",
        d.ai            || "",
        d.b0            || "",
        d.b1            || "",
        d.agua          || "",
        d.adit          || ""
      ]);
      return _ok("Carta Traço salva com sucesso!", id);
    }

    /* ── salvar_custo_mcc ────────────────────────────────────── */
    if (d.action === "salvar_custo_mcc") {
      var sh = getOrCreateSheet(ss, "MCC_Produtos", HEADERS.mcc_produtos, "#1b5e20");
      var lastRow = sh.getLastRow();

      // Atualiza registro por ID se passado
      if (d.id && lastRow > 1) {
        var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === String(d.id)) {
            sh.getRange(i + 2, 1, 1, HEADERS.mcc_produtos.length).setValues([[
              d.id,
              d.dataCadastro || new Date().toLocaleString("pt-BR"),
              d.nomeProduto     || "",
              d.tipoMaterial    || "",
              d.chaveMCC        || "",
              d.fornecedor      || "",
              d.unidadeMCC      || "",
              Number(d.versao)  || 1,
              d.unidadeMedida   || "kg",
              d.preco           || 0,
              d.icmsPct         || 0,
              d.deduzICMS       || "Não",
              d.frete           || 0,
              d.unidFrete       || "kg",
              d.icmsFreteP      || 0,
              d.deduzICMSFrete  || "Não",
              d.custoLiq        || 0,
              d.ativo           || "Sim",
              d.ipiPct          || 0,
              d.deduzIPI        || "Não",
              d.observacoes     || ""
            ]]);
            return _ok("Produto atualizado.", d.id);
          }
        }
      }

      // Novo registro — versionamento automático
      var nomeProduto = d.nomeProduto || "";
      var fornecedor  = d.fornecedor  || "";
      var unidadeMCC  = d.unidadeMCC  || "";
      var maxVersao   = 0;
      if (lastRow > 1) {
        var allData = sh.getRange(2, 1, lastRow - 1, HEADERS.mcc_produtos.length).getValues();
        for (var j = 0; j < allData.length; j++) {
          if (String(allData[j][2]).trim().toLowerCase() === nomeProduto.trim().toLowerCase() &&
              String(allData[j][5]).trim().toLowerCase() === fornecedor.trim().toLowerCase()  &&
              String(allData[j][6]).trim().toLowerCase() === unidadeMCC.trim().toLowerCase()) {
            sh.getRange(j + 2, 18).setValue("Não"); // col 18 = Ativo
            var v = Number(allData[j][7]) || 1;
            if (v > maxVersao) maxVersao = v;
          }
        }
      }
      var novaVersao = maxVersao + 1;
      var agora = new Date();
      var newId = "MCC-" + agora.getFullYear()
        + String(agora.getMonth()+1).padStart(2,"0")
        + String(agora.getDate()).padStart(2,"0")
        + "-" + Math.floor(Math.random()*90000+10000);
      sh.appendRow([
        newId,
        agora.toLocaleString("pt-BR"),
        nomeProduto,
        d.tipoMaterial    || "",
        d.chaveMCC        || "",
        fornecedor,
        unidadeMCC,
        novaVersao,
        d.unidadeMedida   || "kg",
        d.preco           || 0,
        d.icmsPct         || 0,
        d.deduzICMS       || "Não",
        d.frete           || 0,
        d.unidFrete       || "kg",
        d.icmsFreteP      || 0,
        d.deduzICMSFrete  || "Não",
        d.custoLiq        || 0,
        "Sim",
        d.ipiPct          || 0,
        d.deduzIPI        || "Não",
        d.observacoes     || ""
      ]);
      var msg = novaVersao > 1
        ? "Versão " + novaVersao + " salva. Versão anterior inativada."
        : "Produto salvo.";
      return _ok(msg, newId);
    }

    /* ── deletar_custo_mcc ───────────────────────────────────── */
    if (d.action === "deletar_custo_mcc") {
      var sh = getOrCreateSheet(ss, "MCC_Produtos", HEADERS.mcc_produtos, "#1b5e20");
      var last = sh.getLastRow();
      if (last > 1) {
        var ids = sh.getRange(2, 1, last - 1, 1).getValues();
        for (var i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]) === String(d.id)) {
            sh.deleteRow(i + 2);
            return _ok("Produto excluído.", d.id);
          }
        }
      }
      return _err("ID não encontrado.");
    }

    /* ── salvar_ensaio_granulometria ─────────────────────────── */
    if (d.action === "salvar_ensaio_granulometria") {
      var isMiudo = (d.tipo || "").toLowerCase() === "miudo";
      var tabName = isMiudo ? "Granulometria_Miudo" : "Granulometria_Graudo";
      var hdrs    = isMiudo ? HEADERS.gran_miudo    : HEADERS.gran_graudo;
      var cor     = isMiudo ? "#e65100"             : "#bf360c";
      var sh = getOrCreateSheet(ss, tabName, hdrs, cor);
      var agora = new Date();
      var id = d.id || ("GR-" + Date.now());
      if (isMiudo) {
        sh.appendRow([
          id,
          d.dataHora   || agora.toLocaleString("pt-BR"),
          d.material   || "",
          d.fornecedor || "",
          d.local      || "",
          d.operador   || "",
          d.p4         || "", d.p8     || "", d.p16   || "", d.p30  || "",
          d.p50        || "", d.p100   || "", d.fundo || "",
          d.moduloFinura || "", d.dimMax || "",
          d.observacoes  || ""
        ]);
      } else {
        sh.appendRow([
          id,
          d.dataHora   || agora.toLocaleString("pt-BR"),
          d.material   || "",
          d.fornecedor || "",
          d.local      || "",
          d.operador   || "",
          d.p38mm || "", d.p25mm || "", d.p19mm  || "", d.p12mm || "",
          d.p9mm  || "", d.p4    || "", d.p8     || "", d.fundo || "",
          d.moduloFinura || "", d.dimMax || "",
          d.observacoes  || ""
        ]);
      }
      return _ok("Granulometria salva!", id);
    }

    /* ── salvar_ensaio_me ────────────────────────────────────── */
    if (d.action === "salvar_ensaio_me") {
      var sh = getOrCreateSheet(ss, "Massa_Especifica", HEADERS.massa_especifica, "#004d40");
      var agora = new Date();
      var id = d.id || ("ME-" + Date.now());
      sh.appendRow([
        id,
        d.dataHora   || agora.toLocaleString("pt-BR"),
        d.material   || "",
        d.tipo       || "",
        d.fornecedor || "",
        d.local      || "",
        d.operador   || "",
        d.metodo     || "",
        d.m1         || "", d.m2 || "", d.m3 || "",
        d.v          || "",
        d.me         || "",
        d.observacoes || ""
      ]);
      return _ok("Massa específica salva!", id);
    }

    /* ── salvar_fornecedor ───────────────────────────────────── */
    if (d.action === "salvar_fornecedor") {
      var sh = getOrCreateSheet(ss, "Fornecedores", HEADERS.fornecedores, "#37474f");
      var agora = new Date();
      var id = d.id || ("FORN-" + Date.now());
      sh.appendRow([
        id,
        d.dataCadastro     || agora.toLocaleString("pt-BR"),
        d.nomeRazao        || "",
        d.cnpj             || "",
        d.contato          || "",
        d.telefone         || "",
        d.email            || "",
        d.cidade           || "",
        d.uf               || "",
        d.materialPrincipal || "",
        d.observacoes      || ""
      ]);
      return _ok("Fornecedor salvo!", id);
    }

    return _err("action desconhecida: " + (d.action || "vazio"));

  } catch (err) {
    return _err(err.message);
  }
}

/* ══════════════════════════════════════════════════════════════════╗
   doGet — LEITURA / PING
╚══════════════════════════════════════════════════════════════════ */

function doGet(e) {
  function respond(obj) {
    var json = JSON.stringify(obj);
    var cb   = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
    if (cb) {
      return ContentService
        .createTextOutput(cb + "(" + json + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
    var ss = getSS();

    /* ── listar_tracos ──────────────────────────────────────── */
    if (action === "listar_tracos") {
      var sh = getOrCreateSheet(ss, "Traços", HEADERS.tracos, "#1a5276");
      var rows = _getRows(sh, HEADERS.tracos.length);
      var keys = ["id","dataHora","nomeTraco","fck","cimento","adicao","areiaNatural","areiaIndustrial",
                  "brita0","brita1","agua","aditivoSP","arIncorporado","flow","ac","aAgl",
                  "consumoCimento","massaEspConcreto","volumeTotal","teorArgMassa","teorArgVolume",
                  "cliente","obra","engenheiro","observacoes"];
      return respond({ ok: true, tracos: _mapRows(rows, keys) });
    }

    /* ── listar_experimentais ───────────────────────────────── */
    if (action === "listar_experimentais") {
      var sh = getOrCreateSheet(ss, "Experimental", HEADERS.experimental, "#283593");
      var rows = _getRows(sh, HEADERS.experimental.length);
      var keys = ["id","dataHora","nomeTraco","fck","cimento","adicao","areiaNatural","areiaIndustrial",
                  "brita0","brita1","agua","aditivoSP","arIncorporado","volumeBetonada","numBetonadas",
                  "umAN","umAI","umB0","umB1","tipoConsistencia","consistencia","tempAmbiente","tempConcreto",
                  "ac","aAgl","consumoCimento","r24h_1","r24h_2","r3d_1","r3d_2","r7d_1","r7d_2",
                  "r28d_1","r28d_2","consist10min","consist20min","consist30min","observacoes"];
      return respond({ ok: true, tracos: _mapRows(rows, keys) });
    }

    /* ── listar_cartas_traco ────────────────────────────────── */
    if (action === "listar_cartas_traco") {
      var sh = getOrCreateSheet(ss, "Cartas_Traco", HEADERS.cartas_traco, "#6a1b9a");
      var rows = _getRows(sh, HEADERS.cartas_traco.length);
      var keys = ["id","dataHora","cliente","obra","local","aplicacao","engenheiro","art","data","obs",
                  "tracoNome","fck","flow","ar","cim","adic","an","ai","b0","b1","agua","adit"];
      return respond({ ok: true, cartas: _mapRows(rows, keys) });
    }

    /* ── listar_custo_mcc ───────────────────────────────────── */
    if (action === "listar_custo_mcc") {
      var sh = getOrCreateSheet(ss, "MCC_Produtos", HEADERS.mcc_produtos, "#1b5e20");
      var rows = _getRows(sh, HEADERS.mcc_produtos.length);
      var keys = ["id","dataCadastro","nomeProduto","tipoMaterial","chaveMCC","fornecedor",
                  "unidadeMCC","versao","unidadeMedida","preco","icmsPct","deduzICMS",
                  "frete","unidFrete","icmsFreteP","deduzICMSFrete","custoLiq","ativo",
                  "ipiPct","deduzIPI","observacoes"];
      var produtos = _mapRows(rows, keys);
      var filtroChave   = (e.parameter.chave   || "").trim().toLowerCase();
      var filtroUnidade = (e.parameter.unidade || "").trim().toLowerCase();
      if (filtroChave)   produtos = produtos.filter(function(p){ return String(p.chaveMCC).toLowerCase()   === filtroChave;   });
      if (filtroUnidade) produtos = produtos.filter(function(p){ return String(p.unidadeMCC).toLowerCase() === filtroUnidade; });
      return respond({ ok: true, produtos: produtos });
    }

    /* ── listar_fornecedores ────────────────────────────────── */
    if (action === "listar_fornecedores") {
      var sh = getOrCreateSheet(ss, "Fornecedores", HEADERS.fornecedores, "#37474f");
      var rows = _getRows(sh, HEADERS.fornecedores.length);
      var keys = ["id","dataCadastro","nomeRazao","cnpj","contato","telefone","email","cidade","uf","materialPrincipal","observacoes"];
      return respond({ ok: true, fornecedores: _mapRows(rows, keys) });
    }

    /* ── ping ───────────────────────────────────────────────── */
    return respond({
      status: "online",
      planilha: SPREADSHEET_ID,
      abas: ["Traços","Experimental","Cartas_Traco","MCC_Produtos",
             "Granulometria_Miudo","Granulometria_Graudo","Massa_Especifica","Fornecedores"],
      actions_post: ["salvar_traco","salvar_experimental","salvar_carta_traco",
                     "salvar_custo_mcc","deletar_custo_mcc",
                     "salvar_ensaio_granulometria","salvar_ensaio_me","salvar_fornecedor"],
      actions_get:  ["listar_tracos","listar_experimentais","listar_cartas_traco",
                     "listar_custo_mcc","listar_fornecedores"]
    });

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ══════════════════════════════════════════════════════════════════╗
   HELPERS INTERNOS
╚══════════════════════════════════════════════════════════════════ */

function _ok(msg, id) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: msg, id: id || null }))
    .setMimeType(ContentService.MimeType.JSON);
}

function _err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function _getRows(sh, numCols) {
  var last = sh.getLastRow();
  if (last <= 1) return [];
  return sh.getRange(2, 1, last - 1, numCols).getValues();
}

function _mapRows(rows, keys) {
  return rows.map(function(row) {
    var obj = {};
    keys.forEach(function(k, i) { obj[k] = row[i]; });
    return obj;
  });
}

/* ══════════════════════════════════════════════════════════════════╗
   SETUP — executar UMA VEZ para criar todas as abas
╚══════════════════════════════════════════════════════════════════ */

function setup() {
  var ss = getSS();
  getOrCreateSheet(ss, "Traços",               HEADERS.tracos,           "#1a5276");
  getOrCreateSheet(ss, "Experimental",          HEADERS.experimental,     "#283593");
  getOrCreateSheet(ss, "Cartas_Traco",          HEADERS.cartas_traco,     "#6a1b9a");
  getOrCreateSheet(ss, "MCC_Produtos",          HEADERS.mcc_produtos,     "#1b5e20");
  getOrCreateSheet(ss, "Granulometria_Miudo",   HEADERS.gran_miudo,       "#e65100");
  getOrCreateSheet(ss, "Granulometria_Graudo",  HEADERS.gran_graudo,      "#bf360c");
  getOrCreateSheet(ss, "Massa_Especifica",      HEADERS.massa_especifica, "#004d40");
  getOrCreateSheet(ss, "Fornecedores",          HEADERS.fornecedores,     "#37474f");

  SpreadsheetApp.getUi().alert(
    "✅ Pertiga Calculadora de Dosagem\n\n" +
    "Abas criadas com sucesso:\n" +
    "  • Traços (" + HEADERS.tracos.length + " cols)\n" +
    "  • Experimental (" + HEADERS.experimental.length + " cols)\n" +
    "  • Cartas_Traco (" + HEADERS.cartas_traco.length + " cols)\n" +
    "  • MCC_Produtos (" + HEADERS.mcc_produtos.length + " cols)\n" +
    "  • Granulometria_Miudo (" + HEADERS.gran_miudo.length + " cols)\n" +
    "  • Granulometria_Graudo (" + HEADERS.gran_graudo.length + " cols)\n" +
    "  • Massa_Especifica (" + HEADERS.massa_especifica.length + " cols)\n" +
    "  • Fornecedores (" + HEADERS.fornecedores.length + " cols)\n\n" +
    "Próximo passo: Implantar como App da Web e copiar a URL."
  );
}
