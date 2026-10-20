// ==========================================
// SENHA DO ADMIN
// A senha NÃO fica escrita neste código. Ela fica guardada nas
// "Propriedades do script" com o nome ADMIN_SENHA:
// Apps Script > Configurações do projeto (engrenagem) > Propriedades do script >
// Adicionar propriedade do script > Propriedade: ADMIN_SENHA | Valor: (a senha)
// ==========================================
var ACOES_ADMIN_ = ["login", "add", "setConfig"];

function doPost(e) {
  try {
    var p = e.parameter;
    var action = p.action;
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // ==========================================
    // 0. AÇÕES DO ADMIN EXIGEM SENHA
    // ==========================================
    if (ACOES_ADMIN_.indexOf(action) !== -1) {
      var senhaCorreta = PropertiesService.getScriptProperties().getProperty("ADMIN_SENHA");
      if (!senhaCorreta) return json_({ ok: false, erro: "sem-senha" });
      if (String(p.senha || "") !== senhaCorreta) return json_({ ok: false, erro: "senha" });
    }

    if (action === "login") {
      return json_({ ok: true });
    }

    // ==========================================
    // 1. AÇÃO: RESERVAR PRESENTES
    // ==========================================
    if (action === "reserve") {
      var id = p.id;
      var nomeComprador = p.nome || p.reservadoPor;

      // Abre a aba de Presentes
      var sheet = spreadsheet.getSheetByName("Presentes");
      var data = sheet.getDataRange().getValues();

      // Procura a linha correspondente pelo ID (Coluna A)
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim() === id.toString().trim()) {
          // Força a gravação exatamente na Linha (i + 1) e Coluna 5 (Coluna E)
          sheet.getRange(i + 1, 5).setValue(nomeComprador);
          break;
        }
      }
    }

    // ==========================================
    // 2. AÇÃO: CONFIRMAR PRESENÇA (RSVP) — várias pessoas por vez (máx. 20)
    // Colunas: A Nome | B Adulto/Criança | C Data/hora | D Confirmado por
    // ==========================================
    if (action === "rsvp") {
      var sheet = spreadsheet.getSheetByName("RSVP");
      var pessoas = p.pessoas ? JSON.parse(p.pessoas) : [{ nome: p.nome, tipo: p.tipo }];
      var agora = new Date().toLocaleString("pt-BR");
      var confirmadoPor = String(pessoas[0].nome || "").trim();

      pessoas.slice(0, 20).forEach(function (pessoa) {
        var nome = String((pessoa && pessoa.nome) || "").trim();
        if (!nome) return;
        var tipo = pessoa.tipo === "Criança" ? "Criança" : "Adulto";
        sheet.appendRow([nome, tipo, agora, confirmadoPor]);
      });
    }

    // ==========================================
    // 3. AÇÃO: ADICIONAR NOVO PRESENTE (ADMIN)
    // ==========================================
    if (action === "add") {
      var item = p.item;
      var link = p.link;
      var imagem = p.imagem;

      var sheet = spreadsheet.getSheetByName("Presentes");
      var novoId = new Date().getTime().toString();

      sheet.appendRow([novoId, item, link, imagem, ""]);
    }

    // ==========================================
    // 4. AÇÃO: SALVAR DATA, HORÁRIO E ENDEREÇO DO EVENTO (ADMIN)
    // ==========================================
    if (action === "setConfig") {
      setConfig_(p);
    }

    return json_({ ok: true });

  } catch(error) {
    // Se der erro, ele tenta criar uma aba de Log para te avisar o que quebrou
    try {
      var logSheet = spreadsheet.getSheetByName("LOG_ERROS") || spreadsheet.insertSheet("LOG_ERROS");
      logSheet.appendRow([new Date(), error.toString()]);
    } catch(e) {}
    return json_({ ok: false, erro: error.toString() });
  }
}

function doGet(e) {
  var callback = e.parameter.callback;

  // ==========================================
  // DADOS DO EVENTO (data, horário e endereço) para o site
  // ==========================================
  if (e.parameter.action === "config") {
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(getConfig_()) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Presentes");
  var data = sheet.getDataRange().getValues();
  var gifts = [];

  for (var i = 1; i < data.length; i++) {
    if(data[i][1]) { // Só adiciona se o nome do item não estiver vazio
      gifts.push({
        id: data[i][0],
        item: data[i][1],
        link: data[i][2],
        imagem: data[i][3],
        reservadoPor: data[i][4]
      });
    }
  }

  var json = JSON.stringify(gifts);
  return ContentService.createTextOutput(callback + "(" + json + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// DATA, HORÁRIO E ENDEREÇO DO EVENTO
// Ficam guardados nas "Propriedades do script" (não precisa criar aba nova na planilha)
// ==========================================
var CONFIG_CAMPOS_ = ["data", "hora", "endereco", "cidade"];

function getConfig_() {
  var props = PropertiesService.getScriptProperties();
  var config = {};
  CONFIG_CAMPOS_.forEach(function (campo) {
    config[campo] = props.getProperty(campo) || "";
  });
  return config;
}

function setConfig_(params) {
  var novos = {};
  CONFIG_CAMPOS_.forEach(function (campo) {
    if (params[campo] !== undefined) novos[campo] = String(params[campo]).trim();
  });
  PropertiesService.getScriptProperties().setProperties(novos);
}
