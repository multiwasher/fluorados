/**
 * Web App - Registos F-Gas
 * POST JSON -> grava no separador conforme "tipo"
 * GET ?action=ler{Tipo} -> lê dados das sheets
 *
 * Deploy:
 * Implementar > Nova implementação > App Web
 * Executar como: EU
 * Quem tem acesso: QUALQUER PESSOA
 */

const SPREADSHEET_ID = "1_dlHxttI4K93g801NIQEjXJhp1aFNUIQIKm8Y0WUEeI";

const SHEETS = {
  fugas: "Deteção de Fugas",
  intervencoes: "Restantes Intervenções",
  ensaios: "Ensaio Sist.Aut.Det.Fugas"
};

const HEADERS = {
  fugas: ["data", "equipamento", "nFicha", "nomeTecnico", "nCertTecnico", "nomeEmpresa", "nCertEmpresa", "moradaEmpresa", "telEmpresa", "locais", "resultado", "medidas", "posReparacao", "obs"],
  intervencoes: ["data", "equipamento", "nFicha", "nomeTecnico", "nCertTecnico", "nomeEmpresa", "nCertEmpresa", "moradaEmpresa", "telEmpresa", "fluido", "tipoIntervencao", "qAntes", "qRec", "qAdd", "qTotal", "obs"],
  ensaios: ["data", "equipamento", "nFicha", "nomeTecnico", "nCertTecnico", "nomeEmpresa", "nCertEmpresa", "moradaEmpresa", "telEmpresa", "resultado", "obs"]
};

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "";
    
    if (action === "lerFugas") return lerRegistos("fugas");
    if (action === "lerIntervencoes") return lerRegistos("intervencoes");
    if (action === "lerEnsaios") return lerRegistos("ensaios");
    
    return json_({ result: "error", message: `Action inválida: "${action}"` });
  } catch (error) {
    return json_({ result: "error", message: String(error) });
  }
}

function lerRegistos(tipo) {
  try {
    let ss;
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (err) {
      return json_({
        result: "error",
        message: "Não consegui abrir o ficheiro. Confirma ID e permissões."
      });
    }

    const sheetName = SHEETS[tipo];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return json_({ result: "error", message: `Separador "${sheetName}" não encontrado.` });
    }

    const headers = HEADERS[tipo];
    if (!headers) {
      return json_({ result: "error", message: `Tipo desconhecido: "${tipo}"` });
    }

    // Ler todos os dados (pulando a 1ª linha se for header)
    const data = sheet.getDataRange().getValues();
    if (!data || data.length === 0) {
      return json_({ registos: [] });
    }

    // Assumindo que a 1ª linha é header, começar a partir da linha 2
    const registos = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const obj = {};
      headers.forEach((header, idx) => {
        obj[header] = row[idx] || "";
      });
      registos.push(obj);
    }

    return json_({ registos: registos });
  } catch (error) {
    return json_({ result: "error", message: String(error) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ result: "error", message: "Pedido sem body (postData.contents vazio)." });
    }

    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      return json_({ result: "error", message: "JSON inválido: " + String(err) });
    }

    const tipo = String(data.tipo || "").trim().toLowerCase();
    if (!tipo) {
      return json_({ result: "error", message: "Campo 'tipo' em falta." });
    }
    if (!SHEETS[tipo]) {
      return json_({
        result: "error",
        message: `Tipo inválido: "${tipo}". Tipos aceites: ${Object.keys(SHEETS).join(", ")}`
      });
    }

    let ss;
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (err) {
      return json_({
        result: "error",
        message:
          "Não consegui abrir o ficheiro por ID. Confirma que o ID está certo e que o Web App está a executar como 'Eu'. Detalhe: " +
          String(err)
      });
    }

    const sheetName = SHEETS[tipo];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return json_({ result: "error", message: `Separador "${sheetName}" não encontrado.` });
    }

    const row = buildRow_(tipo, data);

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      sheet.appendRow(row);
    } finally {
      lock.releaseLock();
    }

    return json_({
      result: "success",
      fileId: ss.getId(),
      sheet: sheet.getName(),
      row: sheet.getLastRow()
    });

  } catch (error) {
    return json_({ result: "error", message: String(error) });
  }
}

function buildRow_(tipo, data) {
  if (tipo === "fugas") {
    return [
      data.data || "",
      data.equipamento || "",
      data.nFicha || "",
      data.nomeTecnico || "",
      data.nCertTecnico || "",
      data.nomeEmpresa || "",
      data.nCertEmpresa || "",
      data.moradaEmpresa || "",
      data.telEmpresa || "",
      data.locais || "",
      data.resultado || "",
      data.medidas || "",
      data.posReparacao || "",
      data.obs || ""
    ];
  }

  if (tipo === "intervencoes") {
    return [
      data.data || "",
      data.equipamento || "",
      data.nFicha || "",
      data.nomeTecnico || "",
      data.nCertTecnico || "",
      data.nomeEmpresa || "",
      data.nCertEmpresa || "",
      data.moradaEmpresa || "",
      data.telEmpresa || "",
      data.fluido || "",
      data.tipoIntervencao || "",
      data.qAntes || "",
      data.qRec || "",
      data.qAdd || "",
      data.qTotal || "",
      data.obs || ""
    ];
  }

  // ensaios
  return [
    data.data || "",
    data.equipamento || "",
    data.nFicha || "",
    data.nomeTecnico || "",
    data.nCertTecnico || "",
    data.nomeEmpresa || "",
    data.nCertEmpresa || "",
    data.moradaEmpresa || "",
    data.telEmpresa || "",
    data.resultado || "",
    data.obs || ""
  ];
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Executa 1x no editor para forçar autorização (recomendado)
function TESTE_autorizacao() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log(ss.getName());
}
