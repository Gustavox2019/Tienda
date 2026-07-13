function doGet(e) {
  // Si la petición viene desde la aplicación web externa (GitHub Pages)
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var result = [];
    
    if (action === "getOrders") {
      result = getOrders();
    } else if (action === "getProductAutocompleteList") {
      result = getProductAutocompleteList();
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Por si acaso abren el enlace viejo de Google, dejamos el cargador original
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Pedidos Almacén - 10 Pisos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  var result = { success: false, message: "Acción no reconocida o datos incorrectos" };
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    
    if (action === "loginUser") {
      result = loginUser(data.username, data.password);
    } else if (action === "registerUser") {
      result = registerUser(data.username, data.password, data.role);
    } else if (action === "addOrder") {
      result = addOrder(data.product, data.qty, data.notes, data.photoBase64, data.senderName, data.senderAvatar);
    } else if (action === "updateOrder") {
      result = updateOrder(Number(data.rowNum), data.status, data.workerName, data.isEditingMode);
    }
  } catch (err) {
    result = { success: false, message: "Error interno en el servidor: " + err.toString() };
  }
  
  // Devolver el resultado de forma limpia con CORS habilitado
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
//          SISTEMA DE USUARIOS REALES
// ==========================================

function loginUser(username, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (!sheet) return { success: false, message: "Error: No existe la pestaña 'Usuarios'." };
  
  var data = sheet.getDataRange().getValues();
  var cleanUser = username.toString().trim().toLowerCase();
  var cleanPass = password.toString().trim();
  
  for (var i = 1; i < data.length; i++) {
    var dbUser = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    var dbPass = data[i][1] ? data[i][1].toString().trim() : "";
    
    if (dbUser === cleanUser && dbPass === cleanPass) {
      return {
        success: true,
        user: {
          name: data[i][0].toString().trim(),
          role: data[i][2] || "Vendedor"
        }
      };
    }
  }
  return { success: false, message: "Nombre de usuario o contraseña incorrectos." };
}

function registerUser(username, password, role) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (!sheet) return { success: false, message: "Error: No existe la pestaña 'Usuarios'." };
  
  var data = sheet.getDataRange().getValues();
  var cleanUser = username.toString().trim().toLowerCase();
  
  if (!username.toString().trim() || !password.toString().trim()) {
    return { success: false, message: "Por favor completa todos los campos." };
  }

  for (var i = 1; i < data.length; i++) {
    var dbUser = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    if (dbUser === cleanUser) {
      return { success: false, message: "Este nombre de usuario ya existe." };
    }
  }
  
  sheet.appendRow([username.toString().trim(), password.toString().trim(), role, "", ""]);
  return { 
    success: true, 
    user: { name: username.toString().trim(), role: role } 
  };
}

// ==========================================
//          SISTEMA DE PRODUCTOS Y PEDIDOS
// ==========================================

function getProductAutocompleteList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Productos");
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) list.push(values[i][0].toString().trim());
  }
  return list;
}

function getOrders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var orders = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || row[0].toString().trim() === "") continue;
    
    var order = { rowNum: i + 1 };
    for (var j = 0; j < headers.length; j++) {
      var headerName = headers[j].toString().trim();
      if (headerName) order[headerName] = row[j];
    }
    orders.push(order);
  }
  return orders.reverse(); 
}

function uploadToDrive(photoBase64) {
  if (!photoBase64 || photoBase64.indexOf("base64,") === -1) return "";
  try {
    var folderName = "Fotos_Pedidos_Almacen";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    var split = photoBase64.split("base64,");
    var contentType = split[0].match(/:(.*?);/)[1];
    var bytes = Utilities.base64Decode(split[1]);
    var blob = Utilities.newBlob(bytes, contentType, "pedido-" + new Date().getTime() + ".jpg");
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://docs.google.com/uc?export=view&id=" + file.getId();
  } catch (e) {
    return "";
  }
}

function addOrder(product, qty, notes, photoBase64, senderName, senderAvatar) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
  if (!sheet) return { success: false };
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = new Array(headers.length).fill("");
  
  var id = "PED-" + Utilities.formatDate(new Date(), "GMT-5", "yyyyMMdd-HHmmss");
  var timestamp = Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy HH:mm:ss");
  var photoUrl = photoBase64 ? uploadToDrive(photoBase64) : "";
  var historialInicial = "• Creado por " + senderName;

  for (var i = 0; i < headers.length; i++) {
    var h = headers[i].toString().trim();
    if (h === "ID_Pedido") newRow[i] = id;
    else if (h === "Fecha_Hora") newRow[i] = timestamp;
    else if (h === "Vendedor") newRow[i] = senderName;
    else if (h === "Producto") newRow[i] = product;
    else if (h === "Cantidad") newRow[i] = qty;
    else if (h === "Notas") newRow[i] = notes;
    else if (h === "Foto") newRow[i] = photoUrl;
    else if (h === "Estado") newRow[i] = "Por Tomar";
    else if (h === "Almacenero") newRow[i] = "";
    else if (h === "Historial_Cambios") newRow[i] = historialInicial;
  }
  
  sheet.appendRow(newRow);
  
  // Alerta de WhatsApp de entrada exclusiva para los almaceneros configurados
  var userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (userSheet) {
    var usersData = userSheet.getDataRange().getValues();
    var mensajeAlerta = "🚨 *NUEVO PEDIDO EN COLA*\n" +
                        "• *ID:* " + id + "\n" +
                        "• *Producto:* " + product + "\n" +
                        "• *Cantidad:* " + qty + "\n" +
                        "• *Vendedor:* " + senderName;
                        
    for (var u = 1; u < usersData.length; u++) {
      var dbRole = usersData[u][2] ? usersData[u][2].toString().trim() : "";
      var dbPhone = usersData[u][3] ? usersData[u][3].toString().trim() : "";
      var dbAPIKey = usersData[u][4] ? usersData[u][4].toString().trim() : "";
      
      if (dbRole === "Almacenero" && dbPhone !== "" && dbAPIKey !== "") {
        enviarMensajeWhatsApp(dbPhone, dbAPIKey, mensajeAlerta);
      }
    }
  }
  
  return { success: true };
}

function enviarMensajeWhatsApp(telefono, apikey, mensaje) {
  if (!telefono || !apikey || telefono.toString().trim() === "" || apikey.toString().trim() === "") return;
  try {
    var url = "https://api.callmebot.com/whatsapp.php?phone=" + 
              encodeURIComponent(telefono.toString().trim()) + 
              "&text=" + encodeURIComponent(mensaje) + 
              "&apikey=" + encodeURIComponent(apikey.toString().trim());
    var options = { "method": "get", "muteHttpExceptions": true };
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log("Error WhatsApp: " + e.toString());
  }
}

function updateOrder(rowNum, status, workerName, isEditingMode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
  if (!sheet) return { success: false };
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colEstado = headers.indexOf("Estado") + 1;
  var colAlmacenero = headers.indexOf("Almacenero") + 1;
  var colHistorial = headers.indexOf("Historial_Cambios") + 1;
  
  if (colEstado === 0 || colHistorial === 0) return { success: false };
  
  var estadoAnterior = sheet.getRange(rowNum, colEstado).getValue();
  var historialActual = sheet.getRange(rowNum, colHistorial).getValue() || "";
  
  if (estadoAnterior === status && !isEditingMode) return { success: true };
  
  var horaCambio = Utilities.formatDate(new Date(), "GMT-5", "HH:mm");
  var textoEdicion = isEditingMode ? " (Edición)" : "";
  var nuevaLineaHistorial = "\n• " + estadoAnterior + " ➔ " + status + " por " + workerName + " (" + horaCambio + ")" + textoEdicion;
  var historialActualizado = historialActual + nuevaLineaHistorial;
  
  sheet.getRange(rowNum, colEstado).setValue(status);
  sheet.getRange(rowNum, colHistorial).setValue(historialActualizado);
  
  if (status === "Tomado" && colAlmacenero > 0) {
    sheet.getRange(rowNum, colAlmacenero).setValue(workerName);
  }
  
  return { success: true };
}
