const LOW_STOCK_THRESHOLD = 6;
const LOW_STOCK_TELEGRAM_THRESHOLD = 3;
const INVENTORY_SHEET = "Bestand";
const ORDERS_SHEET = "Bestellungen";
const ORDER_NOTIFICATION_EMAIL = "gerhard-diez@gmx.de, lukasdiezinger3@gmail.com";
const ORDER_SENDER_NAME = "Neue Bestellung";
const ORDER_EMAIL_SUBJECT = "[HONIG-BESTELLUNG] Neue Bestellung";
const MAX_ORDER_QUANTITIES = { arb3d: 3, koe3d: 3, wei3d: 3, volk3d: 5, weivol3d: 5 };

const INITIAL_PRODUCTS = [
  ["cf250", "Cremiger Honig Frühling", "250 g", 10],
  ["cf500", "Cremiger Honig Frühling", "500 g", 19],
  ["cs250", "Cremiger Honig Sommer", "250 g", 16],
  ["cs500", "Cremiger Honig Sommer", "500 g", 9],
  ["fs250", "Flüssiger Honig Sommer", "250 g", 13],
  ["fs500", "Flüssiger Honig Sommer", "500 g", 13],
  ["wb250", "Wald- und Blütenhonig", "250 g", 5],
  ["wb500", "Wald- und Blütenhonig", "500 g", 2],
  ["wh1", "Wabenhonig", "Stück", 11]
];
const INITIAL_CANDLES = [
  ["schma", "Kerze Schneemann", "6cm x 10cm x 5cm", 0],
  ["tabau", "Kerze Tannenbaum", "5cm x 8cm x 5cm", 0],
  ["bisto", "Kerze Bienenstock", "4cm x 5cm x 4cm", 0],
  ["weiwi", "Kerze Weihnachtswichtel", "4cm x 6cm x 3,5cm", 0]
];
const INITIAL_3D_PRINTS = [
  ["arb3d", "Arbeiter Bienchen", "Stück", 0],
  ["koe3d", "Königin", "Stück", 0],
  ["wei3d", "Weihnachtsbienchen", "Stück", 0],
  ["volk3d", "Bienchenvolk", "Stück", 0],
  ["weivol3d", "Weihnachtsbienchenvolk", "Stück", 0]
];

function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let inventory = spreadsheet.getSheetByName(INVENTORY_SHEET);
  if (!inventory) inventory = spreadsheet.insertSheet(INVENTORY_SHEET);
  const isNewInventory = inventory.getLastRow() === 0;
  if (isNewInventory) {
    inventory.getRange(1, 1, 1, 5).setValues([["ID", "Sorte", "Größe", "Bestand", "Nicht verfügbar"]]);
    inventory.setFrozenRows(1);
  }
  inventory.getRange(1, 5).setValue("Nicht verfügbar");
  let inventoryRows = inventory.getLastRow() > 1
    ? inventory.getRange(2, 1, inventory.getLastRow() - 1, 4).getValues()
    : [];
  let inventoryIds = new Set(inventoryRows.map(function(row) { return String(row[0]).trim().toUpperCase(); }));
  if (!INITIAL_PRODUCTS.some(function(product) { return inventoryIds.has(String(product[0]).trim().toUpperCase()); })) {
    inventory.appendRow(["HNG", "HONIG", "", ""]);
    inventory.appendRow(["", "", "", ""]);
    INITIAL_PRODUCTS.forEach(function(product) {
      inventory.appendRow(isNewInventory ? product : [product[0], product[1], product[2], ""]);
    });
  }
  inventoryRows = inventory.getLastRow() > 1
    ? inventory.getRange(2, 1, inventory.getLastRow() - 1, 4).getValues()
    : [];
  inventoryIds = new Set(inventoryRows.map(function(row) { return String(row[0]).trim().toUpperCase(); }));
  const candleHeaderIndexes = [];
  inventoryRows.forEach(function(row, index) {
    if (String(row[0]).trim().toUpperCase() === "KZN") candleHeaderIndexes.push(index);
  });
  candleHeaderIndexes.slice(1).reverse().forEach(function(index) {
    inventory.getRange(index + 2, 1, 1, 4).clearContent();
  });
  if (candleHeaderIndexes.length > 1) {
    inventoryRows = inventory.getLastRow() > 1
      ? inventory.getRange(2, 1, inventory.getLastRow() - 1, 4).getValues()
      : [];
    inventoryIds = new Set(inventoryRows.map(function(row) { return String(row[0]).trim().toUpperCase(); }));
  }
  if (!inventoryIds.has("KZN")) {
    const firstCandle = inventoryRows.findIndex(function(row) {
      return INITIAL_CANDLES.some(function(product) { return String(row[0]).trim() === product[0]; });
    });
    if (firstCandle === -1) {
      inventory.appendRow(["", "", "", ""]);
      inventory.appendRow(["KZN", "KERZEN", "BxHxT", ""]);
    } else {
      const headerRow = firstCandle + 2;
      inventory.insertRowsBefore(headerRow, 1);
      inventory.getRange(headerRow, 1, 1, 4).setValues([["KZN", "KERZEN", "BxHxT", ""]]);
    }
  }
  inventoryRows = inventory.getLastRow() > 1
    ? inventory.getRange(2, 1, inventory.getLastRow() - 1, 4).getValues()
    : [];
  INITIAL_CANDLES.forEach(function(product) {
    const rowIndex = inventoryRows.findIndex(function(row) { return String(row[0]).trim() === product[0]; });
    if (rowIndex === -1) {
      const newRow = isNewInventory ? product : [product[0], product[1], product[2], ""];
      inventory.appendRow(newRow);
      inventoryRows.push(newRow);
      return;
    }
    const sizeCell = inventory.getRange(rowIndex + 2, 3);
    if (String(sizeCell.getValue()).trim() !== product[2]) sizeCell.setValue(product[2]);
  });
  inventoryRows = inventory.getLastRow() > 1
    ? inventory.getRange(2, 1, inventory.getLastRow() - 1, 4).getValues()
    : [];
  inventoryIds = new Set(inventoryRows.map(function(row) { return String(row[0]).trim().toUpperCase(); }));
  if (!inventoryIds.has("K3D")) {
    const first3DPrint = inventoryRows.findIndex(function(row) {
      return INITIAL_3D_PRINTS.some(function(product) { return String(row[0]).trim() === product[0]; });
    });
    if (first3DPrint === -1) {
      inventory.appendRow(["", "", "", ""]);
      inventory.appendRow(["K3D", "KLEINKRAM (3D-DRUCK)", "", ""]);
    } else {
      const headerRow = first3DPrint + 2;
      inventory.insertRowsBefore(headerRow, 1);
      inventory.getRange(headerRow, 1, 1, 4).setValues([["K3D", "KLEINKRAM (3D-DRUCK)", "", ""]]);
    }
  }
  inventoryRows = inventory.getLastRow() > 1
    ? inventory.getRange(2, 1, inventory.getLastRow() - 1, 4).getValues()
    : [];
  INITIAL_3D_PRINTS.forEach(function(product) {
    const rowIndex = inventoryRows.findIndex(function(row) { return String(row[0]).trim() === product[0]; });
    if (rowIndex === -1) {
      const newRow = isNewInventory ? product : [product[0], product[1], product[2], ""];
      inventory.appendRow(newRow);
      inventoryRows.push(newRow);
    }
  });
  setupUnavailableCheckboxes_(inventory);
  inventory.autoResizeColumns(1, 5);

  let orders = spreadsheet.getSheetByName(ORDERS_SHEET);
  if (!orders) orders = spreadsheet.insertSheet(ORDERS_SHEET);
  if (orders.getLastRow() === 0) {
    orders.getRange(1, 1, 1, 8).setValues([[
      "Bestell-ID", "Zeitpunkt", "Name", "E-Mail", "Artikel", "Abholung", "Wunschtermin", "Wünsche"
    ]]);
    orders.setFrozenRows(1);
    orders.autoResizeColumns(1, 8);
  }
}

function doGet(e) {
  const callback = String((e && e.parameter && e.parameter.callback) || "");
  if (!/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
    return ContentService.createTextOutput("Ungültiger Callback");
  }

  try {
    const stock = readInventory_();
    const orderId = String((e && e.parameter && e.parameter.orderId) || "");
    let accepted;
    let outcomeReady = false;
    let orderMessage = "";
    if (orderId) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET);
      const rows = sheet && sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
        : [];
      accepted = rows.some(function(row) { return String(row[0]) === orderId; });
      outcomeReady = accepted;
      if (!accepted) {
        const cached = CacheService.getScriptCache().get("order_" + orderId);
        if (cached) {
          const outcome = JSON.parse(cached);
          accepted = outcome.accepted;
          orderMessage = outcome.message || "";
          outcomeReady = true;
        }
      }
    }
    const result = JSON.stringify({
      ok: true,
      accepted: accepted,
      outcomeReady: outcomeReady,
      stock: stock,
      unavailable: readUnavailable_(),
      lowStockThreshold: LOW_STOCK_THRESHOLD,
      message: orderMessage
    });
    return ContentService.createTextOutput(callback + "(" + result + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch (error) {
    const result = JSON.stringify({ ok: false, message: error.message });
    return ContentService.createTextOutput(callback + "(" + result + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

function doPost(e) {
  let orderId = "";
  try {
    const payload = JSON.parse((e && e.parameter && e.parameter.payload) || "{}");
    orderId = String(payload.orderId || "");
    const result = reserveOrder_(payload);
    if (orderId) CacheService.getScriptCache().put("order_" + orderId, JSON.stringify({ accepted: true }), 300);
    if (!result.duplicate) {
      try {
        sendNewOrderTelegram_(result.orderDetails);
      } catch (telegramError) {
        console.error("Telegram-Bestellmeldung konnte nicht erstellt werden: " + telegramError.message);
      }
      sendLowStockAlerts_(result.lowStockAlerts || []);
      try {
        sendNewOrderEmail_(result.orderDetails);
      } catch (emailError) {
        console.error("Bestell-E-Mail konnte nicht gesendet werden: " + emailError.message);
        sendTelegramMessage_("Bestell-E-Mail konnte nicht gesendet werden. Bestell-ID: " + orderId + ". Fehler: " + emailError.message);
      }
    }
    return HtmlService.createHtmlOutput("Bestellung verarbeitet.")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    if (orderId) {
      CacheService.getScriptCache().put("order_" + orderId, JSON.stringify({ accepted: false, message: error.message }), 300);
    }
    return HtmlService.createHtmlOutput("Bestellung konnte nicht verarbeitet werden.")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function logTelegramChatId() {
  const token = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
  if (!token) throw new Error("Lege zuerst TELEGRAM_BOT_TOKEN in den Script-Eigenschaften an.");

  const response = UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/getUpdates", {
    muteHttpExceptions: true
  });
  const result = JSON.parse(response.getContentText());
  if (!result.ok) throw new Error("Telegram konnte Updates nicht lesen. Prüfe den Bot-Token.");

  const chats = {};
  result.result.forEach(function(update) {
    const message = update.message || update.edited_message || update.channel_post;
    if (message && message.chat) {
      chats[String(message.chat.id)] = message.chat.title || message.chat.username || message.chat.first_name || "Chat";
    }
  });
  const ids = Object.keys(chats);
  if (!ids.length) {
    console.log("Keine Chats gefunden. Sende dem Bot zuerst /start und führe die Funktion erneut aus.");
    return;
  }
  ids.forEach(function(id) { console.log(chats[id] + ": " + id); });
}

function sendCurrentLowStockAlerts() {
  const inventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!inventory || inventory.getLastRow() < 2) throw new Error("Das Bestandsblatt ist nicht eingerichtet.");
  const rows = inventory.getRange(2, 1, inventory.getLastRow() - 1, 4).getValues();
  const alerts = rows.map(function(row) {
    return { id: String(row[0]), name: String(row[1]), size: String(row[2]), count: Number(row[3]) };
  }).filter(function(item) {
    return item.id && item.id !== "HNG" && item.id !== "KZN" && item.id !== "K3D" && Number.isInteger(item.count) && item.count >= 0 && item.count <= LOW_STOCK_TELEGRAM_THRESHOLD;
  });
  sendLowStockAlerts_(alerts);
  return alerts.length;
}

function sendLowStockAlerts_(alerts) {
  if (!alerts.length) return;
  alerts.forEach(function(item) {
    const message = "Lagerwarnung: " + item.name + " (" + item.size + ") hat nur noch " + item.count + ".";
    sendTelegramMessage_(message);
  });
}

function sendNewOrderTelegram_(order) {
  if (!order) return;
  const lines = [
    "Neue Honigbestellung",
    "Name: " + order.name,
    "Artikel: " + order.itemSummary,
    "Abholung: " + (order.pickup || "Keine Angabe")
  ];
  if (order.date) lines.push("Wunschtermin: " + order.date);
  sendTelegramMessage_(lines.join("\n"));
}

function sendNewOrderEmail_(order) {
  const lines = [
    "Neue Honigbestellung",
    "Bestell-ID: " + order.orderId,
    "Zeitpunkt: " + Utilities.formatDate(order.timestamp, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss"),
    "Name: " + order.name,
    "E-Mail: " + order.email,
    "Artikel: " + order.itemSummary,
    "Abholung: " + (order.pickup || "Keine Angabe"),
    "Wunschtermin: " + (order.date || "Keine Angabe"),
    "Wünsche: " + (order.note || "Keine Angabe")
  ];

  MailApp.sendEmail({
    to: ORDER_NOTIFICATION_EMAIL,
    name: ORDER_SENDER_NAME,
    replyTo: order.email,
    subject: ORDER_EMAIL_SUBJECT,
    body: lines.join("\n")
  });
}

function authorizeOrderEmail() {
  return MailApp.getRemainingDailyQuota();
}

function testOrderNotifications() {
  const testId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const errors = [];

  try {
    MailApp.sendEmail({
      to: ORDER_NOTIFICATION_EMAIL,
      name: ORDER_SENDER_NAME,
      subject: ORDER_EMAIL_SUBJECT + " TEST",
      body: "Dies ist eine Testmail aus Apps Script. Es wurde keine Bestellung angelegt. Test-ID: " + testId
    });
  } catch (error) {
    errors.push("E-Mail: " + error.message);
  }

  if (!sendTelegramMessage_("TEST - Bestellbenachrichtigung. Es wurde keine Bestellung angelegt. Test-ID: " + testId)) {
    errors.push("Telegram: Senden fehlgeschlagen. Prüfe TELEGRAM_BOT_TOKEN und TELEGRAM_CHAT_ID in den Script-Eigenschaften.");
  }

  if (errors.length) throw new Error(errors.join("\n"));
  console.log("Testmail und Telegram-Testnachricht wurden gesendet. Test-ID: " + testId);
}

function sendTelegramMessage_(message) {
  const properties = PropertiesService.getScriptProperties();
  const token = properties.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = properties.getProperty("TELEGRAM_CHAT_ID");
  if (!token || !chatId) {
    console.warn("Telegram-Nachricht übersprungen: TELEGRAM_BOT_TOKEN oder TELEGRAM_CHAT_ID fehlt.");
    return false;
  }

  try {
    const response = UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, text: message }),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (!result.ok) throw new Error("Telegram hat die Nachricht abgelehnt.");
    return true;
  } catch (error) {
    console.error("Telegram-Nachricht konnte nicht gesendet werden: " + error.message);
    return false;
  }
}

function readInventory_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("Das Bestandsblatt ist nicht eingerichtet.");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return rows.reduce(function(stock, row) {
    const id = String(row[0]);
    if (id && id !== "HNG" && id !== "KZN" && id !== "K3D") stock[id] = Number(row[3]);
    return stock;
  }, {});
}

function readUnavailable_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("Das Bestandsblatt ist nicht eingerichtet.");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return rows.reduce(function(unavailable, row) {
    const id = String(row[0]);
    if (id && id !== "HNG" && id !== "KZN" && id !== "K3D") unavailable[id] = row[4] === true;
    return unavailable;
  }, {});
}

function setupUnavailableCheckboxes_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  rows.forEach(function(row, index) {
    const id = String(row[0]).trim();
    if (!id || id === "HNG" || id === "KZN" || id === "K3D") return;
    const cell = sheet.getRange(index + 2, 5);
    cell.setDataValidation(checkboxRule);
    if (row[4] === "") cell.setValue(false);
  });
}

function reserveOrder_(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventory = spreadsheet.getSheetByName(INVENTORY_SHEET);
    const orders = spreadsheet.getSheetByName(ORDERS_SHEET);
    if (!inventory || !orders) throw new Error("Führe zuerst setup() aus.");

    const orderId = String(payload.orderId || "");
    const name = String(payload.name || "").trim();
    const email = String(payload.email || "").trim();
    const requested = Array.isArray(payload.items) ? payload.items : [];
    if (String(payload.honeypot || "").trim()) throw new Error("Die Bestellung konnte nicht geprüft werden.");
    if (!orderId || !name || !email || !requested.length) {
      throw new Error("Die Bestellung enthält nicht alle erforderlichen Angaben.");
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("Bitte gib eine gültige E-Mail-Adresse an.");

    const orderRows = orders.getLastRow() > 1
      ? orders.getRange(2, 1, orders.getLastRow() - 1, 8).getValues()
      : [];
    const previous = orderRows.find(function(row) { return String(row[0]) === orderId; });
    if (previous) return { ok: true, duplicate: true, stock: readInventory_(), unavailable: readUnavailable_(), lowStockThreshold: LOW_STOCK_THRESHOLD, lowStockAlerts: [] };

    const dataRange = inventory.getRange(2, 1, inventory.getLastRow() - 1, 4);
    const rows = dataRange.getValues();
    const unavailableRows = inventory.getRange(2, 5, inventory.getLastRow() - 1, 1).getValues();
    const rowById = {};
    rows.forEach(function(row, index) {
      const id = String(row[0]);
      if (!id || id === "HNG" || id === "KZN" || id === "K3D") return;
      const stock = Number(row[3]);
      if (!Number.isInteger(stock) || stock < 0) throw new Error("Im Bestandsblatt steht eine ungültige Stückzahl.");
      rowById[id] = { row: index + 2, name: String(row[1]), size: String(row[2]), stock: stock, unavailable: unavailableRows[index][0] === true };
    });

    const totals = requested.reduce(function(result, item) {
      const id = String(item.id || "");
      const quantity = Number(item.quantity);
      if (!Number.isInteger(quantity) || quantity < 1 || quantity > 20) {
        throw new Error("Eine ausgewählte Sorte oder Menge ist ungültig.");
      }
      result[id] = (result[id] || 0) + quantity;
      return result;
    }, {});

    const purchases = Object.keys(totals).map(function(id) {
      const quantity = totals[id];
      const product = rowById[id];
      if (!product) throw new Error("Eine ausgewählte Sorte oder Menge ist ungültig.");
      const maxQuantity = MAX_ORDER_QUANTITIES[id] || 20;
      if (quantity > maxQuantity) throw new Error("Für " + product.name + " sind höchstens " + maxQuantity + " Stück pro Bestellung möglich.");
      if (product.unavailable) throw new Error(product.name + " (" + product.size + ") ist derzeit nicht verfügbar.");
      if (quantity > product.stock) {
        throw new Error("Zu viel ausgewählt: " + product.name + " (" + product.size + "). Du hast " + quantity + " gewählt, verfügbar sind " + product.stock + ". Bitte reduziere die Menge.");
      }
      return { id: id, quantity: quantity, product: product };
    });

    const nextRows = rows.map(function(row) { return row.slice(); });
    purchases.forEach(function(item) {
      const index = item.product.row - 2;
      nextRows[index][3] = item.product.stock - item.quantity;
    });
    const lowStockAlerts = purchases.map(function(item) {
      return {
        name: item.product.name,
        size: item.product.size,
        count: item.product.stock - item.quantity
      };
    }).filter(function(item) {
      return item.count <= LOW_STOCK_TELEGRAM_THRESHOLD;
    });

    const itemSummary = purchases.map(function(item) {
      return item.quantity + "x " + item.product.name + " (" + item.product.size + ")";
    }).join(", ");
    const createdAt = new Date();
    dataRange.setValues(nextRows);
    try {
      orders.appendRow([
        orderId,
        createdAt,
        name,
        email,
        itemSummary,
        String(payload.pickup || ""),
        String(payload.date || ""),
        String(payload.note || "")
      ]);
      SpreadsheetApp.flush();
    } catch (error) {
      dataRange.setValues(rows);
      SpreadsheetApp.flush();
      throw error;
    }

    return {
      ok: true,
      stock: readInventory_(),
      unavailable: readUnavailable_(),
      lowStockThreshold: LOW_STOCK_THRESHOLD,
      lowStockAlerts: lowStockAlerts,
      orderDetails: {
        orderId: orderId,
        timestamp: createdAt,
        name: name,
        email: email,
        itemSummary: itemSummary,
        pickup: String(payload.pickup || ""),
        date: String(payload.date || ""),
        note: String(payload.note || "")
      }
    };
  } finally {
    lock.releaseLock();
  }
}
