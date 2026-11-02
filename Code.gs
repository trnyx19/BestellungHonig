const LOW_STOCK_THRESHOLD = 6;
const LOW_STOCK_TELEGRAM_THRESHOLD = 3;
const INVENTORY_SHEET = "Bestand";
const ORDERS_SHEET = "Bestellungen";
const ORDER_NOTIFICATION_EMAIL = "lukasdiezinger3@gmail.com";
const ORDER_SENDER_NAME = "Neue Bestellung";
const ORDER_EMAIL_SUBJECT = "[BESTELLUNG] Neue Bestellung";
const MAX_ORDER_QUANTITIES = { arb3d: 3, koe3d: 3, wei3d: 3, volk3d: 5, weivol3d: 5 };
const ORDER_STATUSES = ["Eingegangen", "In Bearbeitung", "Verzögert", "Bereit zur Abholung", "Abgeholt"];
const PRODUCT_AVAILABILITY_STATUSES = ["Normal", "Schnell zugreifen", "Nicht verfügbar"];
const UNAVAILABLE_REASONS = [
  "Nicht verfügbar",
  "Wartungsarbeiten - Website",
  "Vorübergehend nicht vorrätig",
  "Saisonpause",
  "Neue Ernte noch nicht verfügbar",
  "Verkaufsstart steht noch aus",
  "Produktion vorübergehend pausiert",
  "Rohstoffe oder Material fehlen",
  "Qualitätsprüfung",
  "Produkt wird überarbeitet",
  "Preis- oder Produktdaten werden aktualisiert",
  "Nicht mehr im Sortiment",
  "Abholung vorübergehend nicht möglich",
  "Sonstiger Grund"
];
const QUICK_ACCESS_REASONS = [
  "Bald aus dem Sortiment",
  "Nur noch Restbestand",
  "Saisonartikel, nur begrenzte Zeit verfügbar",
  "Modellwechsel",
  "Produkt wird überarbeitet",
  "Aktionspreis",
  "Sonstiger Grund"
];
const INITIAL_PRICES = {
  cf250: 3, cf500: 7, cs250: 3, cs500: 7, fs250: 3, fs500: 7, wb250: 3, wb500: 7,
  wh1: 3, arb3d: 0.20, koe3d: 0.40, wei3d: 0.30, volk3d: 1.30, weivol3d: 1.65
};

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
  ["fehldruck3d", "Bienchen Fehldruck", "Stück", 0],
  ["arb3d", "Arbeiter Bienchen", "Stück", 0],
  ["koe3d", "Königin", "Stück", 0],
  ["wei3d", "Weihnachtsbienchen", "Stück", 0],
  ["volk3d", "Bienchenvolk", "Ein Volk", 0],
  ["weivol3d", "Weihnachtsbienchenvolk", "Ein Volk", 0]
];

function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let inventory = spreadsheet.getSheetByName(INVENTORY_SHEET);
  if (!inventory) inventory = spreadsheet.insertSheet(INVENTORY_SHEET);
  const isNewInventory = inventory.getLastRow() === 0;
  if (!isNewInventory) migrateInventoryColumns_(inventory);
  if (isNewInventory) {
    inventory.getRange(1, 1, 1, 7).setValues([["ID", "Sorte", "Größe", "Bestand", "Verfügbarkeit", "Grund", "Preis (€)"]]);
    inventory.setFrozenRows(1);
  }
  inventory.getRange(1, 5).setValue("Verfügbarkeit");
  inventory.getRange(1, 6).setValue("Grund");
  inventory.getRange(1, 7).setValue("Preis (€)");
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
  initializeProductPrices_(inventory);
  inventory.getRange(1, 6).setNote("Für „Schnell zugreifen“ und „Nicht verfügbar“ gelten unterschiedliche Grundlisten.");
  inventory.getRange(1, 7).setNote("Wabenhonig: Preis je 100 g; bei allen anderen Artikeln Stückpreis.");
  if (inventory.getLastRow() > 1) {
    inventory.getRange(2, 7, inventory.getLastRow() - 1, 1).setNumberFormat('0.00 "€"');
  }
  setupProductAvailabilityDropdowns_(inventory);
  setupProductReasonDropdowns_(inventory);
  inventory.autoResizeColumns(1, 7);

  let orders = spreadsheet.getSheetByName(ORDERS_SHEET);
  if (!orders) orders = spreadsheet.insertSheet(ORDERS_SHEET);
  if (orders.getLastRow() === 0) {
    orders.getRange(1, 1, 1, 14).setValues([[
      "Bestell-ID", "Zeitpunkt", "Name", "E-Mail", "Artikel", "Abholung", "Wunschtermin", "Wünsche", "Bestellnummer", "Status", "Betrag (€)", "Preishinweis", "Druck-PDF", "PDF-Datei-ID"
    ]]);
    orders.setFrozenRows(1);
  }
  setupOrderTracking_(orders);
  setupOrderArchive_(spreadsheet);
  setupOrderStatusEmailTrigger_(spreadsheet);
  orders.autoResizeColumns(1, 13);
}

function doGet(e) {
  const callback = String((e && e.parameter && e.parameter.callback) || "");
  if (!/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
    return ContentService.createTextOutput("Ungültiger Callback");
  }

  try {
    const stock = readInventory_();
    const orderId = String((e && e.parameter && e.parameter.orderId) || "");
    let customerOrderNumber = "";
    let accepted;
    let outcomeReady = false;
    let orderMessage = "";
    if (orderId) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET);
      const rows = sheet && sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues()
        : [];
      const matchingOrder = rows.find(function(row) { return String(row[0]) === orderId; });
      accepted = Boolean(matchingOrder);
      if (matchingOrder) customerOrderNumber = String(matchingOrder[8] || "");
      outcomeReady = accepted;
      if (!accepted) {
        const cached = CacheService.getScriptCache().get("order_" + orderId);
        if (cached) {
          const outcome = JSON.parse(cached);
          accepted = outcome.accepted;
          customerOrderNumber = outcome.customerOrderNumber || "";
          orderMessage = outcome.message || "";
          outcomeReady = true;
        }
      }
    }
    const result = JSON.stringify({
      ok: true,
      accepted: accepted,
      outcomeReady: outcomeReady,
      customerOrderNumber: customerOrderNumber,
      stock: stock,
      productStatuses: readProductStatuses_(),
      productSizes: readProductSizes_(),
      prices: readPrices_(),
      unavailable: readUnavailable_(),
      unavailableReasons: readUnavailableReasons_(),
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
  let payload = {};
  try {
    payload = JSON.parse((e && e.parameter && e.parameter.payload) || "{}");
    orderId = String(payload.orderId || "");
    const result = reserveOrder_(payload);
    if (orderId) CacheService.getScriptCache().put("order_" + orderId, JSON.stringify({ accepted: true, customerOrderNumber: result.customerOrderNumber || "" }), 300);
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
      try {
        sendCustomerOrderConfirmation_(result.orderDetails);
      } catch (customerEmailError) {
        console.error("Kunden-Bestätigung konnte nicht gesendet werden: " + customerEmailError.message);
        sendTelegramMessage_("Kunden-Bestätigung konnte nicht gesendet werden. Bestellnummer: " + result.customerOrderNumber + ". Fehler: " + customerEmailError.message);
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
    "Bestellnummer: " + order.customerOrderNumber,
    "Name: " + order.name,
    "Artikel: " + order.itemSummary,
    "Abholung: " + (order.pickup || "Keine Angabe")
  ];
  if (order.date) lines.push("Wunschtermin: " + order.date);
  sendTelegramMessage_(lines.join("\n"));
}

function sendNewOrderEmail_(order) {
  const escapeHtml = function(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  };
  const valueOrDash = function(value) {
    const text = String(value == null ? "" : value).trim();
    return text ? text : "–";
  };
  const timestamp = Utilities.formatDate(order.timestamp, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
  const items = String(order.itemSummary || "").split(/,\s*(?=\d+x\s)/).filter(function(item) { return item; });
  const row = function(label, value) {
    return '<tr><td style="padding:7px 12px 7px 0;color:#888;vertical-align:top;">' + label +
      '</td><td style="padding:7px 0;color:#222;vertical-align:top;white-space:pre-wrap;">' + value + '</td></tr>';
  };
  const emailValue = order.email
    ? '<a href="mailto:' + escapeHtml(order.email) + '" style="color:#822045;">' + escapeHtml(order.email) + '</a>'
    : "–";
  const safeOrderNumber = valueOrDash(order.customerOrderNumber);
  const safeName = valueOrDash(order.name);
  const subject = ORDER_EMAIL_SUBJECT + " " + (order.customerOrderNumber || "") + " – " + (order.name || "");
  const html = '<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:20px 12px;background:#f5f5f5;font-family:-apple-system,\'Segoe UI\',Helvetica,Arial,sans-serif;font-size:15px;">' +
    '<table role="presentation" width="480" cellpadding="0" cellspacing="0" style="width:100%;max-width:480px;margin:0 auto;background:#fff;border-top:4px solid #822045;border-radius:6px;">' +
    '<tr><td style="padding:22px 24px;">' +
    '<div style="font-size:18px;font-weight:bold;color:#822045;">Neue Honigbestellung</div>' +
    '<div style="font-size:13px;color:#888;margin:2px 0 16px;">' + escapeHtml(timestamp) + '</div>' +
    '<div style="display:inline-block;padding:6px 12px;background:#FFE337;color:#822045;border-radius:6px;font-family:\'Courier New\',monospace;font-weight:bold;">' + escapeHtml(safeOrderNumber) + '</div>' +
    '<div style="margin:18px 0 6px;padding-top:14px;border-top:1px solid #eee;font-weight:bold;">' +
    items.map(function(item) { return escapeHtml(item); }).join("<br>") + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;font-size:14px;">' +
    row("Name", escapeHtml(safeName)) +
    row("E-Mail", emailValue) +
    row("Abholung", escapeHtml(valueOrDash(order.pickup))) +
    row("Wunschtermin", escapeHtml(valueOrDash(order.date))) +
    row("Preis", escapeHtml(getOrderPriceText_(order))) +
    row("Wünsche", escapeHtml(valueOrDash(order.note))) +
    '</table><div style="margin-top:18px;font-size:11px;color:#aaa;">ID: ' + escapeHtml(valueOrDash(order.orderId)) +
    '</div></td></tr></table></body></html>';
  const text = [
    "NEUE HONIGBESTELLUNG",
    safeOrderNumber + " · " + timestamp,
    "",
    ...items.map(function(item) { return "- " + item; }),
    "",
    "Name:          " + valueOrDash(order.name),
    "E-Mail:        " + valueOrDash(order.email),
    "Abholung:      " + valueOrDash(order.pickup),
    "Wunschtermin:  " + valueOrDash(order.date),
    "Preis:         " + getOrderPriceText_(order),
    "Wünsche:       " + valueOrDash(order.note),
    "",
    "ID: " + valueOrDash(order.orderId)
  ].join("\n");

  MailApp.sendEmail({
    to: ORDER_NOTIFICATION_EMAIL,
    name: ORDER_SENDER_NAME,
    replyTo: order.email,
    subject: subject,
    body: text,
    htmlBody: html
  });
}

function sendCustomerOrderConfirmation_(order) {
  if (!order || !order.email) return;
  const properties = PropertiesService.getScriptProperties();
  const replyToProperty = properties.getProperty("CUSTOMER_REPLY_TO");
  const replyTo = String(replyToProperty == null ? "landimkereidiezinger.bestellung@gmx.de" : replyToProperty).trim();
  const contactsProperty = properties.getProperty("CUSTOMER_CONTACTS");
  const contacts = String(contactsProperty == null ? "gerhard-diez@gmx.de, lukasdiezinger3@gmail.com" : contactsProperty).split(",")
    .map(function(address) { return address.trim(); })
    .filter(function(address) { return address; });
  const escapeHtml = function(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  };
  const valueOrDash = function(value) {
    const text = String(value == null ? "" : value).trim();
    return text ? text : "–";
  };
  const items = String(order.itemSummary || "").split(/,\s*(?=\d+x\s)/).filter(function(item) { return item; });
  const row = function(label, value) {
    return '<tr><td style="padding:7px 12px 7px 0;color:#888;vertical-align:top;">' + label +
      '</td><td style="padding:7px 0;color:#222;vertical-align:top;white-space:pre-wrap;">' + value + '</td></tr>';
  };
  const mailLink = function(address) {
    const safeAddress = escapeHtml(address);
    return '<a href="mailto:' + safeAddress + '" style="color:#822045;">' + safeAddress + '</a>';
  };
  const safeName = valueOrDash(order.name);
  const safeOrderNumber = valueOrDash(order.customerOrderNumber);
  const date = order.date || "Kein Wunschtermin angegeben";
  const contactHtml = (replyTo ? 'Bei Fragen schreib an ' + mailLink(replyTo) + '. ' : "") +
    (contacts.length ? 'Eine Antwort erhältst du persönlich von ' + contacts.map(mailLink).join(" oder ") + "." : "");
  const html = '<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:20px 12px;background:#f5f5f5;font-family:-apple-system,\'Segoe UI\',Helvetica,Arial,sans-serif;font-size:15px;color:#222;line-height:1.5;">' +
    '<table role="presentation" width="480" cellpadding="0" cellspacing="0" style="width:100%;max-width:480px;margin:0 auto;background:#fff;border-top:4px solid #822045;border-radius:6px;">' +
    '<tr><td style="padding:22px 24px;">' +
    '<div style="font-size:18px;font-weight:bold;color:#822045;">Deine Bestellung ist eingegangen</div>' +
    '<p style="margin:10px 0 16px;">Servus ' + escapeHtml(safeName) + ', danke für deine Bestellung. Wir haben alles erhalten und machen uns bald an die Vorbereitung.</p>' +
    '<div style="display:inline-block;padding:6px 12px;background:#FFE337;color:#822045;border-radius:6px;font-family:\'Courier New\',monospace;font-weight:bold;">' + escapeHtml(safeOrderNumber) + '</div>' +
    '<div style="margin:18px 0 6px;padding-top:14px;border-top:1px solid #eee;font-weight:bold;">' +
    (items.length ? items.map(function(item) { return escapeHtml(item); }).join("<br>") : "–") + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;font-size:14px;">' +
    row("Abholung", escapeHtml(valueOrDash(order.pickup))) +
    row("Wunschtermin", escapeHtml(valueOrDash(date))) +
    row("Preis", escapeHtml(getOrderPriceText_(order))) +
    row("Status", escapeHtml(ORDER_STATUSES[0])) +
    row("Wünsche", escapeHtml(valueOrDash(order.note))) +
    '</table><p style="margin:18px 0 0;">Sobald deine Bestellung abholbereit ist, bekommst du eine Nachricht von uns.</p>' +
    '<p style="margin:18px 0 0;">Viele Grüße<br>Landimkerei Diezinger</p>' +
    '<div style="margin-top:18px;padding-top:12px;border-top:1px solid #eee;font-size:12px;color:#999;">' +
    'Diese E-Mail wurde automatisch erstellt, bitte antworte nicht darauf. ' + contactHtml +
    '</div></td></tr></table></body></html>';
  const text = [
    "Servus " + (order.name || "") + ",",
    "",
    "danke für deine Bestellung. Wir haben alles erhalten und machen uns bald an die Vorbereitung.",
    "",
    "Bestellnummer: " + safeOrderNumber,
    ...items.map(function(item) { return "Bestellte Artikel: " + item; }),
    "Abholung: " + valueOrDash(order.pickup),
    "Wunschtermin: " + date,
    "Preis: " + getOrderPriceText_(order),
    "Status: " + ORDER_STATUSES[0],
    "Wünsche: " + valueOrDash(order.note),
    "",
    "Sobald deine Bestellung abholbereit ist, bekommst du eine Nachricht von uns.",
    "",
    "Viele Grüße",
    "Landimkerei Diezinger",
    "",
    "Diese E-Mail wurde automatisch erstellt, bitte antworte nicht darauf.",
    replyTo ? "Bei Fragen schreib an " + replyTo + "." : "",
    contacts.length ? "Eine Antwort erhältst du persönlich von " + contacts.join(" oder ") + "." : ""
  ].join("\n");
  MailApp.sendEmail({
    to: order.email,
    name: "Landimkerei Diezinger",
    subject: "Deine Bestellung ist eingegangen – " + (order.customerOrderNumber || "") + " – Landimkerei Diezinger",
    body: text,
    htmlBody: html
  });
}

function sendCustomerOrderStatusEmail_(order, status) {
  if (!order || !order.email) return;
  const properties = PropertiesService.getScriptProperties();
  const replyToProperty = properties.getProperty("CUSTOMER_REPLY_TO");
  const replyTo = String(replyToProperty == null ? "landimkereidiezinger.bestellung@gmx.de" : replyToProperty).trim();
  const contactsProperty = properties.getProperty("CUSTOMER_CONTACTS");
  const contacts = String(contactsProperty == null ? "gerhard-diez@gmx.de, lukasdiezinger3@gmail.com" : contactsProperty).split(",")
    .map(function(address) { return address.trim(); })
    .filter(function(address) { return address; });
  const messages = {
    "In Bearbeitung": {
      subject: "Deine Bestellung wird gerade vorbereitet",
      title: "Deine Bestellung wird vorbereitet",
      html: "Wir bereiten deine Bestellung gerade für dich vor. Sobald sie abholbereit ist, bekommst du direkt eine Nachricht von uns. Danke dir für deine Geduld!",
      text: "Wir bereiten deine Bestellung gerade für dich vor. Sobald sie abholbereit ist, bekommst du direkt eine Nachricht von uns. Danke dir für deine Geduld!"
    },
    "Verzögert": {
      subject: "Deine Bestellung ist noch nicht abholbereit",
      title: "Deine Bestellung ist noch nicht abholbereit",
      html: "Es tut uns leid, aber deine Bestellung ist noch nicht bereit und verzögert sich etwas. Leider können wir dir momentan nicht sagen, wie lange es noch dauert. Wir geben unser Bestes, damit du sie so schnell wie möglich abholen kannst, und melden uns direkt bei dir, sobald sie bereit ist. Danke dir für deine Geduld und dein Verständnis!",
      text: "Es tut uns leid, aber deine Bestellung ist noch nicht bereit und verzögert sich etwas. Leider können wir dir momentan nicht sagen, wie lange es noch dauert. Wir geben unser Bestes, damit du sie so schnell wie möglich abholen kannst, und melden uns direkt bei dir, sobald sie bereit ist. Danke dir für deine Geduld und dein Verständnis!"
    },
    "Bereit zur Abholung": {
      subject: "Deine Bestellung ist abholbereit",
      title: "Deine Bestellung ist abholbereit",
      html: "Deine Bestellung ist jetzt abholbereit. Du kannst sie wie vereinbart abholen. Wir freuen uns auf dich!",
      text: "Deine Bestellung ist jetzt abholbereit. Du kannst sie wie vereinbart abholen. Wir freuen uns auf dich!"
    },
    "Abgeholt": {
      subject: "Deine Bestellung wurde abgeholt",
      title: "Deine Bestellung wurde abgeholt",
      html: "Deine Bestellung wurde als abgeholt vermerkt. Wir wünschen dir viel Freude mit dem Honig und den anderen Produkten. Vielen Dank für deine Bestellung! Wir freuen uns, wenn du bald wieder bei uns vorbeischaust und etwas bestellst.",
      text: "Deine Bestellung wurde als abgeholt vermerkt. Wir wünschen dir viel Freude mit dem Honig und den anderen Produkten. Vielen Dank für deine Bestellung! Wir freuen uns, wenn du bald wieder bei uns vorbeischaust und etwas bestellst."
    }
  };
  const message = messages[status];
  if (!message) return;
  const escapeHtml = function(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  };
  const valueOrDash = function(value) {
    const text = String(value == null ? "" : value).trim();
    return text ? text : "–";
  };
  const safeOrderNumber = valueOrDash(order.customerOrderNumber);
  const safeName = valueOrDash(order.name);
  const items = String(order.itemSummary || "").split(/,\s*(?=\d+x\s)/).filter(function(item) { return item; });
  const row = function(label, value) {
    return '<tr><td style="padding:7px 12px 7px 0;color:#888;vertical-align:top;">' + label +
      '</td><td style="padding:7px 0;color:#222;vertical-align:top;white-space:pre-wrap;">' + value + '</td></tr>';
  };
  const mailLink = function(address) {
    const safeAddress = escapeHtml(address);
    return '<a href="mailto:' + safeAddress + '" style="color:#822045;">' + safeAddress + '</a>';
  };
  const contactHtml = (replyTo ? 'Bei Fragen schreib an ' + mailLink(replyTo) + '. ' : "") +
    (contacts.length ? 'Eine Antwort erhältst du persönlich von ' + contacts.map(mailLink).join(" oder ") + "." : "");
  const contactText = [
    replyTo ? "Bei Fragen schreib an " + replyTo + "." : "",
    contacts.length ? "Eine Antwort erhältst du persönlich von " + contacts.join(" oder ") + "." : ""
  ].filter(Boolean).join(" ");
  const privacyNotice = status === "Abgeholt"
    ? "Deine persönlichen Angaben wie Name, E-Mail-Adresse, Abholwunsch und Wünsche wurden aus unserer Bestellübersicht gelöscht."
    : "";
  const itemList = items.length ? items.map(function(item) { return escapeHtml(item); }).join("<br>") : "–";
  const html = '<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:20px 12px;background:#f5f5f5;font-family:-apple-system,\'Segoe UI\',Helvetica,Arial,sans-serif;font-size:15px;color:#222;line-height:1.5;">' +
    '<table role="presentation" width="480" cellpadding="0" cellspacing="0" style="width:100%;max-width:480px;margin:0 auto;background:#fff;border-top:4px solid #822045;border-radius:6px;">' +
    '<tr><td style="padding:22px 24px;">' +
    '<div style="font-size:18px;font-weight:bold;color:#822045;">' + escapeHtml(message.title) + '</div>' +
    '<p style="margin:10px 0 16px;">Hallo ' + escapeHtml(safeName) + ',<br>' + message.html + '</p>' +
    '<div style="display:inline-block;padding:6px 12px;background:#FFE337;color:#822045;border-radius:6px;font-family:\'Courier New\',monospace;font-weight:bold;">' + escapeHtml(safeOrderNumber) + '</div>' +
    '<div style="margin:18px 0 6px;padding-top:14px;border-top:1px solid #eee;font-weight:bold;">' + itemList + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;font-size:14px;">' +
    row("Abholung", escapeHtml(valueOrDash(order.pickup))) +
    row("Wunschtermin", escapeHtml(valueOrDash(order.date))) +
    row("Preis", escapeHtml(getOrderPriceText_(order))) +
    row("Status", escapeHtml(status)) +
    '</table><p style="margin:18px 0 0;">Viele Grüße<br>Landimkerei Diezinger</p>' +
    '<div style="margin-top:18px;padding-top:12px;border-top:1px solid #eee;font-size:12px;color:#999;">' +
    'Diese E-Mail wurde automatisch erstellt, bitte antworte nicht darauf. ' + contactHtml +
    (privacyNotice ? '<div style="margin-top:8px;">' + escapeHtml(privacyNotice) + '</div>' : "") +
    '</div></td></tr></table></body></html>';
  const text = [
    "Hallo " + (order.name || "") + ",",
    "",
    message.text,
    "",
    "Bestellnummer: " + safeOrderNumber,
    ...items.map(function(item) { return "Bestellte Artikel: " + item; }),
    "Abholung: " + valueOrDash(order.pickup),
    "Wunschtermin: " + valueOrDash(order.date),
    "Preis: " + getOrderPriceText_(order),
    "Status: " + status,
    "",
    "Viele Grüße",
    "Landimkerei Diezinger",
    "",
    "Diese E-Mail wurde automatisch erstellt, bitte antworte nicht darauf.",
    contactText,
    privacyNotice
  ].join("\n");
  MailApp.sendEmail({
    to: order.email,
    name: "Landimkerei Diezinger",
    subject: message.subject + " " + (order.customerOrderNumber || "") + " – Landimkerei Diezinger",
    body: text,
    htmlBody: html
  });
}

function authorizeOrderEmail() {
  return MailApp.getRemainingDailyQuota();
}

function EmailStatusSend() {
  const testEmail = String(PropertiesService.getScriptProperties().getProperty("EMAIL_STATUSSEND_TEST") || "").trim();
  if (!testEmail) {
    throw new Error("Lege unter Projekteinstellungen > Script-Eigenschaften die Adresse EMAIL_STATUSSEND_TEST fest.");
  }

  const testOrder = {
    name: "E-Mail-Test",
    email: testEmail,
    itemSummary: "1x Cremiger Honig Frühling (250 g)",
    customerOrderNumber: "TEST-STATUS",
    pickup: "Persönlich abholen (Mo-Fr 17-19 Uhr, Sa 10-19 Uhr)",
    date: "Kein Wunschtermin angegeben",
    note: "",
    totalPrice: 3,
    priceNote: ""
  };
  sendCustomerOrderConfirmation_(testOrder);
  ORDER_STATUSES.slice(1).forEach(function(status) {
    sendCustomerOrderStatusEmail_(testOrder, status);
  });
  console.log("Alle vier Statusmails wurden ausschließlich an " + testEmail + " gesendet.");
  return testEmail;
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

function readProductSizes_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("Das Bestandsblatt ist nicht eingerichtet.");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return rows.reduce(function(sizes, row) {
    const id = String(row[0] || "").trim();
    const size = String(row[2] || "").trim();
    if (id && id !== "HNG" && id !== "KZN" && id !== "K3D" && size) sizes[id] = size;
    return sizes;
  }, {});
}

function readPrices_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("Das Bestandsblatt ist nicht eingerichtet.");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return rows.reduce(function(prices, row) {
    const id = String(row[0]).trim();
    if (!id || id === "HNG" || id === "KZN" || id === "K3D") return prices;
    const rawPrice = row[6];
    const price = rawPrice === "" || rawPrice == null
      ? null
      : Number(String(rawPrice).replace(/\s|€/g, "").replace(",", "."));
    prices[id] = Number.isFinite(price) && price >= 0 ? price : null;
    return prices;
  }, {});
}

function formatEuro_(amount) {
  return Number(amount).toFixed(2).replace(".", ",") + " €";
}

function getOrderPriceText_(order) {
  const amount = Number(order.totalPrice);
  const hasAmount = order.totalPrice !== "" && order.totalPrice != null && Number.isFinite(amount);
  const priceNote = String(order.priceNote || "").trim();
  if (priceNote) {
    return hasAmount && amount > 0
      ? "Zwischensumme: " + formatEuro_(amount) + ". " + priceNote
      : priceNote;
  }
  return hasAmount ? "Gesamtpreis: " + formatEuro_(amount) : "Preis in dieser Bestellung nicht hinterlegt.";
}

function initializeProductPrices_(sheet) {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty("PRODUCT_PRICES_INITIALIZED") === "true") return;
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    rows.forEach(function(row, index) {
      const id = String(row[0]).trim();
      if (Object.prototype.hasOwnProperty.call(INITIAL_PRICES, id) && (row[6] === "" || row[6] == null)) {
        sheet.getRange(index + 2, 7).setValue(INITIAL_PRICES[id]);
      }
    });
  }
  properties.setProperty("PRODUCT_PRICES_INITIALIZED", "true");
}

function readUnavailable_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("Das Bestandsblatt ist nicht eingerichtet.");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return rows.reduce(function(unavailable, row) {
    const id = String(row[0]);
    if (id && id !== "HNG" && id !== "KZN" && id !== "K3D") unavailable[id] = normalizeProductAvailability_(row[4]) === "Nicht verfügbar";
    return unavailable;
  }, {});
}

function readProductStatuses_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("Das Bestandsblatt ist nicht eingerichtet.");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return rows.reduce(function(statuses, row) {
    const id = String(row[0] || "").trim();
    if (id && id !== "HNG" && id !== "KZN" && id !== "K3D") statuses[id] = normalizeProductAvailability_(row[4]);
    return statuses;
  }, {});
}

function normalizeProductAvailability_(value) {
  if (value === true || String(value).trim().toLowerCase() === "true") return "Nicht verfügbar";
  if (value === false || value == null || String(value).trim() === "") return "Normal";
  const status = String(value).trim();
  return PRODUCT_AVAILABILITY_STATUSES.indexOf(status) !== -1 ? status : "Normal";
}

function readUnavailableReasons_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("Das Bestandsblatt ist nicht eingerichtet.");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return rows.reduce(function(reasons, row) {
    const id = String(row[0]);
    if (id && id !== "HNG" && id !== "KZN" && id !== "K3D") reasons[id] = String(row[5] || "").trim();
    return reasons;
  }, {});
}

function setupProductAvailabilityDropdowns_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(PRODUCT_AVAILABILITY_STATUSES, true).build();
  rows.forEach(function(row, index) {
    const id = String(row[0]).trim();
    if (!id || id === "HNG" || id === "KZN" || id === "K3D") return;
    const cell = sheet.getRange(index + 2, 5);
    const status = normalizeProductAvailability_(row[4]);
    cell.setDataValidation(statusRule);
    if (row[4] !== status) cell.setValue(status);
  });
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== INVENTORY_SHEET || e.range.getColumn() !== 5 || e.range.getNumColumns() !== 1) return;
  const startRow = Math.max(2, e.range.getRow());
  const endRow = e.range.getLastRow();
  for (let row = startRow; row <= endRow; row++) {
    const id = String(sheet.getRange(row, 1).getValue() || "").trim();
    if (!id || id === "HNG" || id === "KZN" || id === "K3D") continue;
    updateProductReasonDropdown_(sheet, row, normalizeProductAvailability_(sheet.getRange(row, 5).getValue()));
  }
}

function onOrderStatusEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== ORDERS_SHEET || range.getColumn() !== 10 || range.getNumColumns() !== 1 || range.getNumRows() !== 1) return;
  const status = String(e.value == null ? range.getValue() : e.value).trim();
  const previousStatus = String(e.oldValue || "").trim();
  if (!status || status === previousStatus || status === ORDER_STATUSES[0] || ORDER_STATUSES.indexOf(status) === -1) return;

  const row = sheet.getRange(range.getRow(), 1, 1, 14).getValues()[0];
  if (String(row[9] || "").trim() !== status) return;
  const order = {
    orderId: String(row[0] || ""),
    timestamp: row[1],
    name: String(row[2] || ""),
    email: String(row[3] || ""),
    itemSummary: String(row[4] || ""),
    pickup: String(row[5] || ""),
    date: String(row[6] || ""),
    customerOrderNumber: String(row[8] || ""),
    totalPrice: row[10],
    priceNote: String(row[11] || ""),
    pdfFileId: String(row[13] || "")
  };

  if (status === "Bereit zur Abholung") {
    const pdfFile = order.pdfFileId
      ? DriveApp.getFileById(order.pdfFileId)
      : createOrderSlipPdf_(order);
    sheet.getRange(range.getRow(), 13).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText("PDF öffnen / drucken").setLinkUrl(pdfFile.getUrl()).build()
    );
    sheet.getRange(range.getRow(), 14).setValue(pdfFile.getId());
  } else if (status === "Abgeholt") {
    archiveCompletedOrder_(order);
    trashOrderSlipPdf_(order.pdfFileId);
    sheet.deleteRow(range.getRow());
    sendCustomerOrderStatusEmail_(order, status);
    return;
  } else if (order.pdfFileId) {
    trashOrderSlipPdf_(order.pdfFileId);
    sheet.getRange(range.getRow(), 13, 1, 2).clearContent();
  }

  sendCustomerOrderStatusEmail_(order, status);
}

function setupProductReasonDropdowns_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const reasonMigrations = {
    "Ausverkauft": "Nicht verfügbar",
    "Bestand aufgebraucht": "Vorübergehend nicht vorrätig",
    "Vorübergehend nicht lieferbar": "Vorübergehend nicht vorrätig",
    "Saisonbedingt nicht verfügbar": "Saisonpause"
  };
  rows.forEach(function(row, index) {
    const id = String(row[0]).trim();
    if (!id || id === "HNG" || id === "KZN" || id === "K3D") return;
    const reasonCell = sheet.getRange(index + 2, 6);
    const oldReason = String(row[5] || "").trim();
    if (Object.prototype.hasOwnProperty.call(reasonMigrations, oldReason)) reasonCell.setValue(reasonMigrations[oldReason]);
    updateProductReasonDropdown_(sheet, index + 2, normalizeProductAvailability_(row[4]));
  });
}

function updateProductReasonDropdown_(sheet, row, status) {
  const reasonCell = sheet.getRange(row, 6);
  const reasons = status === "Schnell zugreifen"
    ? QUICK_ACCESS_REASONS
    : status === "Nicht verfügbar" ? UNAVAILABLE_REASONS : null;
  if (!reasons) {
    reasonCell.clearDataValidations();
    return;
  }
  const currentReason = String(reasonCell.getValue() || "").trim();
  if (currentReason && reasons.indexOf(currentReason) === -1) reasonCell.clearContent();
  const reasonRule = SpreadsheetApp.newDataValidation().requireValueInList(reasons, true).build();
  reasonCell.setDataValidation(reasonRule);
}

function migrateInventoryColumns_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 6) return;
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(value) { return String(value).trim(); });
  if (headers[5] !== "Preis (€)") return;

  const hasUnavailableReasons = headers[6] === "Grund nicht verfügbar" || headers[6] === "Grund";
  const lastRow = sheet.getLastRow();
  const existingReasons = hasUnavailableReasons && lastRow > 1
    ? sheet.getRange(2, 7, lastRow - 1, 1).getValues()
    : [];
  sheet.insertColumnAfter(5);
  if (hasUnavailableReasons && lastRow > 1) {
    sheet.getRange(2, 6, lastRow - 1, 1).setValues(existingReasons);
    if (lastColumn === 7) sheet.deleteColumn(8);
  }
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
    const pickup = String(payload.pickup || "");
    if (String(payload.honeypot || "").trim()) throw new Error("Die Bestellung konnte nicht geprüft werden.");
    if (!orderId || !name || !email || !requested.length) {
      throw new Error("Die Bestellung enthält nicht alle erforderlichen Angaben.");
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("Bitte gib eine gültige E-Mail-Adresse an.");

    const orderRows = orders.getLastRow() > 1
      ? orders.getRange(2, 1, orders.getLastRow() - 1, 10).getValues()
      : [];
    const previous = orderRows.find(function(row) { return String(row[0]) === orderId; });
    if (previous) return {
      ok: true,
      duplicate: true,
      customerOrderNumber: String(previous[8] || ""),
      stock: readInventory_(),
      unavailable: readUnavailable_(),
      lowStockThreshold: LOW_STOCK_THRESHOLD,
      lowStockAlerts: []
    };

    const dataRange = inventory.getRange(2, 1, inventory.getLastRow() - 1, 4);
    const rows = dataRange.getValues();
    const unavailableRows = inventory.getRange(2, 5, inventory.getLastRow() - 1, 1).getValues();
    const rowById = {};
    rows.forEach(function(row, index) {
      const id = String(row[0]);
      if (!id || id === "HNG" || id === "KZN" || id === "K3D") return;
      const stock = Number(row[3]);
      if (!Number.isInteger(stock) || stock < 0) throw new Error("Im Bestandsblatt steht eine ungültige Stückzahl.");
      rowById[id] = { row: index + 2, name: String(row[1]), size: String(row[2]), stock: stock, unavailable: normalizeProductAvailability_(unavailableRows[index][0]) === "Nicht verfügbar" };
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

    const prices = readPrices_();
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
      return { id: id, quantity: quantity, product: product, unitPrice: prices[id] };
    });

    let amountCents = pickup.indexOf("Abholschrank") !== -1 ? 100 : 0;
    const priceNotes = [];
    purchases.forEach(function(item) {
      if (item.id === "wh1") {
        priceNotes.push(item.unitPrice == null
          ? "Wabenhonigpreis fehlt im Bestandssheet."
          : "Wabenhonig wird nach Gewicht berechnet (" + formatEuro_(item.unitPrice) + " pro 100 g). Der genaue Preis steht erst mit dem Gewicht fest.");
      } else if (item.unitPrice == null) {
        priceNotes.push("Für " + item.product.name + " fehlt ein Preis im Bestandssheet.");
      } else {
        amountCents += Math.round(item.unitPrice * 100) * item.quantity;
      }
    });
    const totalPrice = amountCents / 100;
    const priceNote = priceNotes.join(" ");

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
    const existingNumbers = new Set(orderRows.map(function(row) { return String(row[8] || ""); }));
    const customerOrderNumber = generateCustomerOrderNumber_(existingNumbers);
    dataRange.setValues(nextRows);
    try {
      orders.appendRow([
        orderId,
        createdAt,
        name,
        email,
        itemSummary,
        pickup,
        String(payload.date || ""),
        String(payload.note || ""),
        customerOrderNumber,
        ORDER_STATUSES[0],
        totalPrice,
        priceNote
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
      customerOrderNumber: customerOrderNumber,
      unavailable: readUnavailable_(),
      lowStockThreshold: LOW_STOCK_THRESHOLD,
      lowStockAlerts: lowStockAlerts,
      orderDetails: {
        orderId: orderId,
        timestamp: createdAt,
        name: name,
        email: email,
        itemSummary: itemSummary,
        customerOrderNumber: customerOrderNumber,
        pickup: pickup,
        date: String(payload.date || ""),
        note: String(payload.note || ""),
        totalPrice: totalPrice,
        priceNote: priceNote
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function setupOrderTracking_(orders) {
  if (orders.getMaxColumns() < 14) orders.insertColumnsAfter(orders.getMaxColumns(), 14 - orders.getMaxColumns());
  orders.getRange(1, 9, 1, 6).setValues([["Bestellnummer", "Status", "Betrag (€)", "Preishinweis", "Druck-PDF", "PDF-Datei-ID"]]);
  orders.getRange(2, 11, Math.max(1, orders.getMaxRows() - 1), 1).setNumberFormat('0.00 "€"');
  orders.hideColumns(14);
  const lastRow = orders.getLastRow();
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(ORDER_STATUSES, true).build();
  if (lastRow < 2) return;

  const rows = orders.getRange(2, 1, lastRow - 1, 10).getValues();
  const existingNumbers = new Set(rows.map(function(row) { return String(row[8] || ""); }).filter(Boolean));
  const trackingValues = rows.map(function(row) {
    if (!String(row[0] || "")) return [row[8], row[9]];
    let customerOrderNumber = String(row[8] || "");
    if (!customerOrderNumber) {
      customerOrderNumber = generateCustomerOrderNumber_(existingNumbers);
      existingNumbers.add(customerOrderNumber);
    }
    return [customerOrderNumber, String(row[9] || ORDER_STATUSES[0])];
  });
  orders.getRange(2, 9, trackingValues.length, 2).setValues(trackingValues);
  orders.getRange(2, 10, Math.max(1, orders.getMaxRows() - 1), 1).setDataValidation(statusRule);
}

function setupOrderArchive_(spreadsheet) {
  let archive = spreadsheet.getSheetByName("Bestellung-Archiv");
  if (!archive) archive = spreadsheet.insertSheet("Bestellung-Archiv");
  archive.getRange(1, 1, 1, 8).setValues([[
    "Bestellnummer", "Produkte", "Abholart", "Wünsche", "Bestellt am", "Abgeholt am", "Betrag (€)", "Preishinweis"
  ]]);
  archive.setFrozenRows(1);
  archive.getRange(2, 5, Math.max(1, archive.getMaxRows() - 1), 2).setNumberFormat("dd.MM.yyyy HH:mm");
  if (archive.getMaxColumns() >= 7) archive.getRange(2, 7, Math.max(1, archive.getMaxRows() - 1), 1).setNumberFormat('0.00 "€"');
  archive.autoResizeColumns(1, 8);
}

function archiveCompletedOrder_(order) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const archive = spreadsheet.getSheetByName("Bestellung-Archiv") || spreadsheet.insertSheet("Bestellung-Archiv");
  if (archive.getLastRow() === 0) setupOrderArchive_(spreadsheet);
  if (!order.customerOrderNumber) throw new Error("Die Bestellung hat keine Bestellnummer und konnte nicht archiviert werden.");
  if (archive.getLastRow() > 1) {
    const archivedNumbers = archive.getRange(2, 1, archive.getLastRow() - 1, 1).getValues();
    if (archivedNumbers.some(function(row) { return String(row[0]) === order.customerOrderNumber; })) return;
  }
  archive.appendRow([
    order.customerOrderNumber,
    order.itemSummary,
    order.pickup,
    order.note,
    order.timestamp,
    new Date(),
    order.totalPrice,
    order.priceNote
  ]);
}

function getOrderSlipFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const folderId = properties.getProperty("ORDER_SLIP_FOLDER_ID");
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (error) {
      properties.deleteProperty("ORDER_SLIP_FOLDER_ID");
    }
  }
  const folder = DriveApp.createFolder("Bestellzettel (vorübergehend)");
  properties.setProperty("ORDER_SLIP_FOLDER_ID", folder.getId());
  return folder;
}

function createOrderSlipPdf_(order) {
  const safeNumber = String(order.customerOrderNumber || "Bestellung").replace(/[^A-Za-z0-9_-]/g, "_");
  const document = DocumentApp.create("Bestellübersicht " + safeNumber);
  const documentId = document.getId();
  try {
    const body = document.getBody();
    body.clear();
    body.setMarginTop(42).setMarginBottom(42).setMarginLeft(48).setMarginRight(48);
    const brand = body.appendParagraph("LANDIMKEREI DIEZINGER");
    brand.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    brand.editAsText().setBold(true).setFontSize(16).setForegroundColor("#822045");
    const heading = body.appendParagraph("Bestellübersicht");
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    heading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("Bestellnummer: " + String(order.customerOrderNumber || "–")).editAsText().setBold(true);
    const orderedAt = order.timestamp instanceof Date
      ? Utilities.formatDate(order.timestamp, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm")
      : String(order.timestamp || "–");
    body.appendParagraph("Bestellt am: " + orderedAt);
    body.appendParagraph("Name: " + String(order.name || "–"));
    body.appendParagraph("Abholart: " + String(order.pickup || "–"));
    body.appendParagraph("Wunschtermin: " + String(order.date || "Kein Wunschtermin angegeben"));
    body.appendParagraph("Artikel").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    String(order.itemSummary || "–").split(/,\s*(?=\d+x\s)/).filter(function(item) { return item; }).forEach(function(item) {
      body.appendListItem(item);
    });
    body.appendParagraph("Wünsche: " + String(order.note || "Keine Wünsche angegeben"));
    body.appendParagraph("Preis: " + getOrderPriceText_(order));
    const disclaimer = body.appendParagraph("Dieses Blatt ist keine Rechnung, sondern eine Zusammenfassung der Bestellung.");
    disclaimer.editAsText().setBold(true);
    document.saveAndClose();
    const pdf = DriveApp.getFileById(documentId).getAs(MimeType.PDF).setName("Bestellübersicht " + safeNumber + ".pdf");
    const file = getOrderSlipFolder_().createFile(pdf);
    DriveApp.getFileById(documentId).setTrashed(true);
    return file;
  } catch (error) {
    try {
      DriveApp.getFileById(documentId).setTrashed(true);
    } catch (cleanupError) {
      console.error("Temporäres Bestelldokument konnte nicht gelöscht werden: " + cleanupError.message);
    }
    throw error;
  }
}

function trashOrderSlipPdf_(fileId) {
  if (!fileId) return;
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (error) {
    console.warn("Bestell-PDF war bereits gelöscht oder konnte nicht gefunden werden: " + error.message);
  }
}

function setupOrderStatusEmailTrigger_(spreadsheet) {
  const triggerExists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === "onOrderStatusEdit" && trigger.getEventType() === ScriptApp.EventType.ON_EDIT;
  });
  if (!triggerExists) {
    ScriptApp.newTrigger("onOrderStatusEdit").forSpreadsheet(spreadsheet).onEdit().create();
  }
}

function generateCustomerOrderNumber_(existingNumbers) {
  let customerOrderNumber;
  do {
    customerOrderNumber = "H-" + Utilities.getUuid().replace(/-/g, "").substring(0, 10).toUpperCase();
  } while (existingNumbers.has(customerOrderNumber));
  return customerOrderNumber;
}

