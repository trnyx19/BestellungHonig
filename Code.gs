const LOW_STOCK_THRESHOLD = 6;
const LOW_STOCK_TELEGRAM_THRESHOLD = 3;
const INVENTORY_SHEET = "Bestand";
const ORDERS_SHEET = "Bestellungen";
const ORDER_NOTIFICATION_EMAIL = "gerhard-diez@gmx.de, lukasdiezinger3@gmail.com";
const ORDER_SENDER_NAME = "Neue Bestellung";
const ORDER_EMAIL_SUBJECT = "[HONIG-BESTELLUNG] Neue Bestellung";
const MAX_ORDER_QUANTITIES = { arb3d: 3, koe3d: 3, wei3d: 3, volk3d: 5, weivol3d: 5 };
const ORDER_STATUSES = ["Eingegangen", "In Bearbeitung", "Bereit zur Abholung", "Abgeholt"];
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
  ["fehldruck3d", "Bienchen Fehldruck", "Einzeln", 0],
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
  if (!isNewInventory) migrateInventoryColumns_(inventory);
  if (isNewInventory) {
    inventory.getRange(1, 1, 1, 7).setValues([["ID", "Sorte", "Größe", "Bestand", "Nicht verfügbar", "Grund nicht verfügbar", "Preis (€)"]]);
    inventory.setFrozenRows(1);
  }
  inventory.getRange(1, 5).setValue("Nicht verfügbar");
  inventory.getRange(1, 6).setValue("Grund nicht verfügbar");
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
  inventory.getRange(1, 6).setNote("Wird auf der Bestellseite angezeigt, wenn „Nicht verfügbar“ aktiviert ist.");
  inventory.getRange(1, 7).setNote("Wabenhonig: Preis je 100 g; bei allen anderen Artikeln Stückpreis.");
  if (inventory.getLastRow() > 1) {
    inventory.getRange(2, 7, inventory.getLastRow() - 1, 1).setNumberFormat('0.00 "€"');
  }
  setupUnavailableCheckboxes_(inventory);
  setupUnavailableReasonDropdowns_(inventory);
  inventory.autoResizeColumns(1, 7);

  let orders = spreadsheet.getSheetByName(ORDERS_SHEET);
  if (!orders) orders = spreadsheet.insertSheet(ORDERS_SHEET);
  if (orders.getLastRow() === 0) {
    orders.getRange(1, 1, 1, 10).setValues([[
      "Bestell-ID", "Zeitpunkt", "Name", "E-Mail", "Artikel", "Abholung", "Wunschtermin", "Wünsche", "Bestellnummer", "Status"
    ]]);
    orders.setFrozenRows(1);
  }
  setupOrderTracking_(orders);
  orders.autoResizeColumns(1, 10);
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
    if (payload.action === "lookupOrder") {
      return createOrderLookupResponse_(lookupOrder_(payload));
    }
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
    if (payload.action === "lookupOrder") {
      return createOrderLookupResponse_({
        type: "orderLookupResult",
        requestId: String(payload.requestId || ""),
        ok: false,
        message: "Bestellung nicht gefunden. Bitte prüfe Bestellnummer und E-Mail-Adresse."
      });
    }
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
  const replyTo = String(properties.getProperty("CUSTOMER_REPLY_TO") || "").trim();
  const contacts = String(properties.getProperty("CUSTOMER_CONTACTS") || "").split(",")
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
  const statusLookup = "Status abrufen: Auf der Bestellseite ganz nach unten scrollen, „Bestellung abrufen“ anklicken und Bestellnummer sowie E-Mail-Adresse eingeben.";
  const html = '<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:20px 12px;background:#f5f5f5;font-family:-apple-system,\'Segoe UI\',Helvetica,Arial,sans-serif;font-size:15px;color:#222;line-height:1.5;">' +
    '<table role="presentation" width="480" cellpadding="0" cellspacing="0" style="width:100%;max-width:480px;margin:0 auto;background:#fff;border-top:4px solid #822045;border-radius:6px;">' +
    '<tr><td style="padding:22px 24px;">' +
    '<div style="font-size:18px;font-weight:bold;color:#822045;">Danke für deine Bestellung</div>' +
    '<p style="margin:10px 0 16px;">Servus ' + escapeHtml(safeName) + ', wir haben deine Bestellung erhalten.</p>' +
    '<div style="display:inline-block;padding:6px 12px;background:#FFE337;color:#822045;border-radius:6px;font-family:\'Courier New\',monospace;font-weight:bold;">' + escapeHtml(safeOrderNumber) + '</div>' +
    '<div style="margin:18px 0 6px;padding-top:14px;border-top:1px solid #eee;font-weight:bold;">' +
    (items.length ? items.map(function(item) { return escapeHtml(item); }).join("<br>") : "–") + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;font-size:14px;">' +
    row("Abholung", escapeHtml(valueOrDash(order.pickup))) +
    row("Wunschtermin", escapeHtml(valueOrDash(date))) +
    row("Status", escapeHtml(ORDER_STATUSES[0])) +
    row("Wünsche", escapeHtml(valueOrDash(order.note))) +
    '</table><p style="margin:18px 0 0;">Wir melden uns, sobald deine Bestellung zur Abholung bereit ist.</p>' +
    '<p style="margin:14px 0 0;padding:12px 14px;background:#fff8cc;border-radius:6px;font-size:14px;"><b>Status abrufen:</b> Auf der Bestellseite ganz nach unten scrollen, „Bestellung abrufen“ anklicken und Bestellnummer sowie E-Mail-Adresse eingeben.</p>' +
    '<p style="margin:18px 0 0;">Viele Grüße<br>Landimkerei Diezinger</p>' +
    '<div style="margin-top:18px;padding-top:12px;border-top:1px solid #eee;font-size:12px;color:#999;">' +
    'Diese E-Mail wurde automatisch erstellt, bitte antworte nicht darauf. ' + contactHtml +
    '</div></td></tr></table></body></html>';
  const text = [
    "Servus " + (order.name || "") + ",",
    "",
    "vielen Dank für deine Bestellung. Wir haben sie erhalten.",
    "",
    "Bestellnummer: " + safeOrderNumber,
    ...items.map(function(item) { return "Bestellte Artikel: " + item; }),
    "Abholung: " + valueOrDash(order.pickup),
    "Wunschtermin: " + date,
    "Status: " + ORDER_STATUSES[0],
    "Wünsche: " + valueOrDash(order.note),
    "",
    "Wir melden uns, sobald deine Bestellung zur Abholung bereit ist.",
    statusLookup,
    "",
    "Viele Grüße",
    "Landimkerei Diezinger",
    "",
    "Diese E-Mail wurde automatisch erstellt, bitte antworte nicht darauf.",
    replyTo ? "Bei Fragen: " + replyTo : "",
    contacts.length ? "Eine Antwort erhältst du persönlich von: " + contacts.join(" oder ") : ""
  ].join("\n");
  MailApp.sendEmail({
    to: order.email,
    name: "Landimkerei Diezinger",
    subject: "Deine Bestellung " + (order.customerOrderNumber || "") + " – Landimkerei Diezinger",
    body: text,
    htmlBody: html
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
    if (id && id !== "HNG" && id !== "KZN" && id !== "K3D") unavailable[id] = row[4] === true;
    return unavailable;
  }, {});
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

function setupUnavailableReasonDropdowns_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const reasonRule = SpreadsheetApp.newDataValidation().requireValueInList(UNAVAILABLE_REASONS, true).build();
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
    reasonCell.setDataValidation(reasonRule);
  });
}

function migrateInventoryColumns_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 6) return;
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(value) { return String(value).trim(); });
  if (headers[5] !== "Preis (€)") return;

  const hasUnavailableReasons = headers[6] === "Grund nicht verfügbar";
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
        String(payload.pickup || ""),
        String(payload.date || ""),
        String(payload.note || ""),
        customerOrderNumber,
        ORDER_STATUSES[0]
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
        pickup: String(payload.pickup || ""),
        date: String(payload.date || ""),
        note: String(payload.note || "")
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function setupOrderTracking_(orders) {
  if (orders.getMaxColumns() < 10) orders.insertColumnsAfter(orders.getMaxColumns(), 10 - orders.getMaxColumns());
  orders.getRange(1, 9, 1, 2).setValues([["Bestellnummer", "Status"]]);
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

function generateCustomerOrderNumber_(existingNumbers) {
  let customerOrderNumber;
  do {
    customerOrderNumber = "H-" + Utilities.getUuid().replace(/-/g, "").substring(0, 10).toUpperCase();
  } while (existingNumbers.has(customerOrderNumber));
  return customerOrderNumber;
}

function lookupOrder_(payload) {
  const requestId = String(payload.requestId || "");
  const orderNumber = String(payload.customerOrderNumber || "").trim().toUpperCase();
  const email = String(payload.email || "").trim().toLowerCase();
  const response = {
    type: "orderLookupResult",
    requestId: requestId,
    ok: false,
    message: "Bestellung nicht gefunden. Bitte prüfe Bestellnummer und E-Mail-Adresse."
  };
  if (!/^[A-Za-z0-9_-]{20,80}$/.test(requestId) || !/^H-[A-F0-9]{10}$/.test(orderNumber) || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return response;

  const emailHash = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email)).replace(/=+$/g, "");
  const cache = CacheService.getScriptCache();
  const rateLimitKey = "order_lookup_" + emailHash;
  const attempts = Number(cache.get(rateLimitKey) || 0);
  if (attempts >= 5) {
    response.message = "Zu viele Versuche. Bitte warte einige Minuten und versuche es erneut.";
    return response;
  }
  cache.put(rateLimitKey, String(attempts + 1), 300);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return response;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  const row = rows.find(function(item) {
    return String(item[8] || "").trim().toUpperCase() === orderNumber && String(item[3] || "").trim().toLowerCase() === email;
  });
  if (!row) return response;

  return {
    type: "orderLookupResult",
    requestId: requestId,
    ok: true,
    order: {
      customerOrderNumber: String(row[8]),
      name: String(row[2] || ""),
      email: String(row[3] || ""),
      items: String(row[4] || ""),
      pickup: String(row[5] || ""),
      date: String(row[6] || ""),
      note: String(row[7] || ""),
      status: String(row[9] || ORDER_STATUSES[0])
    }
  };
}

function createOrderLookupResponse_(result) {
  const message = JSON.stringify(result)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
  return HtmlService.createHtmlOutput("<!doctype html><script>window.top.postMessage(" + message + ", '*');</script>")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
