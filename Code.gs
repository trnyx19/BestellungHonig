const LOW_STOCK_THRESHOLD = 6;
const INVENTORY_SHEET = "Bestand";
const ORDERS_SHEET = "Bestellungen";

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

function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let inventory = spreadsheet.getSheetByName(INVENTORY_SHEET);
  if (!inventory) inventory = spreadsheet.insertSheet(INVENTORY_SHEET);
  if (inventory.getLastRow() === 0) {
    inventory.getRange(1, 1, 1, 4).setValues([["ID", "Sorte", "Größe", "Bestand"]]);
    inventory.getRange(2, 1, INITIAL_PRODUCTS.length, 4).setValues(INITIAL_PRODUCTS);
    inventory.setFrozenRows(1);
    inventory.autoResizeColumns(1, 4);
  }

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
    const result = JSON.stringify({ ok: true, stock: readInventory_(), lowStockThreshold: LOW_STOCK_THRESHOLD });
    return ContentService.createTextOutput(callback + "(" + result + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch (error) {
    const result = JSON.stringify({ ok: false, message: error.message });
    return ContentService.createTextOutput(callback + "(" + result + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

function doPost(e) {
  let result;
  try {
    const payload = JSON.parse((e && e.parameter && e.parameter.payload) || "{}");
    result = reserveOrder_(payload);
  } catch (error) {
    result = { ok: false, message: error.message };
  }

  const safeResult = JSON.stringify(result).replace(/</g, "\\u003c");
  const html = "<!doctype html><meta charset=\"utf-8\"><script>parent.postMessage(" +
    safeResult + ", '*');</script>";
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function readInventory_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("Das Bestandsblatt ist nicht eingerichtet.");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return rows.reduce(function(stock, row) {
    stock[String(row[0])] = Number(row[3]);
    return stock;
  }, {});
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
    if (previous) return { ok: true, duplicate: true, stock: readInventory_(), lowStockThreshold: LOW_STOCK_THRESHOLD };

    const dataRange = inventory.getRange(2, 1, inventory.getLastRow() - 1, 4);
    const rows = dataRange.getValues();
    const rowById = {};
    rows.forEach(function(row, index) {
      const stock = Number(row[3]);
      if (!Number.isInteger(stock) || stock < 0) throw new Error("Im Bestandsblatt steht eine ungültige Stückzahl.");
      rowById[String(row[0])] = { row: index + 2, name: String(row[1]), size: String(row[2]), stock: stock };
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
      if (!product || quantity > 20) throw new Error("Eine ausgewählte Sorte oder Menge ist ungültig.");
      if (quantity > product.stock) {
        throw new Error("Für " + product.name + " (" + product.size + ") sind nicht mehr genug Gläser verfügbar.");
      }
      return { id: id, quantity: quantity, product: product };
    });

    const nextRows = rows.map(function(row) { return row.slice(); });
    purchases.forEach(function(item) {
      const index = item.product.row - 2;
      nextRows[index][3] = item.product.stock - item.quantity;
    });

    dataRange.setValues(nextRows);
    try {
      const itemSummary = purchases.map(function(item) {
        return item.quantity + "x " + item.product.name + " (" + item.product.size + ")";
      }).join(", ");
      orders.appendRow([
        orderId,
        new Date(),
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

    return { ok: true, stock: readInventory_(), lowStockThreshold: LOW_STOCK_THRESHOLD };
  } finally {
    lock.releaseLock();
  }
}
