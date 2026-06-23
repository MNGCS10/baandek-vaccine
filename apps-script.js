// ============================================================
// คลินิกบ้านเด็ก — Google Apps Script Web App
// วาง code นี้ทั้งหมดใน Apps Script แล้ว Deploy → New Deployment
// Execute as: Me | Who has access: Anyone
// ============================================================

function doGet(e) {
  var result = {};

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── CLINIC ──────────────────────────────────────────────
    var clinicSheet = ss.getSheetByName("CLINIC");
    if (clinicSheet) {
      var clinicData = clinicSheet.getDataRange().getValues();
      var clinic = {};
      clinicData.forEach(function(row) {
        if (row[0] && row[1] !== undefined) {
          clinic[String(row[0]).trim()] = String(row[1]).trim();
        }
      });
      result.clinic = {
        name:    clinic["name"]    || "คลินิกบ้านเด็ก",
        phone:   clinic["phone"]   || "",
        line:    clinic["line"]    || "",
        address: clinic["address"] || ""
      };
    }

    // ── PROMO ────────────────────────────────────────────────
    var promoSheet = ss.getSheetByName("PROMO");
    if (promoSheet) {
      var promoData = promoSheet.getDataRange().getValues();
      var promo = {};
      promoData.forEach(function(row) {
        if (row[0] && row[1] !== undefined) {
          promo[String(row[0]).trim()] = row[1];
        }
      });
      result.promo = {
        active: promo["active"] === true || String(promo["active"]).toUpperCase() === "TRUE",
        badge:  String(promo["badge"]  || ""),
        title:  String(promo["title"]  || ""),
        detail: String(promo["detail"] || ""),
        cta:    String(promo["cta"]    || ""),
        ctaUrl: String(promo["ctaUrl"] || "https://line.me/ti/p/@116hbawh")
      };
    }

    // ── VACCINES ─────────────────────────────────────────────
    var vaccineSheet = ss.getSheetByName("VACCINES");
    if (vaccineSheet) {
      var vaccineRows = vaccineSheet.getDataRange().getValues();
      var vaccines = {};
      for (var i = 1; i < vaccineRows.length; i++) {
        var row = vaccineRows[i];
        var code = String(row[0] || "").trim();
        if (!code) continue;
        vaccines[code] = {
          name:   String(row[1] || "").trim(),
          nameTH: String(row[2] || "").trim(),
          price:  Number(row[3]) || 0
        };
      }
      result.vaccines = vaccines;
    }

    // ── PACKAGES ─────────────────────────────────────────────
    var packageSheet = ss.getSheetByName("PACKAGES");
    if (packageSheet) {
      var packageRows = packageSheet.getDataRange().getValues();
      var packages = {};
      for (var j = 1; j < packageRows.length; j++) {
        var prow = packageRows[j];
        var ageId = String(prow[0] || "").trim();
        if (!ageId) continue;
        packages[ageId] = {
          label:       String(prow[1] || "").trim(),
          normalPrice: Number(prow[2]) || 0,
          proPrice:    Number(prow[3]) || 0,
          save:        Number(prow[4]) || 0
        };
      }
      result.packages = packages;
    }

    result.ok = true;
    result.ts = new Date().toISOString();

  } catch (err) {
    result = { ok: false, error: err.message };
  }

  // ── CORS: ต้องใช้ TextOutput + callback pattern ──────────
  var callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    // JSONP fallback
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
