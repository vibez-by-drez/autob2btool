function onFormSubmit(e) {
  // -----------------------------------------------------------
  // 1) Grunddaten aus dem Formular / Sheet holen
  // -----------------------------------------------------------
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Kunde: "Shipping Name" in Spalte B (Index 1)
  var customerName = rowData[1] || "";
  // Adresse: "Shipping Address" in Spalte C (Index 2)
  var customerAddress = rowData[2] || "";
  // Anzahl T-Shirts: aus Spalte "STÜCKZAHL" (Index 40)
  var anzahlTShirts = parseInt(rowData[40]) || 0;
  // Bestelltes Produkt (Textil): prüfe Spalten "T-Shirts"(35), "Hoodie"(36), "Hose"(37) oder "Jacke/Mutze/Socken"(38)
  var textil = "";
  if (rowData[35] && rowData[35].toString().trim() !== "") {
    textil = rowData[35].toString().trim().toLowerCase();
  } else if (rowData[36] && rowData[36].toString().trim() !== "") {
    textil = rowData[36].toString().trim().toLowerCase();
  } else if (rowData[37] && rowData[37].toString().trim() !== "") {
    textil = rowData[37].toString().trim().toLowerCase();
  } else if (rowData[38] && rowData[38].toString().trim() !== "") {
    textil = rowData[38].toString().trim().toLowerCase();
  }
  
  // -----------------------------------------------------------
  // 2) Versandmethode aus Spalte "VERSANDART" (Index 48)
  // -----------------------------------------------------------
  var shippingMethod = (rowData[48] || "").toString().trim();

  // -----------------------------------------------------------
  // 3) Veredelungsdaten aus 4 Gruppen extrahieren
  // -----------------------------------------------------------
  var veredelungItems = [];
  
  // Gruppe 1 (Spalten: Art:6, Size:8, Motiv Size:9, Siebdruckfarben:11, Schmuckfarbe:12)
  if ((rowData[6] || "").toString().trim() !== "") {
    veredelungItems.push({
      type: rowData[6].toString().trim(),
      dimension: (rowData[8] || "").toString().trim(),
      farben: rowData[11] ? rowData[11].toString().trim() : "1", // falls leer als "1 Farbig" annehmen
      schmuck: rowData[12] ? rowData[12].toString().trim() : ""
    });
  }
  // Gruppe 2 (Spalten: Art:13, Size:15, Motiv Size:16, Siebdruckfarben:17, Schmuckfarbe:19)
  if ((rowData[13] || "").toString().trim() !== "") {
    veredelungItems.push({
      type: rowData[13].toString().trim(),
      dimension: (rowData[15] || "").toString().trim(),
      farben: rowData[17] ? rowData[17].toString().trim() : "1",
      schmuck: rowData[19] ? rowData[19].toString().trim() : ""
    });
  }
  // Gruppe 3 (Spalten: Art:20, Size:22, Motiv Size:23, Siebdruckfarben:25, Schmuckfarbe:26)
  if ((rowData[20] || "").toString().trim() !== "") {
    veredelungItems.push({
      type: rowData[20].toString().trim(),
      dimension: (rowData[22] || "").toString().trim(),
      farben: rowData[25] ? rowData[25].toString().trim() : "1",
      schmuck: rowData[26] ? rowData[26].toString().trim() : ""
    });
  }
  // Gruppe 4 (Spalten: Art:27, Size:29, Motiv Size:30, Siebdruckfarben:32, Schmuckfarbe:33)
  if ((rowData[27] || "").toString().trim() !== "") {
    veredelungItems.push({
      type: rowData[27].toString().trim(),
      dimension: (rowData[29] || "").toString().trim(),
      farben: rowData[32] ? rowData[32].toString().trim() : "1",
      schmuck: rowData[33] ? rowData[33].toString().trim() : ""
    });
  }
  
  Logger.log("Kunde: " + customerName);
  Logger.log("Adresse: " + customerAddress);
  Logger.log("Anzahl T-Shirts: " + anzahlTShirts);
  Logger.log("Textil: " + textil);
  Logger.log("Versandmethode: " + shippingMethod);
  Logger.log("Veredelungen: " + veredelungItems.map(function(item) {
    return item.type + " - " + item.dimension;
  }).join(" | "));

  // -----------------------------------------------------------
  // 4) Veredelungen durchgehen & Zeiten/Kosten berechnen
  // -----------------------------------------------------------
  var spreadsheetId = "1tiVehMD146ROAeatymhPCaGaLB8GCcwd2lyYvshE2I8"; 
  var kostenProMinute = parseFloat(
    SpreadsheetApp.openById(spreadsheetId)
      .getSheetByName("Sheet1")
      .getRange("A1")
      .getValue()
  ) || 0;
  Logger.log("Kosten pro Minute aus Tabelle: " + kostenProMinute);
  
  var arbeitskostenGesamt = 0;
  var veredelungRows = [];
  var verwendeteMethoden = new Set();
  
  veredelungItems.forEach(function(item) {
    var type = item.type;
    var dimensionFull = item.dimension;
    verwendeteMethoden.add(type.toUpperCase());
    
    var minutenProStk = getMinutesForVeredelung(type, dimensionFull);
    var baseEinzelpreis = minutenProStk * kostenProMinute;
    
    // --- Rabatt für Veredelung (separat) anwenden ---
    var rabattVeredelung = getVeredelungDiscount(anzahlTShirts);
    baseEinzelpreis = baseEinzelpreis * (1 - rabattVeredelung);
    
    // Falls es sich um SIEBDRUCK handelt, wird der Preis weiter angepasst:
    if (type.toUpperCase() === "SIEBDRUCK") {
      // Extrahiere die Anzahl Farben (z. B. "2 Farbig")
      var match = item.farben.match(/^(\d+)/);
      var colorCount = match ? parseInt(match[1], 10) : 1;
      // Berechne den prozentualen Aufschlag: (Anzahl - 1) * 10 %, maximal 40%
      var surchargePct = Math.min((colorCount - 1) * 0.1, 0.4);
      baseEinzelpreis = baseEinzelpreis * (1 + surchargePct);
      // Zuschlag für Schmuckfarbe, sofern gewählt
      if (item.schmuck && item.schmuck.toLowerCase() !== "keine schmuckfarbe") {
        baseEinzelpreis += 0.30;
      }
    }
    
    var einzelpreis = parseFloat(baseEinzelpreis.toFixed(2));
    var gesamtpreis = parseFloat((einzelpreis * anzahlTShirts).toFixed(2));
    arbeitskostenGesamt += gesamtpreis;
    
    veredelungRows.push({
      type: type + " - " + dimensionFull,
      menge: anzahlTShirts,
      einzelpreis: einzelpreis,
      gesamtpreis: gesamtpreis
    });
  });
  
  Logger.log("Arbeitskosten (ohne Einrüstkosten): " + arbeitskostenGesamt.toFixed(2) + " €");
  
  // -----------------------------------------------------------
  // 4a) Einrüstkosten je Methode aktualisieren
  // -----------------------------------------------------------
  var einruestKostenTotal = 0;
  var einruestRows = [];
  Array.from(verwendeteMethoden).forEach(function(method) {
    var methodUpper = method.toUpperCase();
    var kosten = 0;
    if (methodUpper === "STICK") {
      kosten = 30;
    } else if (methodUpper === "DTF") {
      kosten = 20;
    } else if (methodUpper === "SIEBDRUCK") {
      kosten = 40; // SIEB: Aufschlag wird jetzt direkt in den Einzelpreisen berücksichtigt.
    }
    if (kosten > 0) {
      einruestKostenTotal += kosten;
      einruestRows.push({
        type: "Einrüstkosten " + methodUpper,
        menge: "-",
        einzelpreis: kosten,
        gesamtpreis: kosten
      });
    }
  });
  if (einruestKostenTotal > 0) {
    arbeitskostenGesamt += einruestKostenTotal;
    einruestRows.forEach(function(row) {
      veredelungRows.push(row);
    });
    Logger.log("Einrüstkosten insgesamt = " + einruestKostenTotal + " €");
  }
  
  Logger.log("Arbeitskosten inkl. Einrüstkosten: " + arbeitskostenGesamt.toFixed(2) + " €");
  
  // -----------------------------------------------------------
  // 5) Materialkosten, Nettopreis, Versandkosten, Netto/Brutto berechnen
  // -----------------------------------------------------------
  var materialKosten = getMaterialCost(textil, anzahlTShirts);
  var nettoPriceTextil = materialKosten * anzahlTShirts;
  
  var shippingCost = 0;
  if (shippingMethod === "Standard Versand") {
    shippingCost = Math.ceil(anzahlTShirts / 100) * 10;
  } else if (shippingMethod === "Express Versand") {
    shippingCost = Math.ceil(anzahlTShirts / 100) * 30;
  }
  
  var originalNetto = nettoPriceTextil + arbeitskostenGesamt;
  var nettoPrice_pdf = originalNetto + shippingCost;
  var steuer = nettoPrice_pdf * 0.19;
  var bruttoPrice = nettoPrice_pdf + steuer;
  
  // -----------------------------------------------------------
  // 6) Rechnungsnummer generieren + speichern
  // -----------------------------------------------------------
  var nummerSheetId = "1utn2kw4WHBYqWm14Z3Bx5Om9_nd2rilhcxFtPWZwa2A";
  var nummerSheetName = "Rechnungsnummern";
  var lastNumber = getLastRechnungsnummer(nummerSheetId, nummerSheetName);
  var rechnungsnummer = generateNextRechnungsnummer(lastNumber);
  saveRechnungsnummer(nummerSheetId, nummerSheetName, rechnungsnummer);
  
  // Rechnungnummer in Spalte Q (Index 17) und "0" in Spalte R (Index 18) eintragen
  sheet.getRange(lastRow, 50).setValue(rechnungsnummer);
  sheet.getRange(lastRow, 51).setValue(0);
  sheet.getRange(lastRow, 52).setValue(nettoPrice_pdf);
  Logger.log("Rechnungsnummer und Status in Zeile " + lastRow + " aktualisiert.");
  
  // -----------------------------------------------------------
  // 7) PDF erstellen (Rechnung + Angebot)
  // -----------------------------------------------------------
  var logoBase64 = "data:image/png;base64,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"; // (gekürzt)
  
  var veredelungDetailsHTML = buildVeredelungsHtml(veredelungRows);
  
  var rechnungHtml = createHtmlTemplate(
    "Rechnung",
    rechnungsnummer,
    customerName,
    customerAddress,
    anzahlTShirts,
    textil,
    materialKosten,
    nettoPriceTextil,
    arbeitskostenGesamt,
    originalNetto,
    shippingCost,
    shippingMethod,
    nettoPrice_pdf,
    steuer,
    bruttoPrice,
    veredelungDetailsHTML,
    logoBase64
  );
  var rechnungBlob = Utilities.newBlob(rechnungHtml, "text/html")
                    .getAs("application/pdf")
                    .setName("Rechnung_" + rechnungsnummer + ".pdf");
  
  var angebotHtml = createHtmlTemplate(
    "Angebot",
    rechnungsnummer,
    customerName,
    customerAddress,
    anzahlTShirts,
    textil,
    materialKosten,
    nettoPriceTextil,
    arbeitskostenGesamt,
    originalNetto,
    shippingCost,
    shippingMethod,
    nettoPrice_pdf,
    steuer,
    bruttoPrice,
    veredelungDetailsHTML,
    logoBase64
  );
  var angebotBlob = Utilities.newBlob(angebotHtml, "text/html")
                    .getAs("application/pdf")
                    .setName("Angebot_" + rechnungsnummer + ".pdf");
  
  var folderId = "1bxr_val8cWusAPM3uQ3MkxQ_6DrGsPhL";
  var folder = DriveApp.getFolderById(folderId);
  var rechnungFile = folder.createFile(rechnungBlob);
  var angebotFile = folder.createFile(angebotBlob);
  
  Logger.log("Rechnung gespeichert: " + rechnungFile.getUrl());
  Logger.log("Angebot gespeichert: " + angebotFile.getUrl());
  
  // -----------------------------------------------------------
  // 8) Mail an Kunden + CEO schicken
  // -----------------------------------------------------------
  var customerEmail = (rowData[3] || "").toString().trim();
  if (!customerEmail) {
    Logger.log("Keine Kunden-Mail hinterlegt (Spalte D leer)");
  }
  
  MailApp.sendEmail({
    to: customerEmail,
    cc: "ceo@printstudios.de",
    subject: "Ihre Dokumente (" + rechnungsnummer + ")",
    body: "Hallo " + customerName + ",\n\nanbei senden wir Ihnen Ihr Angebot.\n\nZur Annahme unseres Angebots überweisen Sie bitte den in der beiliegenden Rechnung aufgeführten Betrag.\nDadurch wird der Produktionsvorgang automatisch gestartet.\n\nWir freuen uns sehr, Ihr nächstes Textilprojekt umsetzen zu dürfen.\n\nMit freundlichen Grüßen\nIhr Printstudios-Team\n\nBei Fragen melden Sie sich gerne unter help@printstudios.de\nBeachten Sie unsere AGBs unter www.Printstudios.de/AGBs\nIhr Printstudios-Team",
    attachments: [rechnungFile.getBlob(), angebotFile.getBlob()]
  });
  
  Logger.log("E-Mail an " + customerEmail + " + CEO versendet.");
  Logger.log("Fertig!");
}


// -----------------------------------------------------------------------
// HILFSFUNKTIONEN
// -----------------------------------------------------------------------

function getMinutesForVeredelung(type, dimensionFull) {
  var dimLower = dimensionFull.toLowerCase();
  var dimension = "";
  if (dimLower.indexOf("klein dina6") !== -1) {
    dimension = "KLEIN";
  } else if (dimLower.indexOf("medium dina4") !== -1) {
    dimension = "MEDIUM";
  } else if (dimLower.indexOf("groß dina3") !== -1) {
    dimension = "GROSS";
  } else {
    Logger.log("Unbekannte Motivgröße: " + dimensionFull);
    return 0;
  }
  
  switch (type.toUpperCase()) {
    case "SIEBDRUCK":
      if (dimension === "KLEIN") return 0.6;
      if (dimension === "MEDIUM") return 0.75;
      if (dimension === "GROSS") return 1;
      break;
    case "STICK":
      if (dimension === "KLEIN") return 2;
      if (dimension === "MEDIUM") return 4;
      if (dimension === "GROSS") return 6;
      break;
    case "DTF":
      if (dimension === "KLEIN") return 1.25;
      if (dimension === "MEDIUM") return 2.25;
      if (dimension === "GROSS") return 2.8;
      break;
    default:
      Logger.log("Unbekannter Verfahrenstyp: " + type);
      return 0;
  }
  return 0;
}

function getMaterialCost(textil, menge) {
  var basePrice;
  switch (textil) {
    case "classicfit tee":
      basePrice = 2.94;
      break;
    case "royalcomfort tee":
      basePrice = 3.72;
      break;
    case "artisan tee":
      basePrice = 5.75;
      break;
    case "freeflow tee":
      basePrice = 13.83;
      break;
    case "urbanheavy tee":
      basePrice = 8.09;
      break;
    case "streetstyle tee":
      basePrice = 16.43;
      break;
    case "retrobox tee":
      basePrice = 16.13;
      break;
    case "everyday tee":
      basePrice = 15.68;
      break;
    case "cropchic tee":
      basePrice = 14.78;
      break;
    case "empireelite tee":
      basePrice = 12.75;
      break;
    case "empiremax tee":
      basePrice = 15.00;
      break;
    case "authentichoodie":
      basePrice = 19.17;
      break;
    case "ziphood classic":
      basePrice = 23.40;
      break;
    case "heritagesweat":
      basePrice = 14.41;
      break;
    case "boxfit hoodie":
      basePrice = 33.92;
      break;
    case "cruiserhoodie":
      basePrice = 27.75;
      break;
    case "streetstriker hoodie":
      basePrice = 41.76;
      break;
    case "ultrastreet hoodie":
      basePrice = 74.85;
      break;
    case "loosefit hoodie":
      basePrice = 44.85;
      break;
    case "standardhoodie":
      basePrice = 43.35;
      break;
    case "zipoversize hoodie":
      basePrice = 50.85;
      break;
    case "zipfit hoodie":
      basePrice = 49.35;
      break;
    case "cropcomfort hoodie":
      basePrice = 38.85;
      break;
    case "relaxedcrew sweater":
      basePrice = 38.85;
      break;
    case "classiccrew sweater":
      basePrice = 36.90;
      break;
    case "empirecrop heavy":
      basePrice = 31.50;
      break;
    case "cozyempire hoodie":
      basePrice = 27.75;
      break;
    case "empireessentials hoodie":
      basePrice = 22.50;
      break;
    case "empirechill hoodie":
      basePrice = 22.50;
      break;
    case "imperiallong tee":
      basePrice = 6.96;
      break;
    case "shufflestyle tee":
      basePrice = 10.10;
      break;
    case "loosefit long sleeve":
      basePrice = 20.78;
      break;
    case "trainershorts":
      basePrice = 20.44;
      break;
    case "easyfit shorts":
      basePrice = 34.95;
      break;
    case "empirejogger shorts":
      basePrice = 21.00;
      break;
    case "authenticjoggers":
      basePrice = 21.63;
      break;
    case "comfortsweats":
      basePrice = 43.35;
      break;
    case "empirelounge pants":
      basePrice = 21.00;
      break;
    case "moveflex pants":
      basePrice = 27.50;
      break;
    case "snapbackoriginal":
      basePrice = 9.39;
      break;
    case "lowkey cap":
      basePrice = 10.83;
      break;
    case "twotone trucker":
      basePrice = 6.51;
      break;
    case "heavywarm beanie":
      basePrice = 5.63;
      break;
    case "harbourbeanie":
      basePrice = 4.29;
      break;
    case "crewstep socks":
      basePrice = 2.79;
      break;
    case "courtclassic socks":
      basePrice = 2.85;
      break;
    case "bomberedge jacket":
      basePrice = 31.20;
      break;
    case "coachpro jacket":
      basePrice = 27.12;
      break;
    case "collegevibe jacket":
      basePrice = 35.27;
      break;
    case "sols regent":
      basePrice = 2.12;
      break;
    case "Keins":
      basePrice = 0.0;
      break;
    default:
      basePrice = 50.0;
      break;
  }
  
  var rabatt = 0;
  if (menge >= 500) {
    rabatt = 0.15;
  } else if (menge >= 250) {
    rabatt = 0.12;
  } else if (menge >= 100) {
    rabatt = 0.09;
  } else if (menge >= 50) {
    rabatt = 0.06;
  } else if (menge >= 25) {
    rabatt = 0.03;
  }
  
  return basePrice * (1 - rabatt);
}

// Neue Funktion: Veredelungs-Mengenrabatt mit angepassten Prozentwerten
function getVeredelungDiscount(menge) {
  // Beispielwerte – diese können Sie nach Belieben anpassen:
  if (menge >= 1000) {
    return 0.50; // 50% Rabatt
  } else if (menge >= 700) {
    return 0.40; // 40% Rabatt    
  } else if (menge >= 400) {
    return 0.30; // 30% Rabatt
  } else if (menge >= 250) {
    return 0.20; // 20% Rabatt
  } else if (menge >= 100) {
    return 0.10; // 10% Rabatt
  } else if (menge >= 50) {
    return 0.05; // 5% Rabatt
  }
  return 0;
}

function createHtmlTemplate(titel, rechnungsnummer, customerName, customerAddress, anzahlTShirts, textil, materialKosten, nettoPriceTextil, arbeitskostenGesamt, originalNetto, shippingCost, shippingMethod, nettoPrice_pdf, steuer, bruttoPrice, veredelungDetailsHTML, logoBase64) {
  var heute = new Date().toLocaleDateString();
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <style>
    /* Erhöhe den unteren Abstand, damit kein Inhalt abgeschnitten wird */
    body {
      font-family: Arial, sans-serif;
      margin: 40px;
      margin-top: 1px;
      margin-bottom: 200px; /* erhöhter Abstand */
      line-height: 1.5;
      color: #333;
    }
    .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .header img {
      height: 180px;
    }
    h1 {
      font-size: 1.4em;
      margin-top: 10px;
    }
    .table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 20px;
    }
    .table th, .table td {
      font-size: 0.9em;
      text-align: left;
      padding: 8px;
      border-bottom: 1px solid #ddd;
      white-space: pre-wrap;
    }
    /* Zusammenfassungsblock so definieren, dass er nicht aufgeteilt wird */
    .summary-block {
      page-break-inside: avoid;
      page-break-after: avoid;
      padding-bottom: 20px;
    }
    .summary {
      margin-top: 20px;
      font-size: 1.1em;
    }
    .info-box {
      margin-top: 15px;
      padding: 10px;
      background: #f2f2f2;
      border-radius: 5px;
    }
    /* Anstatt eines fixed Footers den Footer statisch am Ende platzieren */
    .footer {
      position: fixed;
      bottom: 0;
      left: 0;
      width: 100%;
      background: #f9f9f9;
      text-align: center;
      font-size: 0.7em;
      padding: 2px;
    }
    .page-break {
      page-break-after: always;
    }
  </style>
</head>
<body>
  <div class="header">
    <img src="${logoBase64}" alt="Logo" />
    <div>
      <p>Fürth, ${heute}</p>
      <p>${titel} Nr.: <strong>${rechnungsnummer}</strong></p>
    </div>
  </div>
  <h1>${titel}</h1>
  <div class="info-box">
    <p><strong>Kunde:</strong> ${customerName}</p>
    <p><strong>Adresse:</strong> ${customerAddress}</p>
  </div>
  <h2>Textil / Produkt</h2>
  <table class="table">
    <tr>
      <th>Pos.</th>
      <th>Textil</th>
      <th>Menge</th>
      <th>Materialkosten/Stk</th>
      <th>Gesamt (nur Material)</th>
    </tr>
    <tr>
      <td>1</td>
      <td>${textil}</td>
      <td>${anzahlTShirts}</td>
      <td>${materialKosten.toFixed(2)} €</td>
      <td>${nettoPriceTextil.toFixed(2)} €</td>
    </tr>
  </table>
  <h2>Veredelungen</h2>
  ${veredelungDetailsHTML}
  <div class="summary-block">
    <div class="summary">
      <p><strong>Nettobetrag (Material + Veredelung):</strong> ${originalNetto.toFixed(2)} €</p>
      <p><strong>Versandkosten (${shippingMethod}):</strong> ${shippingCost.toFixed(2)} €</p>
      <p><strong>Gesamtnettopreis:</strong> ${nettoPrice_pdf.toFixed(2)} €</p>
      <p><strong>19% MwSt:</strong> ${steuer.toFixed(2)} €</p>
      <p><strong>Gesamtbetrag (Brutto):</strong> ${bruttoPrice.toFixed(2)} €</p>
    </div>
  </div>
  <div class="footer">
    <p>Geschäftsführer: Moritz Veth | Firma: Printstudios | Benno-Strauß-Straße 5b / 2 OG 90763 Fürth Ust.-Id-Nr.: DE279791900</p>
    <p>BANK: N26 | IBAN: DE64100110012294247221 | BIC: NTSBDEB1XXX</p>
  </div>
</body>
</html>
`;
}

function getLastRechnungsnummer(sheetId, sheetName) {
  var nummernSheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  if (!nummernSheet) {
    throw new Error("Das angegebene Sheet mit dem Namen '" + sheetName + "' wurde nicht gefunden.");
  }
  var lastRow = nummernSheet.getLastRow();
  return lastRow > 0 ? nummernSheet.getRange(lastRow, 1).getValue() : "25-000";
}

function generateNextRechnungsnummer(lastNumber) {
  var parts = lastNumber.split("-");
  var prefix = parts[0];
  var num = parseInt(parts[1], 10) + 1;
  return prefix + "-" + num.toString().padStart(3, "0");
}

function saveRechnungsnummer(sheetId, sheetName, nextNumber) {
  var nummernSheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  if (!nummernSheet) {
    throw new Error("Das angegebene Sheet mit dem Namen '" + sheetName + "' wurde nicht gefunden.");
  }
  var lastRow = nummernSheet.getLastRow();
  nummernSheet.getRange(lastRow + 1, 1).setValue(nextNumber);
  Logger.log("Neue Rechnungsnummer gespeichert: " + nextNumber);
}

function buildVeredelungsHtml(rows) {
  var html = "";
  var posCounter = 0;
  
  // Auf Seite 1 max. 8 Einträge
  var firstPageRows = rows.slice(0, 7);
  var subsequentRows = rows.slice(8);
  
  if (firstPageRows.length > 0) {
    html += `
      <table class="table">
        <tr>
          <th>Pos.</th>
          <th>Veredelung</th>
          <th>Menge</th>
          <th>Einzelkosten</th>
          <th>Gesamtkosten</th>
        </tr>
    `;
    firstPageRows.forEach(function(row) {
      posCounter++;
      var einzelStr = (typeof row.einzelpreis === 'number') ? row.einzelpreis.toFixed(2) + " €" : row.einzelpreis;
      var gesamtStr = (typeof row.gesamtpreis === 'number') ? row.gesamtpreis.toFixed(2) + " €" : row.gesamtpreis;
      html += `
        <tr>
          <td>${posCounter}</td>
          <td>${row.type}</td>
          <td>${row.menge}</td>
          <td>${einzelStr}</td>
          <td>${gesamtStr}</td>
        </tr>
      `;
    });
    html += `</table>`;
  }
  
  if (subsequentRows.length > 0) {
    html += '<div class="page-break"></div>';
  }
  
  var PAGE_SIZE = 20;
  var allChunks = [];
  for (var i = 0; i < subsequentRows.length; i += PAGE_SIZE) {
    allChunks.push(subsequentRows.slice(i, i + PAGE_SIZE));
  }
  
  allChunks.forEach(function(chunk, idx) {
    html += `
      <table class="table">
        <tr>
          <th>Pos.</th>
          <th>Veredelung</th>
          <th>Menge</th>
          <th>Einzelkosten</th>
          <th>Gesamtkosten</th>
        </tr>
    `;
    chunk.forEach(function(row) {
      posCounter++;
      var einzelStr = (typeof row.einzelpreis === 'number') ? row.einzelpreis.toFixed(2) + " €" : row.einzelpreis;
      var gesamtStr = (typeof row.gesamtpreis === 'number') ? row.gesamtpreis.toFixed(2) + " €" : row.gesamtpreis;
      html += `
        <tr>
          <td>${posCounter}</td>
          <td>${row.type}</td>
          <td>${row.menge}</td>
          <td>${einzelStr}</td>
          <td>${gesamtStr}</td>
        </tr>
      `;
    });
    html += `</table>`;
    if (idx < allChunks.length - 1) {
      html += '<div class="page-break"></div>';
    }
  });
  
  return html;
}
