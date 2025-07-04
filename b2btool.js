// frontend/pages/deineSeite.js
import { sendToSheet } from 'backend/formHandler';

// Hilfsfunktion für Delay
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Mapping von Textilien zu Beschreibungen
const descriptions = {
  'classicfit tee':   "Das Classic Fit Tee ist unfassbar gut für draußen tragen",
  'royalcomfort tee': "Das Royalcomfort Tee ist aus Griechenland",
  'artisan tee':      "Artisan Tee für Partisanen, die feine Stoffe schätzen"
};

// Hilfsfunktion: Beschreibung im Textfeld setzen
function updateDescription() {
  const textil = getTextil();
  const desc = descriptions[textil] || '';
  $w('#textdescription').text = desc;
}

$w.onReady(() => {
  const prefixes = ['V1', 'V2', 'V3', 'V4'];

  // 1) Initial: alle Siebdruck-Felder verstecken
  prefixes.forEach(prefix => {
    const colorField = $w(`#${prefix}druckfarbenamount`);
    const jewelField = $w(`#${prefix}schmuckfarbe`);
    if (colorField) colorField.hide();
    if (jewelField) jewelField.hide();
  });

  // 2) onChange-Handler für alle Art-Dropdowns
  prefixes.forEach(prefix => {
    const artField = $w(`#${prefix}art`);
    artField.onChange(event => {
      const value = event.target.value;
      console.log(`${prefix}art ausgewählt:`, value);
      const colorField = $w(`#${prefix}druckfarbenamount`);
      const jewelField = $w(`#${prefix}schmuckfarbe`);
      if (value === 'Siebdruck') {
        colorField.show();
        jewelField.show();
      } else {
        colorField.hide();
        jewelField.hide();
      }
    });
  });

  // 3) Beschreibungstext aktualisieren bei Auswahl Textil
  ['#ddshirts', '#ddhoodies', '#ddhosen', '#ddstuff'].forEach(id => {
    $w(id).onChange(() => updateDescription());
  });
  // initial setzen
  updateDescription();

  // 4) Klick-Handler binden
  $w('#buttonsubmit').onClick(submitButton_click);

  // 5) Preis-Button binden (falls vorhanden)
  if ($w('#btnprice')) {
    $w('#btnprice').onClick(() => calculatePrice());
  }
});


// ************************************************************************** //
// *********************** calculatePrice-Funktion ********************** //
// ************************************************************************** //
export async function calculatePrice() {
  console.log('=== calculatePrice gestartet ===');

  // Fortschrittsbalken initialisieren
  const totalSteps = 6;
  let step = 0;
  $w('#progressBar1').max = totalSteps;
  $w('#progressBar1').value = step;
  $w('#progressBar1').show();
  $w('#progressText').show();

  // --------- Schritt 1: Basisdaten sammeln ---------
  $w('#progressText').text = 'Erfasse Basisdaten…';
  await sleep(420);

  console.log('Schritt 1: Basisdaten sammeln');
  const now = new Date();
  const dateOfEntry = now.toLocaleString('de-DE', { timeZone: 'Europe/Berlin' });
  const shippingAddress = [
    $w('#inputland').value,
    $w('#inputstadt').value,
    $w('#inputpostleitzahl').value,
    $w('#inputstrasse').value
  ].join(', ');
  const anzahlTShirts = parseInt($w('#inputamount').value) || 0;
  const textil = getTextil();
  const shippingMethod = $w('#ddversand').value;
  console.log({ dateOfEntry, shippingAddress, anzahlTShirts, textil, shippingMethod });
  $w('#progressBar1').value = 15;

  // --------- Schritt 2: Veredelungs-Items sammeln ---------
  $w('#progressText').text = 'Erfasse Veredelungs-Items…';
  await sleep(400);

  console.log('Schritt 2: Veredelungs-Items sammeln');
  const prefixes = ['V1', 'V2', 'V3', 'V4'];
  const veredelungItems = prefixes
    .map(pref => {
      const type = $w(`#${pref}art`).value;
      if (!type) return null;
      return {
        type,
        dimension: $w(`#${pref}size`).value || '',
        farben:   $w(`#${pref}druckfarbenamount`).value || '1 Farbig',
        schmuck:  $w(`#${pref}schmuckfarbe`).value || ''
      };
    })
    .filter(x => x);
  console.log('Veredelung Items:', veredelungItems);
  $w('#progressBar1').value = 30;

  // --------- Schritt 3: Arbeits- und Einrüstkosten berechnen ---------
  $w('#progressText').text = 'Berechne Arbeits- & Einrüstkosten…';
  await sleep(300);

  console.log('Schritt 3: Arbeits- und Einrüstkosten berechnen');
  const kostenProMinute = 1.0;
  let arbeitskostenGesamt = 0;
  const methodSet = new Set();

  veredelungItems.forEach(item => {
    const t = item.type.toUpperCase();
    methodSet.add(t);

    // Minuten pro Shirt
    const minuten = getMinutesForVeredelung(t, item.dimension);
    console.log(`  [${item.type}] Minuten/Stück:`, minuten);

    // Basispreis
    let basis = minuten * kostenProMinute;
    console.log(`  [${item.type}] Basispreis/Stück:`, basis);

    // Veredelungsrabatt
    const rabatt = getVeredelungDiscount(anzahlTShirts);
    console.log(`  [${item.type}] Rabatt für ${anzahlTShirts} Stück: ${(rabatt*100).toFixed(0)}%`);
    basis *= 1 - rabatt;
    console.log(`  [${item.type}] Nach Rabatt/Stück:`, basis);

    // Siebdruck-Zuschlag
    if (t === 'SIEBDRUCK') {
      const colorCount = parseInt(item.farben) || 1;
      const surchargePct = Math.min((colorCount - 1) * 0.1, 0.4);
      console.log(`  [${item.type}] Farbanzahl: ${colorCount}, Zuschlag: ${(surchargePct*100).toFixed(0)}%`);
      basis *= 1 + surchargePct;
      if (item.schmuck.toLowerCase() !== 'keine schmuckfarbe') {
        console.log(`  [${item.type}] Schmuckfarbe gewählt, +0.30`);
        basis += 0.30;
      }
      console.log(`  [${item.type}] Nach Siebdruck/Stück:`, basis);
    }

    // Einzel- und Gesamtpreis Veredelung
    const einzelpreis = parseFloat(basis.toFixed(2));
    const gesamtVered = einzelpreis * anzahlTShirts;
    console.log(`  [${item.type}] Einzelpreis: ${einzelpreis}, Gesamt Veredelung: ${gesamtVered}`);
    arbeitskostenGesamt += gesamtVered;
  });

  console.log('  Arbeitskosten gesamt (ohne Einrüstkosten):', arbeitskostenGesamt);

  let einruestKostenTotal = 0;
  methodSet.forEach(m => {
    if (m === 'STICK')        einruestKostenTotal += 30;
    else if (m === 'DTF')     einruestKostenTotal += 20;
    else if (m === 'SIEBDRUCK') einruestKostenTotal += 40;
  });
  console.log('  Einrüstkosten gesamt:', einruestKostenTotal);
  arbeitskostenGesamt += einruestKostenTotal;
  console.log('  Arbeitskosten inkl. Einrüstkosten:', arbeitskostenGesamt);
  $w('#progressBar1').value = 45;

  // --------- Schritt 4: Materialkosten berechnen ---------
  $w('#progressText').text = 'Berechne Materialkosten…';
  await sleep(600);

  console.log('Schritt 4: Materialkosten berechnen');
  const materialProStk = getMaterialCost(textil, anzahlTShirts);
  const nettoTextilGesamt = materialProStk * anzahlTShirts;
  console.log('  Materialkosten/Stück:', materialProStk, ' Gesamt:', nettoTextilGesamt);
  $w('#progressBar1').value = 60;

  // --------- Schritt 5: Versandkosten berechnen ---------
  $w('#progressText').text = 'Addiere Versandkosten…';
  await sleep(450);

  console.log('Schritt 5: Versandkosten berechnen');
  let shippingCost = 0;
  if (shippingMethod === 'Standard Versand') {
    shippingCost = Math.ceil(anzahlTShirts / 100) * 10;
  } else if (shippingMethod === 'Express Versand') {
    shippingCost = Math.ceil(anzahlTShirts / 100) * 30;
  }
  console.log('  Versandkosten gesamt:', shippingCost);
  $w('#progressBar1').value = 75;

  // --------- Schritt 6: Gesamtnettopreis berechnen ---------
  $w('#progressText').text = 'Berechne Netto-Preis…';
  await sleep(500);

  console.log('Schritt 6: Gesamt Nettopreis berechnen');
  const nettoPrice = parseFloat(
    (nettoTextilGesamt + arbeitskostenGesamt + shippingCost).toFixed(2)
  );
  console.log('  Summe Arbeitskosten:', arbeitskostenGesamt,
              '+ Material:', nettoTextilGesamt,
              '+ Versand:', shippingCost,
              '= Netto:', nettoPrice);
  $w('#pricetag').text = nettoPrice.toFixed(2).replace('.', ',') + ' €';
  $w('#progressBar1').value = 100;

  // Fertig
  $w('#progressText').text = 'Fertig!';
}


// ************************************************************************** //
// *********************** submitButton_click-Funktion ********************** //
// ************************************************************************** //
function submitButton_click(event) {
  console.log('=== submitButton_click gestartet ===');

  // --------- 1) Datum & Adresse ---------
  const now = new Date();
  const dateOfEntry = now.toLocaleString('de-DE', { timeZone: 'Europe/Berlin' });
  const shippingAddress = [
    $w('#inputland').value,
    $w('#inputstadt').value,
    $w('#inputpostleitzahl').value,
    $w('#inputstrasse').value
  ].join(', ');
  console.log('Date:', dateOfEntry, 'Adresse:', shippingAddress);

  // --------- 2) Felder sammeln ---------
  const data = {
    "Date of Entry":             dateOfEntry,
    "Shipping Name":             $w('#inputname').value,
    "Shipping Address":          shippingAddress,
    "E-Mail":                    $w('#inputemail').value,
    "Phone":                     $w('#inputtelefon').value,
    "Shipment Date":             $w('#dateversand').value,

    // V1
    "1 Veredelungs Art":         $w('#V1art').value,
    "1. Veredelungsposition":    $w('#V1position').value,
    "1. Veredelungs Größe":      $w('#V1size').value,
    "1. Genaue Motiv Größe in mm": $w('#V1mmsize').value,
    "1. Veredelungs Motiv":      $w('#V1upload').value,
    "1. Siebdruckfarben Anzahl": $w('#V1druckfarbenamount').value,
    "1. Siebdruck Schmuckfarbe": $w('#V1schmuckfarbe').value,

    // V2
    "2. Veredelungs Art":        $w('#V2art').value,
    "2. Veredelungsposition":    $w('#V2position').value,
    "2. Veredelungs Größe":      $w('#V2size').value,
    "2. Genaue Motiv Größe in mm": $w('#V2mmsize').value,
    "2. Veredelungs Motiv":      $w('#V2upload').value,
    "2. Siebdruckfarben Anzahl": $w('#V2druckfarbenamount').value,
    "2. Siebdruck Schmuckfarbe": $w('#V2schmuckfarbe').value,

    // V3
    "3. Veredelungs Art":        $w('#V3art').value,
    "3. Veredelungsposition":    $w('#V3position').value,
    "3. Veredelungs Größe":      $w('#V3size').value,
    "3. Genaue Motiv Größe in mm": $w('#V3mmsize').value,
    "3. Veredelungs Motiv":      $w('#V3upload').value,
    "3. Siebdruckfarben Anzahl": $w('#V3druckfarbenamount').value,
    "3. Siebdruck Schmuckfarbe": $w('#V3schmuckfarbe').value,

    // V4
    "4. Veredelungs Art":        $w('#V4art').value,
    "4. Veredelungsposition":    $w('#V4position').value,
    "4. Veredelungs Größe":      $w('#V4size').value,
    "4. Genaue Motiv Größe in mm": $w('#V4mmsize').value,
    "4. Veredelungs Motiv":      $w('#V4upload').value,
    "4. Siebdruckfarben Anzahl": $w('#V4druckfarbenamount').value,
    "4. Siebdruck Schmuckfarbe": $w('#V4schmuckfarbe').value,

    // Sonstige Felder
    "Mock Up":                   $w('#Mockupupload').value,
    "T-Shirts":                  $w('#ddshirts').value,
    "Hoodie":                    $w('#ddhoodies').value,
    "Hose":                      $w('#ddhosen').value,
    "Mutze/Jacke/Socken":        $w('#ddstuff').value,
    "Textilfarbe":               $w('#inputtextilecolor').value,
    "Stückzahl":                 $w('#inputamount').value,

    "XS":                        $w('#inputXS').value,
    "S":                         $w('#inputS').value,
    "M":                         $w('#inputM').value,
    "L":                         $w('#inputL').value,
    "XL":                        $w('#inputXL').value,
    "XXL":                       $w('#inputXXL').value,
    "XXXL":                      $w('#inputXXXL').value,

    "Versandart":                $w('#ddversand').value
  };

  console.log('--- Feldwerte ---');
  Object.entries(data).forEach(([key, val]) => console.log(`${key}:`, val));
  console.log('--- Ende der Feldwerte ---');

  // --------- 3) Daten an Backend schicken ---------
  console.log('Sende Daten an Backend sendToSheet…');
  sendToSheet(data)
    .then(json => {
      if (json.result === 'success') {
        console.log('✅ Eintrag erfolgreich:', json.row);
      } else {
        console.error('🚨 Fehler vom Server:', json.error);
      }
    })
    .catch(err => console.error('🔥 Backend-Error:', err));
}


// ************************************************************************** //
// *************************** Hilfsfunktionen *************************** //
// ************************************************************************** //

function getTextil() {
  const keys = ['ddshirts','ddhoodies','ddhosen','ddstuff'];
  for (let id of keys) {
    const v = $w(`#${id}`).value;
    if (v) return v.toString().trim().toLowerCase();
  }
  return '';
}

function getMinutesForVeredelung(type, dimensionFull) {
  const dim = dimensionFull.toLowerCase();
  let sizeKey;
  if (dim.includes('klein dina6')) {
    sizeKey = 'KLEIN';
  } else if (dim.includes('medium dina4')) {
    sizeKey = 'MEDIUM';
  } else if (dim.includes('large dina3') || dim.includes('gross dina3')) {
    sizeKey = 'GROSS';
  } else {
    console.warn('Unbekannte Motivgröße:', dimensionFull);
    return 0;
  }

  switch (type.toUpperCase()) {
    case 'SIEBDRUCK':
      return sizeKey === 'KLEIN'  ? 0.6
           : sizeKey === 'MEDIUM' ? 0.75
           : /* GROSS */           1;
    case 'STICK':
      return sizeKey === 'KLEIN'  ? 2
           : sizeKey === 'MEDIUM' ? 4
           : /* GROSS */           6;
    case 'DTF':
      return sizeKey === 'KLEIN'  ? 1.25
           : sizeKey === 'MEDIUM' ? 2.25
           : /* GROSS */           2.8;
    default:
      console.warn('Unbekannter Verfahrenstyp:', type);
      return 0;
  }
}

function getVeredelungDiscount(menge) {
  if (menge >= 1000) return 0.50;
  if (menge >= 700)  return 0.40;
  if (menge >= 400)  return 0.30;
  if (menge >= 250)  return 0.20;
  if (menge >= 100)  return 0.10;
  if (menge >= 50)   return 0.05;
  return 0;
}

function getMaterialCost(textil, menge) {
  const prices = {
    'classicfit tee':          2.94,
    'royalcomfort tee':        3.72,
    'artisan tee':             5.75,
    'freeflow tee':           13.83,
    'urbanheavy tee':          8.09,
    'streetstyle tee':        16.43,
    'retrobox tee':           16.13,
    'everyday tee':           15.68,
    'cropchic tee':           14.78,
    'empireelite tee':        12.75,
    'empiremax tee':          15.00,
    'authentichoodie':        19.17,
    'ziphood classic':        23.40,
    'heritagesweat':          14.41,
    'boxfit hoodie':          33.92,
    'cruiserhoodie':          27.75,
    'streetstriker hoodie':   41.76,
    'ultrastreet hoodie':     74.85,
    'loosefit hoodie':        44.85,
    'standardhoodie':         43.35,
    'zipoversize hoodie':     50.85,
    'zipfit hoodie':          49.35,
    'cropcomfort hoodie':     38.85,
    'relaxedcrew sweater':    38.85,
    'classiccrew sweater':    36.90,
    'empirecrop heavy':       31.50,
    'cozyempire hoodie':      27.75,
    'empireessentials hoodie':22.50,
    'empirechill hoodie':     22.50,
    'imperiallong tee':        6.96,
    'shufflestyle tee':       10.10,
    'loosefit long sleeve':   20.78,
    'trainershorts':          20.44,
    'easyfit shorts':         34.95,
    'empirejogger shorts':    21.00,
    'authenticjoggers':       21.63,
    'comfortsweats':          43.35,
    'empirelounge pants':     21.00,
    'moveflex pants':         27.50,
    'snapbackoriginal':        9.39,
    'lowkey cap':             10.83,
    'twotone trucker':         6.51,
    'heavywarm beanie':        5.63,
    'harbourbeanie':           4.29,
    'crewstep socks':          2.79,
    'courtclassic socks':      2.85,
    'bomberedge jacket':      31.20,
    'coachpro jacket':        27.12,
    'collegevibe jacket':     35.27,
    'sols regent':             2.12,
    'keins':                   0.00
  };
  let base = prices[textil] || 50;
  let rabatt = 0;
  if (menge >= 500)      rabatt = 0.15;
  else if (menge >= 250) rabatt = 0.12;
  else if (menge >= 100) rabatt = 0.09;
  else if (menge >= 50)  rabatt = 0.06;
  else if (menge >= 25)  rabatt = 0.03;
  return base * (1 - rabatt);
}
