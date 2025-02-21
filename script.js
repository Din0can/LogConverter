/************************************************
 * script.js
 * - Lädt eine .log-Datei im Browser
 * - Parst sie per Regex
 * - Erzeugt ein Excel-Workbook mit SheetJS
 * - Aktiviert in jedem Worksheet ein autoFilter
 *   (Filterzeile für Excel-Tabellen)
 ************************************************/

/***************************************
 * 1) Hilfsfunktionen zum Parsen (Regex)
 **************************************/

/**
 * Sucht nach Zeilen wie:
 *  User EMAIL | Frau/Herrn Nachname, Vorname included based on {kind}: DATETIME
 */
function parseIncludedUser(line, kind) {
    const pattern = new RegExp(
      `User\\s+(?<email>[^\\s]+)\\s*\\|\\s*(?<anrede>Frau|Herrn)\\s+(?<lastFirst>[^i]+)included based on ${kind}:\\s+(?<datetime>.+)`
    );
    const match = line.match(pattern);
    if (!match) return null;
  
    const { email, anrede, lastFirst, datetime } = match.groups;
    return {
      email: email.trim(),
      anrede: anrede.trim(),
      name: lastFirst.trim(),
      datum_zeit: datetime.trim(),
      kind: kind
    };
  }
  
  /**
   * Ignorierter Nutzer:
   *   Ignorierter Nutzer: EMAIL | Frau/Herrn Name | LETZTES_AENDERUNGSDATUM: XX, ...
   */
  function parseIgnoredUser(line) {
    const pattern = new RegExp(
      `Ignorierter Nutzer:\\s+(?<email>[^\\s]+)\\s*\\|\\s*(?<anrede>Frau|Herrn)\\s+(?<name>[^|]+)\\|.*?` +
      `LETZTES_AENDERUNGSDATUM:\\s*(?<lad>[^,]+)?,\\s*ERFASSUNGSZEITPUNKT:\\s*(?<ez>.*)`
    );
    const match = line.match(pattern);
    if (!match) return null;
  
    const { email, anrede, name, lad, ez } = match.groups;
    return {
      email: email.trim(),
      anrede: anrede.trim(),
      name: name.trim(),
      letztes_aenderungsdatum: lad ? lad.trim() : null,
      erfassungszeitpunkt: ez ? ez.trim() : null
    };
  }
  
  /**
   * Zeilen wie:
   *   Benutzer: EMAIL | Frau/Herrn Name
   */
  function parseBenutzer(line) {
    const pattern = new RegExp(
      `Benutzer:\\s+(?<email>[^\\s]+)\\s*\\|\\s*(?<anrede>Frau|Herrn)\\s+(?<name>.+)`
    );
    const match = line.match(pattern);
    if (!match) return null;
  
    const { email, anrede, name } = match.groups;
    return {
      email: email.trim(),
      anrede: anrede.trim(),
      name: name.trim()
    };
  }
  
  /**
   * Zeilen wie:
   *   Benutzer: EMAIL | DN: uid=...,ou=Benutzer,...
   */
  function parseUsersAddedToGroup(line) {
    const pattern = new RegExp(
      `Benutzer:\\s+(?<email>[^\\s]+)\\s*\\|\\s+DN:\\s+(?<dn>.*)`
    );
    const match = line.match(pattern);
    if (!match) return null;
  
    const { email, dn } = match.groups;
    return { email: email.trim(), dn: dn.trim() };
  }
  
  /***************************************
   * 2) Hauptfunktion zum Analysieren
   ***************************************/
  function analyzeLogfile(logText) {
    const data = {
      eingeschlossen_lad: [],       // Included (LETZTES_AENDERUNGSDATUM)
      eingeschlossen_ez: [],        // Included (ERFASSUNGSZEITPUNKT)
      ignorierte_nutzer: [],        // Ignored users
      mandanten_ein_nutzer: [],     // Single-user Mandanten
      mandanten_mehrere_nutzer: [], // Multi-user Mandanten
      benutzer_hinzugefuegt_hv: []  // Users added to HV group
    };
  
    let currentOrg = null;
    let currentGp = null;
  
    let inIgnoredSection = false;
    let inSingleSection = false;
    let inMultiSection = false;
  
    // In Zeilen splitten
    const lines = logText.split(/\r?\n/);
  
    for (let line of lines) {
      const trimmed = line.trim();
  
      // Erkennung der "Marker"
      if (line.includes("========= Included Users Based on LETZTES_AENDERUNGSDATUM =========")) {
        inIgnoredSection = false; inSingleSection = false; inMultiSection = false;
        continue;
      }
      if (line.includes("========= Included Users Based on ERFASSUNGSZEITPUNKT =========")) {
        inIgnoredSection = false; inSingleSection = false; inMultiSection = false;
        continue;
      }
      if (line.includes("========= Ignorierte Nutzer von Mandanten =========")) {
        inIgnoredSection = true; inSingleSection = false; inMultiSection = false;
        continue;
      }
      if (line.includes("========= Mandanten mit einem Benutzer =========")) {
        inIgnoredSection = false; inSingleSection = true; inMultiSection = false;
        continue;
      }
      if (line.includes("========= Mandanten mit mehreren Benutzern =========")) {
        inIgnoredSection = false; inSingleSection = false; inMultiSection = true;
        continue;
      }
      if (line.includes("========= Users Added to Hauptverantwortlicher Group =========")) {
        inIgnoredSection = false; inSingleSection = false; inMultiSection = false;
        continue;
      }
  
      // 1) Eingeschlossen (LAD)
      if (line.includes(" included based on LETZTES_AENDERUNGSDATUM:")) {
        const parsed = parseIncludedUser(line, "LETZTES_AENDERUNGSDATUM");
        if (parsed) data.eingeschlossen_lad.push(parsed);
        continue;
      }
  
      // 2) Eingeschlossen (EZ)
      if (line.includes(" included based on ERFASSUNGSZEITPUNKT:")) {
        const parsed = parseIncludedUser(line, "ERFASSUNGSZEITPUNKT");
        if (parsed) data.eingeschlossen_ez.push(parsed);
        continue;
      }
  
      // 3) Ignorierte Nutzer
      if (inIgnoredSection) {
        if (trimmed.startsWith("Organization:")) {
          currentOrg = trimmed.replace("Organization:", "").trim();
        } else if (line.includes("Ignorierter Nutzer:")) {
          const ignoredData = parseIgnoredUser(line);
          if (ignoredData) {
            ignoredData.organisation = currentOrg;
            data.ignorierte_nutzer.push(ignoredData);
          }
        }
        continue;
      }
  
      // 4) Mandanten (ein/mehrere Nutzer)
      if (inSingleSection) {
        if (trimmed.startsWith("Geschäftspartner:")) {
          currentGp = trimmed.replace("Geschäftspartner:", "").trim();
        } else if (trimmed.startsWith("Benutzer:")) {
          const benData = parseBenutzer(trimmed);
          if (benData) {
            benData.geschaeftspartner = currentGp;
            data.mandanten_ein_nutzer.push(benData);
          }
        }
        continue;
      }
      if (inMultiSection) {
        if (trimmed.startsWith("Geschäftspartner:")) {
          currentGp = trimmed.replace("Geschäftspartner:", "").trim();
        } else if (trimmed.startsWith("Benutzer:")) {
          const benData = parseBenutzer(trimmed);
          if (benData) {
            benData.geschaeftspartner = currentGp;
            data.mandanten_mehrere_nutzer.push(benData);
          }
        }
        continue;
      }
  
      // 5) HV-Gruppe
      if (trimmed.startsWith("Benutzer:") && trimmed.includes(" DN: ")) {
        const addedData = parseUsersAddedToGroup(trimmed);
        if (addedData) data.benutzer_hinzugefuegt_hv.push(addedData);
        continue;
      }
    }
  
    return data;
  }
  
  /***************************************
   * 3) Excel-Generierung mit SheetJS (+ Filter)
   ***************************************/
  function createWorkbook(data) {
    const wb = XLSX.utils.book_new();
  
    /**
     * Erstellt Worksheet aus 2D-Array (headers + rows),
     * hängt es ans Workbook und setzt 'autoFilter',
     * damit Excel die Filter in der ersten Zeile anzeigt.
     */
    function appendSheet(sheetName, headers, rowData) {
      // AOA = "array of arrays": Erste Zeile = Header, gefolgt von Daten
      const allRows = [headers, ...rowData];
      const ws = XLSX.utils.aoa_to_sheet(allRows);
  
      // Worksheet zum Workbook hinzufügen
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
  
      // Falls es Datensätze gibt, aktivieren wir autoFilter
      if (rowData.length > 0) {
        // rowData.length Zeilen + 1 Headerzeile => total rowData.length + 1
        // columns = headers.length
        const lastRowIndex = rowData.length;    // 0-basierter Index => rowData.length
        const lastColIndex = headers.length - 1;
  
        // Range von A1 bis "entsprechende Spalte/Zeile"
        // z.B. A1:D10 => { s: { r:0, c:0}, e: { r:9, c:3} }
        const range = XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: lastRowIndex, c: lastColIndex }
        });
  
        // Hier autoFilter aktivieren, wodurch Excel-Filter gesetzt wird
        ws["!autofilter"] = { ref: range };
      }
    }
  
    // 1) Eingeschlossen (Letztes AD)
    {
      const sheetName = "Eingeschlossen_LetztesAD";
      const headers = ["Email", "Anrede", "Name", "Datum/Uhrzeit"];
      const rowData = data.eingeschlossen_lad.map(item => [
        item.email,
        item.anrede,
        item.name,
        item.datum_zeit
      ]);
      appendSheet(sheetName, headers, rowData);
    }
  
    // 2) Eingeschlossen (Erfassungszeitpunkt)
    {
      const sheetName = "Eingeschlossen_EZPunkt";
      const headers = ["Email", "Anrede", "Name", "Datum/Uhrzeit"];
      const rowData = data.eingeschlossen_ez.map(item => [
        item.email,
        item.anrede,
        item.name,
        item.datum_zeit
      ]);
      appendSheet(sheetName, headers, rowData);
    }
  
    // 3) Ignorierte Nutzer
    {
      const sheetName = "Ignorierte_Nutzer";
      const headers = [
        "Organisation",
        "Email",
        "Anrede",
        "Name",
        "LetztesÄnderungsdatum",
        "Erfassungszeitpunkt"
      ];
      const rowData = data.ignorierte_nutzer.map(item => [
        item.organisation || "",
        item.email,
        item.anrede,
        item.name,
        item.letztes_aenderungsdatum,
        item.erfassungszeitpunkt
      ]);
      appendSheet(sheetName, headers, rowData);
    }
  
    // 4) Mandanten mit einem Nutzer
    {
      const sheetName = "Mandanten_EinNutzer";
      const headers = ["Geschaeftspartner", "Email", "Anrede", "Name"];
      const rowData = data.mandanten_ein_nutzer.map(item => [
        item.geschaeftspartner,
        item.email,
        item.anrede,
        item.name
      ]);
      appendSheet(sheetName, headers, rowData);
    }
  
    // 5) Mandanten mit mehreren Nutzern
    {
      const sheetName = "Mandanten_MehrereNutzer";
      const headers = ["Geschaeftspartner", "Email", "Anrede", "Name"];
      const rowData = data.mandanten_mehrere_nutzer.map(item => [
        item.geschaeftspartner,
        item.email,
        item.anrede,
        item.name
      ]);
      appendSheet(sheetName, headers, rowData);
    }
  
    // 6) Nutzer in Hauptverantwortlicher-Gruppe
    {
      const sheetName = "HV_Gruppe_Hinzugefuegt";
      const headers = ["Email", "DN"];
      const rowData = data.benutzer_hinzugefuegt_hv.map(item => [
        item.email,
        item.dn
      ]);
      appendSheet(sheetName, headers, rowData);
    }
  
    return wb;
  }
  
  /***************************************
   * 4) Button-Handler: Datei lesen, parsen, Excel downloaden
   ***************************************/
  function parseAndDownload() {
    const fileInput = document.getElementById("logFile");
    const statusDiv = document.getElementById("statusMessage");
  
    if (!fileInput.files || fileInput.files.length === 0) {
      statusDiv.textContent = "Bitte zuerst eine .log-Datei auswählen!";
      return;
    }
  
    const file = fileInput.files[0];
    if (!file.name.toLowerCase().endsWith(".log")) {
      statusDiv.textContent = "Bitte eine gültige .log-Datei wählen (.log).";
      return;
    }
  
    statusDiv.textContent = "Lese Datei...";
  
    // Datei einlesen
    const reader = new FileReader();
    reader.onload = function(e) {
      const logText = e.target.result;
      statusDiv.textContent = "Analysiere...";
      // 1) Log parsen
      const data = analyzeLogfile(logText);
  
      // 2) Workbook erstellen
      statusDiv.textContent = "Erstelle Excel-Datei...";
      const wb = createWorkbook(data);
  
      // 3) Download als .xlsx
      statusDiv.textContent = "Starte Download...";
      XLSX.writeFile(wb, "log_analysis.xlsx");
  
      statusDiv.textContent = "Fertig! Excel-Datei wurde erstellt.";
    };
    reader.onerror = function() {
      statusDiv.textContent = "Fehler beim Lesen der Datei!";
    };
  
    reader.readAsText(file, "UTF-8");
  }
  