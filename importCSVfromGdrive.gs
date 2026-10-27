// --- KONFIGURACIJA ZA SHEET3 (Zbirna tablica) ---
const CIJENA_BENZINA = 1.74; // EUR po litri (Prilagodi po potrebi)
const CIJENA_STRUJE = 0.08;  // EUR po kWh (Prilagodi po potrebi)
const BAZNA_POTROSNJA = 8; // L/100km (Potrošnja čistog benzinca)
const DOWNSAMPLE_SHEET1 = 10; // U Sheet1 upisuje svaki 10. redak

function importCSVsFromDrive() {
  const folderId = "1p6zbVZ9HuTIC2E3m6kGoakYWYj7VZqaN"; // UNESI SVOJ FOLDER ID
  const sheetName1 = "Sheet1";
  const sheetName3 = "Sheet3";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName1); 
  
  if (!sheet) {
    Logger.log("Error: Could not find a tab named '" + sheetName1 + "'");
    return;
  }

  // Provjera/Kreiranje Sheet3
  let sheet3 = ss.getSheetByName(sheetName3);
  if (!sheet3) {
    sheet3 = ss.insertSheet(sheetName3);
  }

  // 1. Target CSV headers from Car Scanner
  const targetHeaders = [
    "time",
    "Distance travelled (km)",
    "Fuel used (L)",
    "Energy used (kWh)",
    "Engine RPM (rpm)",
    "[BMS] State of Charge Display (%)",
    "Vehicle speed (km/h)",
    "Latitude",
    "Longtitude" // Handles Car Scanner spelling
  ];

  // 2. Standard output headers for Google Sheets (Sheet1)
  const outputHeader = [
    "DateTime",
    "Distance travelled (km)",
    "Fuel used (L)",
    "Energy used (kWh)",
    "Engine RPM (rpm)",
    "[BMS] State of Charge Display (%)",
    "Vehicle speed (km/h)",
    "Latitude",
    "Longitude"
  ];

  // Standard output headers for Sheet3
  const sheet3Headers = [
    "Datum i Vrijeme Vožnje",
    "Prođena udaljenost (km)",
    "Prosječna brzina (km/h)",
    "Potrošnja struje (kWh)",
    "SOC Energy used (kWh)",
    "Potrošnja benzina (L)",
    "Trošak struje (€)",
    "Trošak benzina (€)",
    "Ušteda benzina (L)",
    "Ušteda vs. Benzin (€)"
  ];

  // Setup headers if sheets are empty
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]);
  }
  if (sheet3.getLastRow() === 0) {
    sheet3.getRange(1, 1, 1, sheet3Headers.length).setValues([sheet3Headers]).setFontWeight("bold");
  }

  // 3. Load tracking lists from Script Properties
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const processedNamesJson = scriptProperties.getProperty("PROCESSED_FILE_NAMES") || "[]";
  let processedNames = JSON.parse(processedNamesJson);

  const emptyCheckJson = scriptProperties.getProperty("EMPTY_CHECK_FILES") || "[]";
  let emptyCheckFiles = JSON.parse(emptyCheckJson);

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles(); 
  
  let newFilesToProcess = [];

  // 4. Collect CSV files whose filename has NOT been processed
  while (files.hasNext()) {
    let file = files.next();
    let rawName = file.getName();
    let cleanName = rawName.trim();

    if (cleanName.toLowerCase().indexOf('.csv') > -1 && processedNames.indexOf(cleanName) === -1) {
      newFilesToProcess.push({
        file: file,
        name: cleanName
      });
    }
  }

  // 5. Sort incoming files DESCENDING by name (Newest trip first)
  newFilesToProcess.sort((a, b) => b.name.localeCompare(a.name));

  let allNewRows = [];
  let allNewSummaryRows = []; // Kontejner za Sheet3 rezultate
  let newlyProcessedFiles = [];
  let skippedEmptyFiles = [];

  // Pomoćna funkcija za siguran izračun razlike (Max - Min) iz niza
  function getDelta(arr) {
    if (arr.length === 0) return 0;
    let max = arr.reduce((a, b) => Math.max(a, b), -Infinity);
    let min = arr.reduce((a, b) => Math.min(a, b), Infinity);
    let delta = max - min;
    return delta > 0.0001 ? delta : 0;
  }

  // 6. Process files
  for (let f = 0; f < newFilesToProcess.length; f++) {
    let fileObj = newFilesToProcess[f];
    let file = fileObj.file;
    let fileName = fileObj.name;

    let rawText = file.getBlob().getDataAsString();
    let isDataEmpty = !rawText || rawText.trim().length === 0;
    let csvData = [];

    if (!isDataEmpty) {
      let delimiter = rawText.indexOf(';') > -1 ? ';' : ',';
      csvData = Utilities.parseCsv(rawText, delimiter);
      if (!csvData || csvData.length < 2) {
        isDataEmpty = true;
      }
    }

    // Handle empty or incomplete CSV files
    if (isDataEmpty) {
      let emptyIdx = emptyCheckFiles.indexOf(fileName);
      if (emptyIdx === -1) {
        emptyCheckFiles.push(fileName);
        Logger.log("File empty/syncing on first check. Pending re-verification next run: " + fileName);
      } else {
        Logger.log("Not processed: " + fileName);
        skippedEmptyFiles.push(fileName);
        
        processedNames.push(fileName);
        emptyCheckFiles.splice(emptyIdx, 1);
      }
      continue;
    }

    let prevEmptyIdx = emptyCheckFiles.indexOf(fileName);
    if (prevEmptyIdx !== -1) {
      emptyCheckFiles.splice(prevEmptyIdx, 1);
    }

    // Extract YYYY-MM-DD from filename
    let dateMatch = fileName.match(/(\d{4}-\d{2}-\d{2})/);
    let dateStr = dateMatch ? dateMatch[1] : "";

    let headerRow = csvData[0];

    // Match column positions
    let colIndices = targetHeaders.map(target => {
      let index = headerRow.indexOf(target);
      if (index === -1 && target === "Longtitude") {
        index = headerRow.indexOf("Longitude");
      }
      return index;
    });

    let tripRowsAll = [];     // Svi redovi (za precizan izračun u Sheet3)
    let tripRowsSheet1 = [];  // Filtrirani redovi s downsamplingom (za Sheet1)

// Extract data rows
    for (let i = 1; i < csvData.length; i++) {
      let sourceRow = csvData[i];
      let newRow = [];

      for (let j = 0; j < targetHeaders.length; j++) {
        let colIdx = colIndices[j];
        let val = (colIdx !== -1 && sourceRow[colIdx] !== undefined) ? sourceRow[colIdx] : "";

        // Combine Date and Time
        if (j === 0 && val !== "") {
          if (dateStr) {
            let cleanTime = String(val).trim();
            let fullDateTimeStr = dateStr + "T" + cleanTime;
            let parsedDate = new Date(fullDateTimeStr);
            val = !isNaN(parsedDate.getTime()) ? parsedDate : (dateStr + " " + cleanTime);
          }
        } else if (j > 0 && val !== "") {
          let numVal = Number(String(val).replace(',', '.'));
          if (!isNaN(numVal)) {
            val = numVal;
          }
        }

        // Dodavanje vrijednosti u trenutni red
        newRow.push(val);
      } // <-- KRAJ UNUTARNJE PETLJE (j)

      // Spremanje cijelog reda u glavni niz za izračun
      tripRowsAll.push(newRow);

      // Downsampling za Sheet1: uzimamo prvi, zadnji i svaki 10. redak
      if (i === 1 || i % DOWNSAMPLE_SHEET1 === 0 || i === csvData.length - 1) {
        tripRowsSheet1.push(newRow);
      }
    } // <-- KRAJ VANJSKE PETLJE (i)


// --- PROVJERA UDALJENOSTI I IZRAČUN STATISTIKE ZA SHEET3 (IZ SVIH REDOVA) ---
    if (tripRowsAll.length > 0) {
      let dist = getDelta(tripRowsAll.map(r => r[1]).filter(v => typeof v === 'number'));

      if (dist <= 0.0001) {
        Logger.log("Preskačem zapis s 0 km udaljenosti: " + fileName);
        processedNames.push(fileName);
        continue;
      }

      let fuel = getDelta(tripRowsAll.map(r => r[2]).filter(v => typeof v === 'number'));
      let energy = getDelta(tripRowsAll.map(r => r[3]).filter(v => typeof v === 'number'));

      // RPM Pravilo
      let rpmVals = tripRowsAll.map(r => r[4]).filter(v => typeof v === 'number');
      let maxRpm = rpmVals.length > 0 ? Math.max(...rpmVals) : 0;
      if (maxRpm <= 0) {
        fuel = 0;
      }

      // SOC Energy used
      let socVals = tripRowsAll.map(r => r[5]).filter(v => typeof v === 'number');
      let socDelta = getDelta(socVals);
      let socEnergy = socDelta * 0.12;

      let speedVals = tripRowsAll.map(r => r[6]).filter(v => typeof v === 'number' && v > 0);
      let avgSpeed = speedVals.length > 0 ? speedVals.reduce((a,b) => a + b, 0) / speedVals.length : 0;
      
      let tripStart = tripRowsAll[0][0]; // Datum i vrijeme prve točke mjerenja
      
      // Troškovi i uštede (clamped na min 0)
      let costElec = energy * CIJENA_STRUJE;
      let costGas = fuel * CIJENA_BENZINA;
      let totalPhevCost = costElec + costGas;

      let hypotheticalIceGasLiters = (dist * BAZNA_POTROSNJA) / 100.0;
      let hypotheticalIceCost = hypotheticalIceGasLiters * CIJENA_BENZINA;

      let gasSavedLiters = Math.max(0, hypotheticalIceGasLiters - fuel);
      let eurSaved = Math.max(0, hypotheticalIceCost - totalPhevCost);

      // Dodavanje u kontejner za Sheet3 (10 kolona: A-J)
      allNewSummaryRows.push([
        tripStart,
        Number(dist.toFixed(3)),
        Number(avgSpeed.toFixed(1)),
        Number(energy.toFixed(3)),
        Number(socEnergy.toFixed(3)),
        Number(fuel.toFixed(3)),
        Number(costElec.toFixed(2)),
        Number(costGas.toFixed(2)),
        Number(gasSavedLiters.toFixed(3)),
        Number(eurSaved.toFixed(2))
      ]);
    }

if (tripRowsSheet1.length > 0) {
      tripRowsSheet1.reverse();
      let emptyRow = new Array(outputHeader.length).fill("");
      tripRowsSheet1.push(emptyRow);
      allNewRows = allNewRows.concat(tripRowsSheet1);
    }

    // Mark as processed
    processedNames.push(fileName);
    newlyProcessedFiles.push(fileName);
  }

  // Save updated state back to Script Properties
  scriptProperties.setProperty("PROCESSED_FILE_NAMES", JSON.stringify(processedNames));
  scriptProperties.setProperty("EMPTY_CHECK_FILES", JSON.stringify(emptyCheckFiles));

  // 7. Insert all new rows directly at Row 2 in Sheet1 (pushing older data down)
  if (allNewRows.length > 0) {
    let hasExistingData = sheet.getLastRow() > 1;

    if (hasExistingData) {
      sheet.insertRowsBefore(2, allNewRows.length);
    }

    let range = sheet.getRange(2, 1, allNewRows.length, outputHeader.length);
    range.setValues(allNewRows);

    // Format Column A as DateTime
    sheet.getRange(2, 1, allNewRows.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss.sss");

    Logger.log("--- NEWLY PROCESSED FILES (" + newlyProcessedFiles.length + ") ---");
    newlyProcessedFiles.forEach((name, index) => {
      Logger.log((index + 1) + ". " + name);
    });
    Logger.log("Successfully inserted " + allNewRows.length + " rows at Row 2 in Sheet1.");
  } else {
    Logger.log("No new CSV files ready to process.");
  }

  // 8. Insert summary data into Sheet3 directly at Row 2 (columns A-J)
  if (allNewSummaryRows.length > 0) {
    let hasExistingSummaryData = sheet3.getLastRow() > 1;
    
    if (hasExistingSummaryData) {
      sheet3.insertRowsBefore(2, allNewSummaryRows.length);
    }
    
    let range3 = sheet3.getRange(2, 1, allNewSummaryRows.length, sheet3Headers.length);
    range3.setValues(allNewSummaryRows);
    
    // Format Column A in Sheet3 as standard DateTime
    sheet3.getRange(2, 1, allNewSummaryRows.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    Logger.log("Successfully inserted " + allNewSummaryRows.length + " summary rows at Row 2 in Sheet3.");
  }

  // 9. OSVJEŽAVAMO I ČISTIMO ZBIRNU TABLICU OD STUPCA L (L1:U3)
  osvjeziZbirnuTablicu(sheet3);

  if (skippedEmptyFiles.length > 0) {
    Logger.log("--- CONFIRMED EMPTY / NOT PROCESSED (" + skippedEmptyFiles.length + ") ---");
    skippedEmptyFiles.forEach((name) => {
      Logger.log("Not processed: " + name);
    });
  }
}

/**
 * Pomoćna funkcija: Postavlja, oblikuje i čisti zbirnu tablicu (PROSJEK i UKUPNO) od stupca L
 */
function osvjeziZbirnuTablicu(sheet3) {
  const headersSummary = [
    "Prosjek/Ukupno",
    "Prođena udaljenost (km)",
    "Prosječna brzina (km/h)",
    "Potrošnja struje (kWh)",
    "SOC Energy used (kWh)",
    "Potrošnja benzina (L)",
    "Trošak struje (€)",
    "Trošak benzina (€)",
    "Ušteda benzina (L)",
    "Ušteda vs. Benzin (€)"
  ];

  // Redak 2: PROSJEK
  const rowProsjekData = [
    "PROSJEK",
    '=AVERAGE(INDIRECT("B2:B"))',
    '=AVERAGE(INDIRECT("C2:C"))',
    '=AVERAGE(INDIRECT("D2:D"))',
    '=AVERAGE(INDIRECT("E2:E"))',
    '=AVERAGE(INDIRECT("F2:F"))',
    '=AVERAGE(INDIRECT("G2:G"))',
    '=AVERAGE(INDIRECT("H2:H"))',
    '=AVERAGE(INDIRECT("I2:I"))',
    '=AVERAGE(INDIRECT("J2:J"))'
  ];

  // Redak 3: UKUPNO
  const rowUkupnoData = [
    "UKUPNO",
    '=SUM(INDIRECT("B2:B"))',
    '=AVERAGE(INDIRECT("C2:C"))', 
    '=SUM(INDIRECT("D2:D"))',
    '=SUM(INDIRECT("E2:E"))',
    '=SUM(INDIRECT("F2:F"))',
    '=SUM(INDIRECT("G2:G"))',
    '=SUM(INDIRECT("H2:H"))',
    '=SUM(INDIRECT("I2:I"))',
    '=SUM(INDIRECT("J2:J"))'
  ];

  let maxRows = sheet3.getMaxRows();
  
  // 1. Čišćenje starih tablica/podataka od stupca K (11) do U (21) kako bi se spriječilo preklapanje
  if (maxRows >= 1) {
    sheet3.getRange(1, 11, maxRows, 11).clearContent().clearFormat();
  }

  // 2. Upisivanje zaglavlja i formula zbirne tablice od stupca L (12)
  sheet3.getRange(1, 12, 1, headersSummary.length).setValues([headersSummary]).setFontWeight("bold");
  sheet3.getRange(2, 12, 1, rowProsjekData.length).setValues([rowProsjekData]);
  sheet3.getRange(3, 12, 1, rowUkupnoData.length).setValues([rowUkupnoData]);

  // 3. Oblikovanje izgleda za L2:U3
  sheet3.getRange(2, 12, 2, 1).setFontWeight("bold"); // Podebljani nazivi PROSJEK i UKUPNO
  sheet3.getRange(2, 13, 2, 9).setNumberFormat("0.00"); // 2 decimalna mjesta
}

/** Reset function - clear memory cache */
function resetProcessedFilesHistory() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("PROCESSED_FILE_NAMES");
  props.deleteProperty("EMPTY_CHECK_FILES");
  props.deleteProperty("PROCESSED_FILE_KEYS");
  props.deleteProperty("PROCESSED_FILE_IDS");
  Logger.log("Processed file history and empty-check tracking cleared.");
}
