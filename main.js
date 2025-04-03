// Function to send emails
function sendConfirmationEmails(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues(); // Récupère toutes les données
    var headers = data[0]; // Première ligne (en-têtes)

    // Déterminer l'index des colonnes
    var colEmail = headers.indexOf("Adresse e-mail");
    var colStatut = headers.indexOf("Validée");
    var colMailEnvoye = headers.indexOf("Mail envoyé");
    var colJour = headers.indexOf("Jour");
    var colGrossesFraises =  headers.indexOf("Grosses fraises");
    var colFraisesConfiture = headers.indexOf("C");

    if (colEmail === -1 || colStatut === -1 || colMailEnvoye === -1) {
        SpreadsheetApp.getUi().alert("Erreur : Vérifie que les colonnes 'Adresse e-mail', 'Validée' et 'Mail envoyé' existent.");
        return;
    }

    var emailsEnvoyes = 0;

    for (var i = 1; i < data.length; i++) {
        var email = data[i][colEmail].trim(); // Supprime les espaces éventuels
        var statut = data[i][colStatut];
        var mailEnvoye = data[i][colMailEnvoye];
        var jour = Utilities.formatDate(data[i][colJour], Session.getScriptTimeZone(), "dd/MM/yyyy");

        // Vérifier si le mail doit être envoyé
        if ((statut === "Validée" || statut === "Refusée") && mailEnvoye === "" && email!=="") {
            var sujet = "";
            var message = "";
            var produitsCommandes = [];

            for (var j = colGrossesFraises; j <= colFraisesConfiture; j++) {
                var quantite = data[i][j];
                var nomProduit = headers[j]; // Nom du produit dans l'en-tête
                if (quantite && quantite > 0) {
                    produitsCommandes.push("- " + quantite + " " + nomProduit + "\n");
                }
            }

            if (statut === "Validée") {
                sujet = "Validation de votre commande pour le " + jour + "."
                message = "Bonjour,\n\nNous avons le plaisir de vous informer que votre commande pour le " + jour + " a été validée. Voici un récapitulatif de votre commande :\n" + produitsCommandes +"\n\nCordialement,\nLa ferme des grands prés";
            } else if (statut === "Refusée") {
                sujet = "Refus de votre commande pour le " + jour + "."
                message = "Bonjour,\n\nNous sommes désolés de vous informer que votre commande pour le " + jour + " a été refusée car nous n'avons pas assez de production en ce moment.\n\nCordialement,\nLa ferme des grands prés";
            }

            // Envoyer l'email
            MailApp.sendEmail(email, sujet, message);
            emailsEnvoyes++;

            // Mettre à jour la colonne "Mail envoyé"
            sheet.getRange(i + 1, colMailEnvoye + 1).setValue("Oui, automatiquement");
        }
    }

    SpreadsheetApp.getUi().alert(emailsEnvoyes + " e-mails envoyés.");
}

// Tasks to perform each time a sheet is edited manually
function onEdit(e) {
    //hideClient(e);
    updateDisplayPerDay();
  }
  
  // Tasks to perform each time a sheet structued is changed (e.g., dropdowns)
  //function onChange(e) {
  //  updateDisplayPerDay();
  //}
  
  // Tasks to perform each time a Google form injects data
  // function onFormSubmit(e) {
  //}
  
  function updateDisplayPerDay() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  
    const clientOrdersSheetName = 'commandes-clients';
    const adminOrdersSheetName = 'commandes-admin';
    const dashboardSheetName = 'dashboard';
  
    const dashboardSheet = ss.getSheetByName(dashboardSheetName);
    const startDate = new Date(dashboardSheet.getRange('C3').getValue());
    const endDate = new Date(dashboardSheet.getRange('C4').getValue());
  
    let ordersByDate = {};
    let otherOrders = [];
  
    // Process commandes-clients (needs validation check)
    processOrders(ss, clientOrdersSheetName, startDate, endDate, ordersByDate, otherOrders, true);
    
    // Process commandes-admin (always validated)
    processOrders(ss, adminOrdersSheetName, startDate, endDate, ordersByDate, otherOrders, false);
  
    // Create & update specific date tabs
    Object.keys(ordersByDate).forEach(tabName => {
      updateSheet(ss, tabName, ordersByDate[tabName]);
    });
  
    // Handle "autres" tab
    if (otherOrders.length > 0) {
      updateSheet(ss, "autres", otherOrders);
    }
  }
  
  function processOrders(ss, sheetName, startDate, endDate, ordersByDate, otherOrders, checkValidation) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet ${sheetName} not found!`);
      return;
    }
  
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
  
    const headers = data[0];
  
    let isClientSheet = (sheetName === "commandes-clients");
  
    // Get column indexes based on sheet type
    const dateIndex = headers.indexOf("Commande pour ...");
    const nameIndex = headers.indexOf("Nom");
    const emailIndex = headers.indexOf("Email");
    const statutIndex = isClientSheet ? headers.indexOf("Statut") : -1; // Only exists in commandes-clients
    const columnIndexes = {
      "R": headers.indexOf("R"),
      "P": headers.indexOf("P"),
      "B": headers.indexOf("B"),
      "F": headers.indexOf("F"),
      "Po": headers.indexOf("Po"),
      "Bac": headers.indexOf("Bac"),
      "S": headers.indexOf("S"),
      "C": headers.indexOf("C")
    };
  
    if (dateIndex === -1 || nameIndex === -1 || emailIndex === -1) {
      Logger.log(`Required columns not found in ${sheetName}!`);
      return;
    }
  
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (isClientSheet && row[statutIndex] !== "Validé") continue; // Skip if not validated
  
      const orderDate = new Date(row[dateIndex]);
      if (isNaN(orderDate)) continue;
  
      // Normalize data into the expected format
      let formattedRow = [
        row[emailIndex], // Email
        row[nameIndex], // Nom
        row[dateIndex], // Commande pour ...
        row[columnIndexes["R"]],
        row[columnIndexes["P"]],
        row[columnIndexes["B"]],
        row[columnIndexes["F"]],
        row[columnIndexes["Po"]],
        row[columnIndexes["Bac"]],
        isClientSheet ? "" : row[columnIndexes["S"]], // S (empty for clients)
        isClientSheet ? "" : row[columnIndexes["C"]]  // C (empty for clients)
      ];
  
      if (orderDate >= startDate && orderDate <= endDate) {
        const formattedTabName = formatTabName(orderDate);
        if (!ordersByDate[formattedTabName]) ordersByDate[formattedTabName] = [];
        ordersByDate[formattedTabName].push(formattedRow);
      } else {
        otherOrders.push(formattedRow);
      }
    }
  }
  
  function updateSheet(ss, sheetName, data) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = createDaySheet(ss, sheetName);
    sheet.clearContents();
  
    // Unified headers
    const headers = ["Email", "Nom", "Commande pour ...", "R", "P", "B", "F", "Po", "Bac", "S", "C"];
    sheet.appendRow(headers);
  
    // Sort data by name
    data.sort(compareRows);
  
    // Add sorted rows
    if (data.length > 0) {
      let dataRange = sheet.getRange(2, 1, data.length, headers.length);
      dataRange.setValues(data);
    }
  
    // Apply styles
    applySheetStyles(sheet, headers.length, data.length);
  }
  
  // Helper function to sort rows by date and name
  function compareRows(a, b) {
    const dateA = a[2] || new Date(0); // Default to very old date if missing
    const dateB = b[2] || new Date(0);
    const nameA = a[1] ? String(a[1]) : "";
    const nameB = b[1] ? String(b[1]) : "";
  
    Logger.log("Sorting: nameA = %s, nameB = %s", nameA, nameB);
  
    return dateA - dateB || nameA.localeCompare(nameB);
  }
  
  function formatTabName(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  
  
  
  // Hide clients when checkbox is selected
  function hideClient(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    
    // Column where checkboxes are (adjust the column index if necessary)
    var checkboxColumn = 1;  // Column A (1 = Column A, 2 = Column B, etc.)
  
    // Apply the filter when the edit is made in the checkbox column
    if (range.getColumn() == checkboxColumn) {
      var filter = sheet.getFilter();
      
      // If the filter exists, modify it
      if (filter) {
        var criteria = SpreadsheetApp.newFilterCriteria()
          .whenFormulaSatisfied('=NOT(A2=TRUE)')  // Formula to filter out checked rows (TRUE)
          .build();
        filter.setColumnFilterCriteria(checkboxColumn, criteria);  // Apply the filter to the checkbox column
      } else {
        // If no filter exists, create a new one with the criteria
        sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).createFilter();
        var newFilter = sheet.getFilter();
        var criteria = SpreadsheetApp.newFilterCriteria()
          .whenFormulaSatisfied('=NOT(A2=TRUE)')  // Formula to filter out checked rows (TRUE)
          .build();
        newFilter.setColumnFilterCriteria(checkboxColumn, criteria);  // Apply the filter to the checkbox column
      }
    }
  }
  
  // Utils functions
  /**
   * Format date into tab name (e.g., "ve07/03" for Friday, March 7)
   */
  function formatTabName(date) {
    const days = ["di", "lu", "ma", "me", "je", "ve", "sa"];
    const dayOfWeek = days[date.getDay()];
    const day = ("0" + date.getDate()).slice(-2);
    const month = ("0" + (date.getMonth() + 1)).slice(-2);
    return `${dayOfWeek}${day}/${month}`;
  }
  
  /**
   * Create a new sheet and return it
   */
  function createDaySheet(ss, sheetName) {
    let sheet = ss.insertSheet(sheetName);
    return sheet;
  }
  
  /**
   * Apply styling to the sheet: header, font size, alternating row colors
   */
  function applySheetStyles(sheet, numColumns, numRows) {
    // Set font size for the entire sheet
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontSize(14);
  
    // Style header: dark green background, white bold text
    let headerRange = sheet.getRange(1, 1, 1, numColumns);
    headerRange.setBackground("#0B6623").setFontColor("white").setFontWeight("bold");
  
    // Apply alternating row colors (light green for every other row)
    for (let i = 0; i < numRows; i++) {
      if (i % 2 === 0) {
        sheet.getRange(i + 2, 1, 1, numColumns).setBackground("#DFF2BF"); // Light green
      }
    }
  }
  
  