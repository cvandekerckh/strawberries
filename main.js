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
                message = "Bonjour,\n\nNous avons le plaisir de vous informer que votre commande pour le " + jour + " a été validée. Voici un récapitulatif de votre commande :\n" + produitsCommandes.join("") +"\n\nCordialement,\nLa ferme des grands prés";
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

function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  // Adjust this range to cover the columns you care about
  const lastCol = sheet.getLastColumn();

  // Set font size to 12 for the new row
  sheet.getRange(row, 1, 1, lastCol).setFontSize(12);
}
