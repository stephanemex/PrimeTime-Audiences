document.addEventListener("DOMContentLoaded", () => {
    // Sélection des éléments DOM nécessaires
    const historiqueInput = document.getElementById("historiqueFile");
    const audienceInput = document.getElementById("audienceFile");
    const processAudienceButton = document.getElementById("processAudienceButton");
    const findMatchesButton = document.getElementById("findMatchesButton");
    const generatePreviewButton = document.getElementById("generatePreviewButton");
    const downloadButton = document.getElementById("downloadButton");
    const progressBar = document.getElementById("progressBar");
    const previewContainer = document.getElementById("previewContainer");

    // Variables pour stocker les données
    let historiqueData = null; // Données extraites du fichier historique
    let audienceData = {}; // Données extraites des différentes feuilles du fichier audience
    let updatedWorkbook = null; // Workbook Excel mis à jour avec la coloration

    /**
     * Étape 1 : Chargement du fichier historique
     */
    historiqueInput.addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (!file) {
            alert("Veuillez sélectionner un fichier historique.");
            return;
        }

        readExcel(file, (workbook) => {
            historiqueData = parseHistorique(workbook);
            console.log("Fichier historique chargé.");
            if (audienceData) {
                processAudienceButton.disabled = false; // Activer le bouton de traitement si l'autre fichier est chargé
            }
        });
    });

    /**
     * Étape 2 : Chargement du fichier audience
     */
    audienceInput.addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (!file) {
            alert("Veuillez sélectionner un fichier audience.");
            return;
        }
    
        readExcel(file, (workbook) => {
            audienceData = parseAudience(workbook);
            console.log("Données du fichier audience : ", audienceData); // Vérifiez ici
            console.log("Fichier audience chargé.");
            if (historiqueData) {
                processAudienceButton.disabled = false; // Activer le bouton de traitement si l'autre fichier est chargé
            }
        });
    });
    
    /**
     * Étape 3 : Traitement des données d'audience
     */
    processAudienceButton.addEventListener("click", () => {
        if (!audienceData) {
            alert("Veuillez charger un fichier audience avant de le traiter.");
            return;
        }

        progressBar.textContent = "Traitement des données d'audience en cours...";
        analyzeData(); // Préparation des données d'audience
        console.log("Traitement des données d'audience terminé.");
        progressBar.textContent = "Traitement terminé.";
        findMatchesButton.disabled = false; // Activer le bouton pour lancer la recherche de correspondances
    });

    /**
     * Étape 4 : Recherche des correspondances
     */
    findMatchesButton.addEventListener("click", () => {
        if (!historiqueData || !audienceData) {
            alert("Veuillez charger les fichiers avant de lancer la recherche.");
            return;
        }

        progressBar.textContent = "Recherche de correspondances en cours...";
        findMatches(); // Recherche les correspondances
        console.log("Recherche de correspondances terminée.");
        progressBar.textContent = "Correspondances trouvées.";
        generatePreviewButton.disabled = false; // Activer le bouton pour générer l'aperçu
    });

    /**
     * Étape 5 : Génération de l'aperçu
     */
    generatePreviewButton.addEventListener("click", () => {
        progressBar.textContent = "Génération de l'aperçu...";
        renderPreview(); // Génération de l'aperçu
        console.log("Aperçu généré.");
        progressBar.textContent = "Aperçu prêt.";
        downloadButton.disabled = false; // Activer le bouton pour télécharger
    });

    /**
     * Étape 6 : Téléchargement du fichier mis à jour
     */
    downloadButton.addEventListener("click", () => {
        if (!updatedWorkbook) {
            alert("Aucun fichier mis à jour à télécharger.");
            return;
        }

        XLSX.writeFile(updatedWorkbook, "updated_audience.xlsx");
        console.log("Fichier mis à jour téléchargé.");
    });

    // Fonction pour analyser les données d'audience (à ajouter si absente)
    function analyzeData() {
        console.log("Analyzing audience data...");
        Object.keys(audienceData).forEach((sheetName) => {
            if (sheetName === "Job") return; // Ignorer la feuille "Job"
            const sheet = audienceData[sheetName];
            sheet.forEach((row, rowIndex) => {
                if (rowIndex >= 2) { // Ignorer les lignes d'en-tête
                    const timeSlot = row[1]; // Colonne B : tranches horaires
                    if (timeSlot && isValidTimeSlot(timeSlot)) {
                        const [startTime, endTime] = timeSlot.split(" - ");
                        row.start = convertTimeStringToDate(startTime.trim());
                        row.end = convertTimeStringToDate(endTime.trim());
                    }
                }
            });
        });
        console.log("Audience data analyzed:", audienceData);
    }

    /**
     * Fonction pour rechercher les correspondances entre données historiques et audiences
     */
    function findMatches() {
        console.log("Recherche de correspondances...");
        const sheetNames = Object.keys(audienceData).filter(name => name !== "Job");

        sheetNames.forEach(sheetName => {
            const sheet = audienceData[sheetName];
            const datesRow = sheet[1]?.slice(2); // Ligne 2 : Dates dans les colonnes >= C

            historiqueData.forEach(program => {
                const startProgrammeDate = program["Date début"];
                const startProgrammeTime = program["Heure début"];

                console.log(`Programme : ${startProgrammeDate} - ${startProgrammeTime}`);

                // Trouver la colonne de la date dans le fichier audience
                const columnIndex = datesRow?.findIndex(date => date.trim() === startProgrammeDate.trim());
                if (columnIndex === -1 || columnIndex === undefined) {
                    console.warn(`Date non trouvée dans l'audience : ${startProgrammeDate}`);
                    return;
                }

                // Vérifier chaque tranche horaire
                sheet.forEach((row, rowIndex) => {
                    if (rowIndex >= 2) { // Ignorer les lignes d'en-tête
                        const timeSlot = row[1]; // Colonne B : tranches horaires
                        const timeSlotRegex = /^(\d{2}:\d{2}:\d{2}) - (\d{2}:\d{2}:\d{2})$/;
                        const match = timeSlot.match(timeSlotRegex);

                        if (!match) return; // Ignorer les lignes sans tranches horaires

                        const [_, startTimeStr, endTimeStr] = match;
                        const startTime = convertTimeStringToDate(startTimeStr);
                        const endTime = convertTimeStringToDate(endTimeStr);
                        const programmeTime = convertTimeStringToDate(startProgrammeTime);

                        // Vérifier si startProgrammeTime est dans la tranche
                        if (programmeTime >= startTime && programmeTime < endTime) {
                            console.log(`Correspondance trouvée : ${startProgrammeDate} ${timeSlot}`);

                            // Ajouter un style rouge à la cellule correspondante
                            const cellIndex = columnIndex + 2; // Décalage pour les colonnes ignorées
                            row[cellIndex] = `<span style="background-color: red;">${row[cellIndex]}</span>`;
                        }
                    }
                });
            });
        });

        console.log("Recherche de correspondances terminée.");
    }
});

/**
 * Fonction pour afficher un aperçu des données modifiées dans une table HTML
 * @param {string} selectedSheetName - Nom de la feuille à afficher (par défaut : la première feuille)
 */
function renderPreview(audienceData, selectedSheetName = null) {
    if (!audienceData || Object.keys(audienceData).length === 0) {
        console.error("Audience data is empty or undefined.");
        return;
    }

    console.log("Génération de l'aperçu...");

    const sheetNames = Object.keys(audienceData).filter(name => name !== "Job");
    if (!sheetNames.length) {
        alert("Aucune donnée disponible pour la prévisualisation.");
        return;
    }

    // Détecter la feuille sélectionnée ou utiliser la première par défaut
    const sheetName = selectedSheetName || sheetNames[0];
    const sheetData = audienceData[sheetName];
    console.log(`Données de la feuille ${sheetName} pour prévisualisation :`, sheetData);

    // Récupérer les dates de la ligne 2 (à partir de la colonne C)
    const dates = sheetData[1]?.slice(2); // Ligne 2, colonnes à partir de C

    // Mettre à jour ou créer le sélecteur de feuille
    const sheetSelector = document.getElementById("sheetSelector");
    sheetSelector.innerHTML = ""; // Réinitialiser le sélecteur
    sheetNames.forEach(name => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        if (name === sheetName) option.selected = true;
        sheetSelector.appendChild(option);
    });

    // Écouteur pour changer de feuille
    sheetSelector.addEventListener("change", () => {
        renderPreview(sheetSelector.value);
    });

    // Création du tableau HTML
    const table = document.createElement("table");
    const tbody = document.createElement("tbody");

    // Parcours de chaque ligne du tableau audience
    sheetData.forEach((row, rowIndex) => {
        const tr = document.createElement("tr");

        row.forEach((cell, colIndex) => {
            const td = document.createElement("td");
            td.innerHTML = cell || "";

            // Si c'est une cellule d'audience, colorer si elle correspond
            if (colIndex >= 2 && rowIndex >= 2) { // Colonnes >= C et lignes >= 3
                const audienceDate = dates?.[colIndex - 2]; // Date associée à la colonne
                const timeSlot = sheetData[rowIndex]?.[1]; // Tranche horaire (colonne B)

                if (audienceDate && timeSlot) {
                    const match = findMatchesForCell(audienceDate, timeSlot); // Utilisation d'une fonction de recherche
                    if (match) {
                        td.style.backgroundColor = "rgba(255, 0, 0, 0.5)"; // Rouge transparent
                        td.style.color = "white"; // Texte en blanc pour lisibilité
                    }
                }
            }

            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    const previewContainer = document.getElementById("previewContainer");
    previewContainer.innerHTML = ""; // Réinitialiser l'aperçu précédent
    previewContainer.appendChild(table);

    console.log("Aperçu généré.");
}
