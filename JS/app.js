document.addEventListener("DOMContentLoaded", () => {
    // Sélection des éléments DOM nécessaires
    const historiqueInput = document.getElementById("historiqueFile");
    const audienceInput = document.getElementById("audienceFile");
    const analyzeButton = document.getElementById("analyzeButton");
    const progressBar = document.getElementById("progressBar").firstElementChild;
    const previewContainer = document.getElementById("previewContainer");
    const downloadButton = document.getElementById("downloadButton");
    const sheetSelector = document.getElementById("sheetSelector");

    // Variables pour stocker les données
    let historiqueData = null; // Données extraites du fichier historique
    let audienceData = {}; // Données extraites des différentes feuilles du fichier audience
    let updatedWorkbook = null; // Workbook Excel mis à jour avec la coloration

    /**
     * Fonction pour lire un fichier Excel et renvoyer le workbook via un callback.
     * @param {File} file - Le fichier Excel chargé par l'utilisateur.
     * @param {Function} callback - Fonction à appeler avec le workbook lu.
     */
    function readExcel(file, callback) {
        console.log(`Lecture du fichier : ${file.name}`);
        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            callback(workbook);
        };
        reader.readAsArrayBuffer(file);
    }

    /**
     * Fonction pour extraire les données du fichier historique à partir de la 5ème ligne.
     * @param {Object} workbook - Le workbook Excel chargé.
     * @returns {Array} - Un tableau d'objets représentant les données.
     */
    function parseHistorique(workbook) {
        console.log("Extraction des données du fichier historique...");
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { range: 4, raw: false }); // Ignorer les 4 premières lignes
        console.log("Données extraites du fichier historique :", data);
        return data;
    }

    /**
     * Fonction pour extraire les données de toutes les feuilles du fichier audience.
     * @param {Object} workbook - Le workbook Excel chargé.
     * @returns {Object} - Un objet contenant les données par feuille.
     */
    function parseAudience(workbook) {
        console.log("Extraction des données du fichier audience...");
        const sheets = {};
        workbook.SheetNames.forEach(sheetName => {
            console.log(`Lecture de la feuille : ${sheetName}`);
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Récupère toutes les lignes en tableau brut
            sheets[sheetName] = data; // Stocker les données pour traitement ultérieur
        });
        console.log("Toutes les données extraites :", sheets);
        return sheets;
    }

    /**
     * Fonction pour convertir une date et une heure en un objet Date.
     * @param {String} dateStr - Date au format "dd.mm.yyyy".
     * @param {String} timeStr - Heure au format "HH:MM:SS".
     * @returns {Date} - Objet Date représentant la combinaison date + heure.
     */
    function convertToDate(dateStr, timeStr) {
        const [day, month, year] = dateStr.split(".");
        const [hours, minutes, seconds] = timeStr.split(":");
        return new Date(year, month - 1, day, hours, minutes, seconds);
    }

    /**
     * Fonction pour convertir un horaire texte en objet Date.
     * @param {String} timeString - Heure au format "HH:MM:SS".
     * @returns {Date} - Objet Date représentant l'heure.
     */
    function convertTimeStringToDate(timeString) {
        const [hours, minutes, seconds] = timeString.split(":").map(Number);
        return new Date(1970, 0, 1, hours, minutes, seconds);
    }

    // Fonction utilitaire pour convertir des plages horaires
    function convertAudienceTimes(date, timeStr) {
        if (!timeStr || typeof timeStr !== "string") {
            console.error("Heure invalide ou non définie : ", timeStr);
            return null; // Retourner null si l'heure est invalide
        }
    
        const [hours, minutes, seconds] = timeStr.split(":").map(Number);
        const adjustedHours = hours >= 24 ? hours - 24 : hours; // Gérer les heures > 24
        const dayOffset = hours >= 24 ? 1 : 0; // Décalage d'un jour si nécessaire
    
        return new Date(
            new Date(date).getFullYear(),
            new Date(date).getMonth(),
            new Date(date).getDate() + dayOffset,
            adjustedHours,
            minutes,
            seconds || 0
        );
    }
    
    
    /**
     * Fonction pour vérifier si une cellule correspond à une plage horaire d'un programme.
     * @param {String} cell - Contenu de la cellule (plage horaire en texte).
     * @param {Number} rowIndex - Index de la ligne dans le tableau.
     * @param {Number} colIndex - Index de la colonne dans le tableau.
     * @returns {Boolean} - Vrai si la plage horaire correspond à un programme.
     */
    function isMatchingTimeSlot(audienceDate, timeSlot) {
        if (!timeSlot || typeof timeSlot !== "string") return false; // Si la cellule est vide ou n'est pas un texte
    
        // Vérifier si la tranche horaire est au format HH:MM:SS - HH:MM:SS
        const timeSlotRegex = /^(\d{2}:\d{2}:\d{2}) - (\d{2}:\d{2}:\d{2})$/;
        const match = timeSlot.match(timeSlotRegex);
        if (!match) return false;
    
        // Extraire les horaires de début et de fin
        const [_, startTime, endTime] = match;
    
        // Conversion en objets Date avec la date d'audience
        const start = convertAudienceTimes(audienceDate, startTime);
        const end = convertAudienceTimes(audienceDate, endTime);
    
        // Recherche de correspondance dans le fichier historique
        const matchFound = historiqueData.some(program => {
            const programStart = new Date(`${program['Date début']}T${program['Heure début']}`);
            const programEnd = new Date(`${program['Date fin']}T${program['Heure fin']}`);
    
            // Vérification du chevauchement entre le programme et la tranche horaire
            return (
                (programStart >= start && programStart < end) || // Début du programme dans la tranche
                (programEnd > start && programEnd <= end) ||    // Fin du programme dans la tranche
                (programStart <= start && programEnd >= end)    // Le programme englobe totalement la tranche
            );
        });
    
        console.log(
            matchFound
                ? `Correspondance trouvée : ${audienceDate} ${timeSlot}`
                : `Pas de correspondance pour : ${audienceDate} ${timeSlot}`
        );
    
        return matchFound;
    }
    
    
    /**
     * Fonction principale pour analyser les données historiques et audiences.
     * Colore les plages horaires correspondantes.
     */
    function analyzeData() {
        console.log("Début de l'analyse des données...");
        progressBar.style.width = "50%";
        progressBar.textContent = "Analyse en cours...";
    
        // Extraire et convertir les plages horaires des programmes
        const programs = historiqueData.map(row => {
            if (!row["Date début"] || !row["Heure début"] || !row["Heure fin"]) {
                console.warn("Ligne ignorée (données manquantes) :", row);
                return null;
            }
            return {
                date: row["Date début"].trim(),
                start: convertToDate(row["Date début"].trim(), row["Heure début"].trim()),
                end: convertToDate(row["Date fin"]?.trim() || row["Date début"].trim(), row["Heure fin"].trim()),
            };
        }).filter(program => program !== null); // Filtrer les lignes invalides
    
        console.log("Programmes extraits :", programs);
    
        if (programs.length === 0) {
            alert("Aucun programme valide trouvé dans le fichier historique.");
            return;
        }
    
        updatedWorkbook = XLSX.utils.book_new(); // Créer un nouveau workbook pour les données modifiées
    
        // Parcourir les feuilles d'audience
        Object.keys(audienceData).forEach(sheetName => {
            console.log(`Traitement de la feuille : ${sheetName}`);
            const sheet = audienceData[sheetName];
            const updatedSheet = sheet.map(row => [...row]); // Cloner la feuille pour modifications
    
            // Parcourir chaque cellule et vérifier les correspondances
            sheet.forEach((row, rowIndex) => {
                row.forEach((cell, colIndex) => {
                    if (typeof cell === "string" && cell.includes(" - ")) {
                        const [startStr, endStr] = cell.split(" - ");
                        const start = convertAudienceTimes(startStr.trim());
                        const end = convertAudienceTimes(endStr.trim());
    
                        // Vérification de correspondance entre les plages horaires
                        const match = programs.some(program => {
                            const programStart = program.start;
                            const programEnd = program.end;
    
                            // Correspondance si :
                            // 1. Le programme commence dans la tranche horaire
                            // 2. Le programme se termine dans la tranche horaire
                            // 3. La tranche horaire est complètement incluse dans le programme
                            return (
                                (programStart >= start && programStart < end) || // Début dans la tranche
                                (programEnd > start && programEnd <= end) || // Fin dans la tranche
                                (programStart <= start && programEnd >= end) // Tranche incluse dans le programme
                            );
                        });
    
                        if (match) {
                            console.log(
                                `Correspondance trouvée : cellule "${cell}" (ligne ${rowIndex}, colonne ${colIndex})`
                            );
                            updatedSheet[rowIndex][colIndex] = `<span style="background-color: red;">${cell}</span>`;
                        }
                    }
                });
            });

            // Ajouter la feuille modifiée au workbook
            const worksheet = XLSX.utils.aoa_to_sheet(updatedSheet);
            XLSX.utils.book_append_sheet(updatedWorkbook, worksheet, sheetName);
        });
    
        console.log("Analyse terminée. Aperçu prêt.");
        renderPreview();
    }
    
        //fonction Highlight
        function highlightRowByText(sheetData, searchText) {
            // Recherche de la ligne contenant le texte
            console.log(`Recherche de la ligne contenant : "${searchText}"`);
            sheetData.forEach((row, rowIndex) => {
                if (row.some(cell => typeof cell === "string" && cell.includes(searchText))) {
                    console.log(`Ligne trouvée à l'index ${rowIndex} :`, row);
        
                    // Marquer la ligne entière en jaune
                    row.forEach((cell, colIndex) => {
                        const tableRow = document.querySelector(`#previewContainer table tr:nth-child(${rowIndex + 1})`);
                        if (tableRow) {
                            const tableCell = tableRow.children[colIndex];
                            if (tableCell) {
                                tableCell.style.backgroundColor = "rgba(255, 255, 0, 0.5)"; // Jaune transparent
                            }
                        }
                    });
                }
            });
        }
    
    /**
     * Fonction pour afficher un aperçu des données modifiées dans une table HTML.
     */
    function renderPreview(selectedSheetName = null) {
        console.log("Génération du tableau de prévisualisation...");
        progressBar.style.width = "100%";
        progressBar.textContent = "Analyse terminée";
    
        // Ignorer la feuille "Job" et récupérer les noms de feuilles valides
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
        const dates = sheetData[1].slice(2); // Ligne 2, colonnes à partir de C
    
        // Mettre à jour la liste déroulante des feuilles
        const sheetSelector = document.getElementById("sheetSelector");
        sheetSelector.innerHTML = ""; // Réinitialiser la liste
        sheetNames.forEach(name => {
            const option = document.createElement("option");
            option.value = name;
            option.textContent = name;
            if (name === sheetName) option.selected = true;
            sheetSelector.appendChild(option);
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
    
                // Colorer la cellule si elle correspond à un programme
                if (colIndex >= 2 && rowIndex >= 2) { // Colonnes >= C et lignes >= 3
                    const audienceDate = dates[colIndex - 2]; // Date associée à la colonne
                    const timeSlot = sheetData[rowIndex][1]; // Tranche horaire (colonne B)
    
                    if (audienceDate && timeSlot && isMatchingTimeSlot(audienceDate, timeSlot)) {
                        td.style.backgroundColor = "rgba(255, 0, 0, 0.5)"; // Rouge transparent
                        td.style.color = "white"; // Texte en blanc pour lisibilité
                    }
                }
    
                tr.appendChild(td);
            });
    
            tbody.appendChild(tr);
        });
    
        table.appendChild(tbody);
        previewContainer.innerHTML = ""; // Réinitialiser l'aperçu précédent
        previewContainer.appendChild(table);
            
        // Appel pour surligner la ligne contenant "Pénétration nette (en 1'000)"
        highlightRowByText(sheetData, "Pénétration nette (en 1'000)");
    
        downloadButton.disabled = false;
        console.log("Prévisualisation générée.");
    }
    

   //Sélection des pages
    sheetSelector.addEventListener("change", () => {
        renderPreview(sheetSelector.value);
    });

    historiqueInput.addEventListener("change", (e) => {
        readExcel(e.target.files[0], (workbook) => {
            historiqueData = parseHistorique(workbook);
            console.log("Fichier historique chargé.");
        });
    });

    audienceInput.addEventListener("change", (e) => {
        readExcel(e.target.files[0], (workbook) => {
            audienceData = parseAudience(workbook);
            console.log("Fichier audience chargé.");
        });
    });

    analyzeButton.addEventListener("click", () => {
        if (historiqueData && audienceData) {
            analyzeData();
        } else {
            alert("Veuillez charger les deux fichiers avant d'analyser.");
        }
    });

    downloadButton.addEventListener("click", () => {
        XLSX.writeFile(updatedWorkbook, "updated_audience.xlsx");
    });
});
