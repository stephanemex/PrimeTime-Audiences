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
 * Fonction pour convertir des chaînes d'horaires en objets Date.
 * @param {String} date - Date au format "dd.mm.yyyy".
 * @param {String} timeStr - Heure au format "HH:MM:SS".
 * @returns {Date} - Objet Date corrigé pour les heures > 24.
 */
function convertAudienceTimes(date, timeStr) {
    if (!date || !timeStr || typeof timeStr !== "string") {
        console.error("Date ou heure invalide : ", { date, timeStr });
        return null;
    }

    const [hours, minutes, seconds] = timeStr.split(":").map(Number);
    const adjustedHours = hours >= 24 ? hours - 24 : hours; // Gérer les heures > 24
    const dayOffset = hours >= 24 ? 1 : 0; // Décalage d'un jour si nécessaire

    const result = new Date(
        new Date(date).getFullYear(),
        new Date(date).getMonth(),
        new Date(date).getDate() + dayOffset,
        adjustedHours,
        minutes,
        seconds || 0
    );

    console.log(`Conversion réussie : ${timeStr} -> ${result}`);
    return result;
}

function convertTimeStringToDate(timeString) {
    const [hours, minutes, seconds] = timeString.split(":").map(Number);
    return new Date(1970, 0, 1, hours, minutes, seconds || 0);
}

/**
 * Vérifie si une chaîne correspond au format d'une plage horaire HH:MM:SS - HH:MM:SS.
 * @param {String} timeSlot - La chaîne à vérifier.
 * @returns {Boolean} - True si la chaîne est valide.
 */
function isValidTimeSlot(timeSlot) {
    const regex = /^(\d{2}:\d{2}:\d{2}) - (\d{2}:\d{2}:\d{2})$/;
    return regex.test(timeSlot);
}

/**
 * Vérifie si une chaîne correspond au format d'une heure HH:MM:SS.
 * @param {String} timeString - La chaîne à vérifier.
 * @returns {Boolean} - True si la chaîne est valide.
 */
function isValidTimeString(timeString) {
    const regex = /^\d{2}:\d{2}:\d{2}$/;
    return regex.test(timeString);
}
