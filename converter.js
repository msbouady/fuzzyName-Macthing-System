import xlsx from 'xlsx';
import FuzzySet from 'fuzzyset.js';
import path from 'path';

// Fonction pour nettoyer les noms
function nettoyerNom(nom) {
    nom = nom.replace(/\./g, '').trim();
    let parts = nom.split(/\s+/);
    let variations = [parts.join(' ')];
    if (parts.length > 1) {
        variations.push(parts.reverse().join(' '));
    }
    return variations;
}

// Fonction qui compare les noms avec FuzzySet
function comparerNoms(nom, fuzzySet) {
    let variationsNom = nettoyerNom(nom);
    let bestScore = 0;
    let bestMatch = null;
    for (let variation of variationsNom) {
        let matches = fuzzySet.get(variation);
        if (matches && matches[0][0] > bestScore) {
            bestScore = matches[0][0];
            bestMatch = matches[0];
        }
    }
    return bestScore >= 0.7 ? bestMatch : null;
}

// Chemins des fichiers Excel
const filePath1 = path.resolve('C:\\Users\\M\\Downloads\\SS\\SS 1.xlsx');
const filePath2 = path.resolve('C:\\Users\\M\\Downloads\\SS\\SS 2.xlsx');

// Charger les fichiers Excel
const workbook1 = xlsx.readFile(filePath1);
const workbook2 = xlsx.readFile(filePath2);

// Sélectionner les feuilles de calcul
const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];

// Convertir les feuilles en JSON
const ss1 = xlsx.utils.sheet_to_json(sheet1, { defval: '' });
const ss2 = xlsx.utils.sheet_to_json(sheet2, { defval: '' });

// Extraction de la colonne 'ncomplet'
let ncomplet_ss1 = ss1.map(row => row.ncomplet);
let ncomplet_ss2 = ss2.map(row => row.ncomplet);

// Création des ensembles flous (FuzzySet)
let fuzzySet1 = FuzzySet(ncomplet_ss1);
let fuzzySet2 = FuzzySet(ncomplet_ss2);

// Ajouter l'état au JSON
ss1.forEach(row => row.etat = 'sortant');
ss2.forEach(row => row.etat = 'entrant');

// Associer les colonnes et ajouter l'état
ss1.forEach((row1, index1) => {
    let match = comparerNoms(row1.ncomplet, fuzzySet2);
    if (match) {
        let matchedRowIndex = ss2.findIndex(row2 => nettoyerNom(row2.ncomplet).includes(match[1]));
        if (matchedRowIndex !== -1) {
            ss1[index1] = { ...row1, ...ss2[matchedRowIndex], etat: 'admis' };
            ss2[matchedRowIndex].etat = 'admis';
        }
    }
});

ss2.forEach((row2, index2) => {
    let match = comparerNoms(row2.ncomplet, fuzzySet1);
    if (match) {
        let matchedRowIndex = ss1.findIndex(row1 => nettoyerNom(row1.ncomplet).includes(match[1]));
        if (matchedRowIndex !== -1) {
            ss2[index2] = { ...row2, ...ss1[matchedRowIndex], etat: 'admis' };
            ss1[matchedRowIndex].etat = 'admis';
        }
    }
});

// Convertir le JSON en feuilles Excel
const newSheet1 = xlsx.utils.json_to_sheet(ss1);
const newSheet2 = xlsx.utils.json_to_sheet(ss2);

// Création de nouveaux classeurs et ajout des feuilles
const newWorkbook1 = xlsx.utils.book_new();
const newWorkbook2 = xlsx.utils.book_new();

xlsx.utils.book_append_sheet(newWorkbook1, newSheet1, 'Sheet1');
xlsx.utils.book_append_sheet(newWorkbook2, newSheet2, 'Sheet1');

// Sauvegarder les nouveaux fichiers Excel
xlsx.writeFile(newWorkbook1, 'SS1_etat.xlsx');
xlsx.writeFile(newWorkbook2, 'SS2_etat.xlsx');
