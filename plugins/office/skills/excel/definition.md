---
name: excel
description: "Créer, lire ou modifier des classeurs Excel .xlsx, des budgets, tableaux et rapports avec formules et mise en forme."
---

# Excel

## Exécution

Les bibliothèques sont livrées avec ce plugin (`office.cjs`). Ne lancez ni pip ni npm pour cette tâche.
Créez dans le dossier de travail un script `.cjs` exportant une fonction asynchrone qui reçoit les bibliothèques.
Les chemins du runtime sont fournis par les variables OFFICE_NODE et OFFICE_RUNTIME.

PowerShell (conserver les guillemets, les chemins peuvent contenir des espaces) :
```powershell
$env:ELECTRON_RUN_AS_NODE='1'
& $env:OFFICE_NODE $env:OFFICE_RUNTIME ./creer.cjs
if ($LASTEXITCODE -ne 0) { throw 'Création du document échouée' }
```
Bash :
```bash
ELECTRON_RUN_AS_NODE=1 "$OFFICE_NODE" "$OFFICE_RUNTIME" ./creer.cjs
```
Ce script est exécuté avec l'outil de commande habituel et ses permissions. Une bibliothèque incluse ne donne aucune autorisation supplémentaire.
Si les variables sont absentes, expliquez que le plugin Bureautique n'est pas installé ou n'est pas activé dans WhiteCode (Paramètres > Plugins) ; ne devinez aucun chemin local.

## Livraison et contrôle

Adaptez réellement l'exemple aux données et au résultat demandés. Utilisez des chemins absolus pour les fichiers livrés.
Préservez les originaux : écrivez un nouveau fichier, sauf demande explicite de remplacer l'existant.
Vérifiez les données, les unités, les accents et la structure du résultat. Ne prétendez jamais avoir recalculé ou inspecté visuellement un fichier sans l'avoir fait.
Ouvrez le document généré dans l'aperçu si disponible. LibreOffice peut être nécessaire pour l'aperçu Office, mais pas pour créer le fichier.
Envoyez le fichier final avec SendUserFile (files: [chemin absolu], caption: description courte, status: normal).
Ne livrez pas uniquement le script ou un tableau Markdown quand l'utilisateur demande un fichier.

## Exemple de départ

```javascript
const { resolve } = require('node:path');
const { writeFile } = require('node:fs/promises');
module.exports = async ({ ExcelJS }) => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'WhiteCode';
  workbook.calcProperties.fullCalcOnLoad = true;
  const sheet = workbook.addWorksheet('Budget');
  sheet.columns = [{header:'Poste',key:'poste',width:30},{header:'Montant (€)',key:'montant',width:20}];
  sheet.addRows([{poste:'Matériel',montant:1200},{poste:'Services',montant:800}]);
  sheet.addRow(['Total',{formula:'SUM(B2:B3)',result:2000}]);
  sheet.getColumn(2).numFmt = '#,##0.00';
  sheet.getRow(1).font = {bold:true,color:{argb:'FFFFFFFF'}};
  sheet.getRow(1).fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FF005A85'}};
  sheet.views = [{state:'frozen',ySplit:1}];
  const output = resolve('budget.xlsx');
  await writeFile(output, await workbook.xlsx.writeBuffer(), {flag:'wx'});
  const check = new ExcelJS.Workbook(); await check.xlsx.readFile(output);
  if (check.getWorksheet('Budget').getCell('B4').value.result !== 2000) throw Error('Total invalide');
  console.log(output);
};
```

## Spécificités

ExcelJS ne calcule pas les formules. Fournissez un résultat mis en cache vérifié quand il est connu et activez fullCalcOnLoad. Ne fabriquez pas un résultat de formule. Conservez des valeurs numériques pour les nombres, des formats de cellule pour unités et devises. Lecture : workbook.xlsx.readFile(path). Documentation : https://github.com/exceljs/exceljs .
