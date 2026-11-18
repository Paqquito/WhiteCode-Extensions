---
name: excel
description: 'XLSX — use_skill({"name":"whitecode-office:excel"}). Créer, lire ou modifier des classeurs Excel .xlsx, des budgets, tableaux et rapports avec formules et mise en forme.'
---

# Excel

## Documents issus d'une source

Pour un README, un Markdown ou un fichier texte, utilisez le fichier original comme entrée. Une seule exécution accepte tous les formats demandés et lit la source une seule fois. Word conserve la structure, Excel reprend les tableaux et les notes, PowerPoint reprend les sections et le PDF les met en page. Choisissez uniquement les sorties demandées.

Sous Windows, appelez `run_powershell` avec les deux champs suivants (adaptez les noms au projet et à la demande).

```json
{"command":"$env:ELECTRON_RUN_AS_NODE='1'; & $env:OFFICE_NODE $env:OFFICE_RUNTIME --convert './README.md' './note.docx' './lots.xlsx' './revue.pptx' './fiche.pdf'; exit $LASTEXITCODE","attach_files":["note.docx","lots.xlsx","revue.pptx","fiche.pdf"]}
```

Sous Bash, utilisez `run_shell` avec `ELECTRON_RUN_AS_NODE=1 "$OFFICE_NODE" "$OFFICE_RUNTIME" --convert ...` et le même champ `attach_files`. Ce champ contient toutes les pièces finales de la commande, y compris en mode script. Le moteur les livre dans ce tour. Vérifiez le nombre et les noms dans son accusé avant d'annoncer la livraison. Si cet argument est indisponible dans un ancien moteur, créez les documents puis appelez `deliver_file` avec toutes les sorties dans ce même tour.

Si l'utilisateur demande une synthèse, une sélection ou une présentation particulière, reprenez les passages utiles dans leur contexte. Gardez les valeurs, unités, négations, conditions, attributions et états des décisions. Un infinitif ou une case ouverte désigne une action à faire. Vérifiez chaque affirmation de la synthèse contre la source originale avant de générer les fichiers.

## Création directe

Écrivez avec `write_file` un fichier JSON UTF-8 selon cet exemple : `{"sheets":[{"name":"Budget","rows":[["Poste","Montant"],["Formation",1200]]}]}`. La première ligne porte les en-têtes. Les nombres et booléens conservent leur type. Pour des formules ou graphiques, utilisez le mode script ci-dessous.
Exécutez la commande avec `run_powershell` sous Windows :

```powershell
$env:ELECTRON_RUN_AS_NODE='1'
& $env:OFFICE_NODE $env:OFFICE_RUNTIME --convert './classeur.json' './budget.xlsx'
if ($LASTEXITCODE -ne 0) { throw 'Conversion échouée' }
```

Sous Bash, utilisez `run_shell` avec `ELECTRON_RUN_AS_NODE=1 "$OFFICE_NODE" "$OFFICE_RUNTIME" --convert './classeur.json' './budget.xlsx'`.
Le programme imprime le chemin absolu du nouveau fichier. Contrôlez l’accusé de livraison des fichiers demandés.

## Mode script

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
Si la commande n’a pas déjà livré le fichier, envoyez-le avec `deliver_file` et les champs `files` (chemins absolus), `mode: "attach"`, `caption` et `status: "normal"`.
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
