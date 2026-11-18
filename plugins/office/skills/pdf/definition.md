---
name: pdf
description: 'PDF — use_skill({"name":"whitecode-office:pdf"}). Créer ou modifier des fichiers PDF, assembler des pages et produire un rapport PDF prêt à partager.'
---

# Pdf

## Documents issus d'une source

Pour un README, un Markdown ou un fichier texte, utilisez le fichier original comme entrée. Une seule exécution accepte tous les formats demandés et lit la source une seule fois. Word conserve la structure, Excel reprend les tableaux et les notes, PowerPoint reprend les sections et le PDF les met en page. Choisissez uniquement les sorties demandées.

Sous Windows, appelez `run_powershell` avec les deux champs suivants (adaptez les noms au projet et à la demande).

```json
{"command":"$env:ELECTRON_RUN_AS_NODE='1'; & $env:OFFICE_NODE $env:OFFICE_RUNTIME --convert './README.md' './note.docx' './lots.xlsx' './revue.pptx' './fiche.pdf'; exit $LASTEXITCODE","attach_files":["note.docx","lots.xlsx","revue.pptx","fiche.pdf"]}
```

Sous Bash, utilisez `run_shell` avec `ELECTRON_RUN_AS_NODE=1 "$OFFICE_NODE" "$OFFICE_RUNTIME" --convert ...` et le même champ `attach_files`. Ce champ contient toutes les pièces finales de la commande, y compris en mode script. Le moteur les livre dans ce tour. Vérifiez le nombre et les noms dans son accusé avant d'annoncer la livraison. Si cet argument est indisponible dans un ancien moteur, créez les documents puis appelez `deliver_file` avec toutes les sorties dans ce même tour.

Si l'utilisateur demande une synthèse, une sélection ou une présentation particulière, reprenez les passages utiles dans leur contexte. Gardez les valeurs, unités, négations, conditions, attributions et états des décisions. Un infinitif ou une case ouverte désigne une action à faire. Vérifiez chaque affirmation de la synthèse contre la source originale avant de générer les fichiers.

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
Si la commande n’a pas déjà livré le fichier, envoyez-le avec `deliver_file` et les champs `files` (chemins absolus), `mode: "attach"`, `caption` et `status: "normal"`.
Ne livrez pas uniquement le script ou un tableau Markdown quand l'utilisateur demande un fichier.

## Exemple de départ

```javascript
const { writeFile } = require('node:fs/promises');
const { resolve } = require('node:path');
module.exports = async ({ PDFLib }) => {
  const {PDFDocument,StandardFonts,rgb} = PDFLib;
  const document = await PDFDocument.create();
  const font = await document.embedFont(StandardFonts.Helvetica);
  const page = document.addPage([595.28,841.89]);
  page.drawText('Rapport de synthèse', {x:48,y:780,size:24,font,color:rgb(0,0.35,0.52)});
  page.drawText('Résultats vérifiés et prochaines étapes.', {x:48,y:730,size:12,font});
  const output = resolve('rapport.pdf');
  const bytes = await document.save();
  if ((await PDFDocument.load(bytes)).getPageCount() !== 1) throw Error('PDF invalide');
  await writeFile(output,bytes,{flag:'wx'}); console.log(output);
};
```

## Spécificités

PDF-lib ne fait ni OCR ni extraction de texte ni mise en page automatique. Calculez retours à la ligne, marges et nouvelles pages. Les polices standard prennent en charge les accents latins usuels, pas tous les alphabets ; ne supprimez pas silencieusement les caractères non pris en charge. Pour une autre écriture utilisez un format Office et sa police adaptée, ou signalez la dépendance de police nécessaire. Documentation : https://pdf-lib.js.org/ .
