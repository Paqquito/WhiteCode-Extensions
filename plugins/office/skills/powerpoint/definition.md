---
name: powerpoint
description: 'PPTX — use_skill({"name":"whitecode-office:powerpoint"}). Créer des présentations PowerPoint .pptx : synthèses, revues de projet et diapositives avec tableaux, images et graphiques.'
---

# Powerpoint

## Documents issus d'une source

Pour un README, un Markdown ou un fichier texte, utilisez le fichier original comme entrée. Une seule exécution accepte tous les formats demandés et lit la source une seule fois. Word conserve la structure, Excel reprend les tableaux et les notes, PowerPoint reprend les sections et le PDF les met en page. Choisissez uniquement les sorties demandées.

Sous Windows, appelez `run_powershell` avec les deux champs suivants (adaptez les noms au projet et à la demande).

```json
{"command":"$env:ELECTRON_RUN_AS_NODE='1'; & $env:OFFICE_NODE $env:OFFICE_RUNTIME --convert './README.md' './note.docx' './lots.xlsx' './revue.pptx' './fiche.pdf'; exit $LASTEXITCODE","attach_files":["note.docx","lots.xlsx","revue.pptx","fiche.pdf"]}
```

Sous Bash, utilisez `run_shell` avec `ELECTRON_RUN_AS_NODE=1 "$OFFICE_NODE" "$OFFICE_RUNTIME" --convert ...` et le même champ `attach_files`. Ce champ contient toutes les pièces finales de la commande, y compris en mode script. Le moteur les livre dans ce tour. Vérifiez le nombre et les noms dans son accusé avant d'annoncer la livraison. Si cet argument est indisponible dans un ancien moteur, créez les documents puis appelez `deliver_file` avec toutes les sorties dans ce même tour.

Si l'utilisateur demande une synthèse, une sélection ou une présentation particulière, reprenez les passages utiles dans leur contexte. Gardez les valeurs, unités, négations, conditions, attributions et états des décisions. Un infinitif ou une case ouverte désigne une action à faire. Vérifiez chaque affirmation de la synthèse contre la source originale avant de générer les fichiers.

## Création directe

Écrivez avec `write_file` un fichier JSON UTF-8 selon cet exemple : `{"slides":[{"title":"Bilan","bullets":["Prototype livré","Recette en cours"],"notes":"Présenter les résultats."}]}`. Chaque entrée crée une diapositive. Pour des graphiques ou des images, utilisez le mode script ci-dessous.
Exécutez la commande avec `run_powershell` sous Windows :

```powershell
$env:ELECTRON_RUN_AS_NODE='1'
& $env:OFFICE_NODE $env:OFFICE_RUNTIME --convert './diaporama.json' './presentation.pptx'
if ($LASTEXITCODE -ne 0) { throw 'Conversion échouée' }
```

Sous Bash, utilisez `run_shell` avec `ELECTRON_RUN_AS_NODE=1 "$OFFICE_NODE" "$OFFICE_RUNTIME" --convert './diaporama.json' './presentation.pptx'`.
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
module.exports = async ({ PptxGenJS }) => {
  const presentation = new PptxGenJS();
  presentation.layout = 'LAYOUT_WIDE';
  presentation.author = 'WhiteCode';
  presentation.subject = 'Revue de projet';
  const slide = presentation.addSlide();
  slide.addText('Revue de projet', {x:0.6,y:0.5,w:12,h:0.7,fontSize:30,bold:true,color:'005A85'});
  slide.addText('Résultats et prochaines étapes', {x:0.6,y:1.5,w:12,h:1,fontSize:20});
  const output = resolve('presentation.pptx');
  await writeFile(output, await presentation.write({outputType:'nodebuffer'}), {flag:'wx'});
  console.log(output);
};
```

## Spécificités

Une idée principale par diapositive ; limitez le texte, utilisez des marges cohérentes et vérifiez les débordements. Les éléments restent éditables. PptxGenJS crée des présentations ; il ne lit pas un PPTX existant. Documentation : https://gitbrent.github.io/PptxGenJS/docs/quick-start/ .
