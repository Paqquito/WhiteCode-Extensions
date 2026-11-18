---
name: powerpoint
description: 'PPTX — use_skill({"name":"whitecode-office:powerpoint"}). Créer des présentations PowerPoint .pptx : synthèses, revues de projet et diapositives avec tableaux, images et graphiques.'
---

# Powerpoint

## Création directe

Écrivez avec `write_file` un fichier JSON UTF-8 selon cet exemple : `{"slides":[{"title":"Bilan","bullets":["Prototype livré","Recette en cours"],"notes":"Présenter les résultats."}]}`. Chaque entrée crée une diapositive. Pour des graphiques ou des images, utilisez le mode script ci-dessous.
Exécutez la commande avec `run_powershell` sous Windows :

```powershell
$env:ELECTRON_RUN_AS_NODE='1'
& $env:OFFICE_NODE $env:OFFICE_RUNTIME --convert './diaporama.json' './presentation.pptx'
if ($LASTEXITCODE -ne 0) { throw 'Conversion échouée' }
```

Sous Bash, utilisez `run_shell` avec `ELECTRON_RUN_AS_NODE=1 "$OFFICE_NODE" "$OFFICE_RUNTIME" --convert './diaporama.json' './presentation.pptx'`.
Le programme imprime le chemin absolu du nouveau fichier. Appelez ensuite `deliver_file` avec `mode: "attach"` et ce chemin dans `files`.

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
Envoyez le fichier final avec `deliver_file` et les champs `files` (chemins absolus), `mode: "attach"`, `caption` et `status: "normal"`.
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
