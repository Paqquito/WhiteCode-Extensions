# WhiteCode-Extensions

Catalogue de plugins de WhiteCode.

## Catalogue de plugins

`.claude-plugin/marketplace.json` est le catalogue que WhiteCode ajoute comme source par défaut au premier lancement, épinglé sur un tag (`vX.Y.Z`). Un plugin s'installe par `git clone`, sans rien construire sur le poste : tout ce qu'il exécute doit être versionné, déjà construit.

| Plugin | Dossier | Contenu |
|---|---|---|
| Bureautique (`whitecode-office`) | `plugins/office/` | 4 skills (Word, Excel, PowerPoint, PDF) et `office.cjs`, le runtime qui embarque exceljs, docx, pptxgenjs et pdf-lib |

### Générer un document

Une fois le plugin Bureautique installé et activé, demandez par exemple « Transforme le README.md en document Word et joins-le ». Le modèle charge le skill avec `use_skill` puis utilise le runtime livré dans l’extension. La commande `office.cjs --convert source sortie` transforme un Markdown en DOCX, ou un JSON en XLSX ou PPTX. Les formats du JSON figurent dans les skills Excel et PowerPoint. Le runtime refuse d’écraser les fichiers existants. Le mode script reste disponible pour le PDF et les mises en page personnalisées.

### Modifier la Bureautique

```bash
bun install --frozen-lockfile
bun run build:office   # reconstruit plugins/office/office.cjs et THIRD_PARTY_NOTICES.txt
bun test plugins
```

Committez `office.cjs` avec la modification, augmentez la version dans `plugins/office/.claude-plugin/plugin.json`, `extension.json` et `marketplace.json`, puis posez un nouveau tag. WhiteCode ne change de version que lorsque l'utilisateur actualise la source sur ce tag.

## Module natif GitLab

Le dossier `modules/gitlab` est compilé dans les distributions qui choisissent `gitlab` dans leur identité desktop. Dans Paramètres → GitLab, saisissez l’URL HTTPS de l’instance et un jeton personnel avec le droit `read_api` pour consulter les demandes de fusion, ou `api` pour les modifier selon les permissions du compte. Dans **Panneaux → GitLab**, choisissez le projet et ouvrez les actions de la demande. Vous pouvez créer ou modifier un brouillon, le passer prêt, commenter une ligne, résoudre une discussion, agir sur les jobs et le pipeline, ou fusionner la révision affichée. Une boîte de dialogue native confirme le compte, le projet et les paramètres avant chaque écriture. La fusion transmet le SHA à GitLab. Les paramètres de compte sont protégés par le coffre du système. Lancez `bun run test` pour vérifier le module et Bureautique.
