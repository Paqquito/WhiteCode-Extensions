# WhiteCode-Extensions

Catalogue de plugins de WhiteCode.

## Catalogue de plugins

`.claude-plugin/marketplace.json` est le catalogue que WhiteCode ajoute comme source par défaut au premier lancement, épinglé sur un tag (`vX.Y.Z`). Un plugin s'installe par `git clone`, sans rien construire sur le poste : tout ce qu'il exécute doit être versionné, déjà construit.

| Plugin | Dossier | Contenu |
|---|---|---|
| Bureautique (`whitecode-office`) | `plugins/office/` | 4 skills (Word, Excel, PowerPoint, PDF) et `office.cjs`, le runtime qui embarque exceljs, docx, pptxgenjs et pdf-lib |

### Modifier la Bureautique

```bash
bun install --frozen-lockfile
bun run build:office   # reconstruit plugins/office/office.cjs et THIRD_PARTY_NOTICES.txt
bun test plugins
```

Committez `office.cjs` avec la modification, augmentez la version dans `plugins/office/.claude-plugin/plugin.json`, `extension.json` et `marketplace.json`, puis posez un nouveau tag. WhiteCode ne change de version que lorsque l'utilisateur actualise la source sur ce tag.

## Module natif GitLab

Le dossier `modules/gitlab` est compilé dans les distributions qui choisissent `gitlab` dans leur identité desktop. Dans Paramètres → GitLab, saisissez l’URL HTTPS de l’instance et un jeton personnel avec le droit `read_api`. Ouvrez ensuite Panneaux → GitLab et indiquez le chemin du projet pour consulter ses demandes de fusion ouvertes. Les paramètres de compte sont protégés par le coffre du système. Lancez `bun run test` pour vérifier le module et Bureautique.
