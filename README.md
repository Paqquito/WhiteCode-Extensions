# WhiteCode-Extensions

Plugins de WhiteCode et personnalisations de la distribution whitecode.

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

## Personnalisations de la distribution

`whitecode/` contient le code propre à la distribution whitecode (compte, lancement du moteur, outils métier), importé par le build du desktop. Ce n'est pas un plugin. Voir `whitecode/README.md`.
