# WhiteCode-Extensions

Catalogue de plugins de WhiteCode.

## Catalogue de plugins

`.claude-plugin/marketplace.json` est le catalogue que WhiteCode ajoute comme source par défaut au premier lancement, épinglé sur un tag (`vX.Y.Z`). Un plugin s'installe par `git clone`, sans rien construire sur le poste : tout ce qu'il exécute doit être versionné, déjà construit.

| Plugin | Dossier | Contenu |
|---|---|---|
| Bureautique (`whitecode-office`) | `plugins/office/` | 4 skills (Word, Excel, PowerPoint, PDF) et `office.cjs`, le runtime qui embarque exceljs, docx, pptxgenjs et pdf-lib |
| Outlook (`whitecode-outlook`) | `plugins/outlook/` | Recherche par période et mots-clés, lecture, synthèses, brouillons et invitations dans Outlook classique pour Windows |

### Utiliser Outlook

Installez et activez **Outlook** dans le catalogue des extensions sur une révision qui contient le plugin. Le plugin utilise le profil Outlook classique de la session Windows. Le nouvel Outlook et Outlook Web ne fournissent pas cette interface. Aucun jeton ni service supplémentaire à configurer.

Demandez par exemple « Résume ma boîte de réception depuis lundi et liste les décisions », « Cherche budget dans les messages envoyés la semaine dernière » ou « Prépare deux brouillons séparés pour les destinataires suivants ». La recherche porte sur les champs du message et le corps en texte brut. Elle couvre le dossier demandé du compte par défaut, avec pagination. Les extraits lus entrent dans la conversation avec le fournisseur de modèle sélectionné. Pour une boîte professionnelle, choisissez un fournisseur autorisé à traiter ses messages.

Pour un mail en deux temps, demandez « Prépare un brouillon à mon adresse », puis « Envoie ce brouillon uniquement à mon adresse ». L’extension rouvre le message enregistré et vérifie son objet ainsi que les destinataires en À, Cc et Cci avant soumission à Outlook. Si les destinataires ou leurs rôles ont changé, le brouillon est conservé.

Pour une réunion, précisez participants, date, heure et durée. L’invitation est enregistrée sans envoi dans le calendrier personnel, avec le statut Disponible. Les courriels restent dans Brouillons. Un lien de visio fourni peut y être inséré. Pour Zoom ou Tixeo, le plugin ouvre la fenêtre de cette invitation et actionne le bouton du fournisseur dans le ruban. Pour Zoom, la commande valide le bouton Continuez dans la fenêtre de cette invitation puis vérifie le nouveau lien de participation en relisant le rendez-vous enregistré. Si le bouton est indisponible ou le lien absent, le brouillon reste disponible pour vérification. Pour Tixeo, vérifiez que « Réunion en ligne » est activé. Si la commande est grisée, le plugin le signale et conserve l’invitation.

### Générer un document

Une fois le plugin Bureautique installé et activé, demandez par exemple « Transforme le README.md en document Word et joins-le ». Le modèle charge le skill avec `use_skill` puis utilise le runtime livré dans l’extension. La commande `office.cjs --convert README.md note.docx donnees.xlsx revue.pptx fiche.pdf` génère les sorties demandées à partir des mêmes passages. Les tableaux restent numériques dans Excel et les actions ouvertes conservent leur libellé. Le champ `attach_files` de la commande déclenche leur livraison dans le tour en cours. Le runtime préserve les fichiers existants. Les entrées JSON et le mode script restent disponibles pour des mises en page personnalisées.

### Modifier la Bureautique

```bash
bun install --frozen-lockfile
bun run build:office   # reconstruit plugins/office/office.cjs et THIRD_PARTY_NOTICES.txt
bun test plugins
```

Committez `office.cjs` avec la modification, augmentez la version dans `plugins/office/.claude-plugin/plugin.json`, `extension.json` et `marketplace.json`, puis posez un nouveau tag. WhiteCode ne change de version que lorsque l'utilisateur actualise la source sur ce tag.

## Module natif GitLab

Le dossier `modules/gitlab` est compilé dans les distributions qui choisissent `gitlab` dans leur identité desktop. Dans Paramètres → GitLab, saisissez l’URL HTTPS de l’instance et un jeton personnel avec le droit `read_api` pour consulter les demandes de fusion, ou `api` pour les modifier selon les permissions du compte. Dans **Panneaux → GitLab**, choisissez le projet et ouvrez les actions de la demande. Vous pouvez créer ou modifier un brouillon, le passer prêt, commenter une ligne, résoudre une discussion, agir sur les jobs et le pipeline, ou fusionner la révision affichée. Une boîte de dialogue native confirme le compte, le projet et les paramètres avant chaque écriture. La fusion transmet le SHA à GitLab. Les paramètres de compte sont protégés par le coffre du système. Lancez `bun run test` pour vérifier le module et Bureautique.

Une demande explicite comme « Crée une réunion Zoom demain à 13 h et envoie-moi l’invitation uniquement à mon adresse » autorise l’envoi après relecture du lien et des destinataires. Vous pouvez aussi demander l’envoi dans un second message pour un rendez-vous Zoom préparé avec Outlook 1.3.0 ou ultérieur. Sans demande d’envoi, le rendez-vous reste à relire.
