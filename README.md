# WhiteCode Extensions

Ce dépôt contient les extensions Bureautique et Outlook pour WhiteCode, ainsi que le module GitLab de [WhiteCode Desktop](https://github.com/Paqquito/WhiteCode-Desktop#readme).

## Installer une extension

Dans le desktop, ouvrez **Plugins et skills**, choisissez l’extension, cliquez sur **Installer**, puis activez-la. Git doit être disponible sur le poste.

Le catalogue par défaut pointe sur le tag `v1.1.0`, qui contient Bureautique. Pour installer **Outlook** depuis la branche `main`, ouvrez **Sources du catalogue**, saisissez `main` dans le champ de référence du catalogue officiel et cliquez sur **Actualiser la source**. Vous pourrez alors installer Outlook.

| Extension | Prérequis | Exemple de demande |
| --- | --- | --- |
| [Bureautique](plugins/office/) | Le moteur de génération est livré avec l’extension. | « Transforme le README en document Word et joins-le. » Formats disponibles : Word, Excel, PowerPoint et PDF. |
| [Outlook](plugins/outlook/) | Outlook classique pour Windows avec un profil connecté. Pour une visioconférence, le complément correspondant doit être configuré dans Outlook. | « Résume mes mails depuis lundi » ou « Prépare une réunion Zoom de 30 minutes demain matin avec les participants suivants… » |

Les extraits de mails lus entrent dans la conversation avec le fournisseur de modèle sélectionné. Choisissez un fournisseur autorisé à traiter ces messages. Outlook prépare un brouillon et vérifie les destinataires avant un envoi demandé explicitement. Pour une réunion, le desktop propose d’autres créneaux en cas de conflit et le complément vérifie le lien de visioconférence avant l’envoi.

## Configurer GitLab

Le [module GitLab](modules/gitlab/) est intégré lorsque l’identité du desktop inclut `gitlab`. Dans **Paramètres → GitLab**, renseignez l’URL HTTPS de l’instance et un jeton personnel avec le droit `read_api` pour la consultation ou `api` pour les écritures. Ouvrez **Panneaux → GitLab** pour choisir le projet et la demande de fusion. Le desktop demande confirmation avant chaque écriture.

## Modifier et tester les extensions

Depuis la racine du dépôt, avec Bun 1.4.2 :

```sh
bun install --frozen-lockfile
bun run build:office
bun run test
```

La construction de Bureautique régénère `office.cjs` et ses notices. Le [manifeste du catalogue](.claude-plugin/marketplace.json) indique les extensions distribuées.

## Licence

Le catalogue, les deux extensions et le module GitLab sont sous [licence MIT](LICENSE). Chaque composant installable contient sa licence. Les [notices des bibliothèques de Bureautique](plugins/office/THIRD_PARTY_NOTICES.txt) accompagnent son moteur de génération.
