---
name: visio
description: 'Organise une réunion Zoom ou Tixeo, programme une visio, ajoute un lien Zoom ou Tixeo à une invitation Outlook : préparer réellement le calendrier et ouvrir le complément demandé.'
---

Pour organiser une réunion, utilisez maintenant `run_powershell` et le script installé `${SKILL_ROOT}\..\outlook\outlook.ps1`. Ce poste peut préparer une vraie invitation. Ne créez aucun fichier .ics sauf demande explicite d’export ; un fichier et une URL d’exemple ne sont pas une réservation. Résolvez les noms fournis dans Outlook via -To. Traduisez les dates relatives avec l’horloge locale du poste. Si seule l’heure de début est donnée, prévoyez 30 minutes et indiquez la durée retenue. Si l’heure demandée est déjà passée, demandez une précision. Utilisez le nom ou l’adresse réellement fournis : Outlook résout les participants ; demandez une précision en cas d’ambiguïté. Échappez les apostrophes des données dans les chaînes PowerShell en les doublant.

Exemple pour demain à 13 h, à adapter à la demande :

```powershell
$debut = [datetimeoffset]([datetime]::Today.AddDays(1).AddHours(13))
& '${SKILL_ROOT}\..\outlook\outlook.ps1' -Action meeting -Subject 'Réunion de suivi' -To @('personne@example.org') -Start $debut.ToString('yyyy-MM-ddTHH:mmzzz') -End $debut.AddMinutes(30).ToString('yyyy-MM-ddTHH:mmzzz') -VideoProvider Zoom
```

Pour Tixeo, remplacez Zoom par Tixeo. Un lien HTTPS déjà fourni peut être passé avec `-VideoUrl` à la place de `-VideoProvider`. Ne demandez pas à l’utilisateur de fournir un lien quand il demande à en créer un. La commande enregistre l’invitation et ouvre son ruban. Une fois `saved=true`, ne recréez pas l’événement, même si le complément attend une action.

- `zoom-link-verified` : le complément a ajouté un nouveau lien de participation Zoom et le script a relu ce même lien dans le rendez-vous enregistré. Invitez l’utilisateur à relire et envoyer dans Outlook.
- `zoom-review-open` : la configuration Zoom est ouverte. Demandez de valider **Continuez** dans cette fenêtre, puis de vérifier le lien dans l’invitation avant l’envoi.
- `tixeo-enabled` : le bouton **Réunion en ligne** est activé sur cette invitation. L’organisateur vérifie les réglages puis envoie dans Outlook.
- `addin-command-disabled` : Outlook a grisé la commande. Pour Tixeo, demander de vérifier la connexion au client avec le même compte et le droit organisateur. L’invitation préparée est conservée.
- Tout autre état : indiquer que le lien reste à vérifier ou à ajouter dans l’invitation existante. Ne dites jamais qu’un lien a été créé à partir du seul clic sur le ruban. L’envoi reste à la main de l’utilisateur.
