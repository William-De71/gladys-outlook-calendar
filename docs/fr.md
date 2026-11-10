# Outlook Calendar

## Présentation

Cette intégration amène vos **calendriers Outlook** dans Gladys — Outlook.com, Hotmail, Live ou Microsoft 365 — via l’API officielle Microsoft Graph, en lecture seule. **Chaque utilisateur de Gladys relie son propre compte Microsoft** : ses calendriers restent les siens, privés tant qu’il ne les partage pas avec la maison. Les événements apparaissent alors dans le calendrier de Gladys, déclenchent des scènes avec « un événement arrive », et l’assistant peut les lire.

Les événements récurrents, les journées entières, les calendriers partagés avec vous et les liens de réunion en ligne sont gérés. Les événements privés peuvent être affichés « Occupé » ou ignorés, et les invitations refusées sont ignorées par défaut.

Non affiliée à Microsoft. Pas encore testée avec un vrai compte Microsoft : vos retours sont les bienvenus.

## Prérequis

- Gladys Assistant 5.3.0 ou plus récent.
- Un compte Microsoft avec un calendrier : personnel (Outlook.com, Hotmail, Live) ou professionnel (Microsoft 365).
- Une seule fois, pour toute la maison : une **inscription d’application** gratuite dans Microsoft Entra, faite par l’administrateur de Gladys (n’importe quel compte Microsoft peut en créer une).
- **Vos données** : les événements de chaque utilisateur sont lus par Gladys chez Microsoft ; rien n’est jamais écrit dans Outlook.

### Créer l’inscription d’application

1. Allez sur [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **Inscriptions d’applications** → **Nouvelle inscription** (un compte Microsoft personnel suffit).
2. Nom : `Gladys`. Types de comptes pris en charge : **Comptes dans un annuaire d’organisation et comptes Microsoft personnels**. URI de redirection : laissez vide. **Inscrire**.
3. **Authentification** → **Paramètres avancés** → **Autoriser les flux clients publics** : **Oui** → **Enregistrer**.
4. **Autorisations d’API** → **Ajouter une autorisation** → **Microsoft Graph** → **Autorisations déléguées** : `Calendars.Read`, `User.Read`, `offline_access` → **Ajouter des autorisations**.
5. Copiez l’**ID d’application (client)** depuis la page **Vue d’ensemble**.

Pas de secret client, pas d’adresse de redirection : les utilisateurs se connectent avec un code, depuis n’importe quel appareil.

## Configuration

| Champ ou bouton (administrateur)     | Par défaut | Rôle                                                                                                         |
| ------------------------------------ | ---------- | ------------------------------------------------------------------------------------------------------------ |
| ID d’application (client)            | —          | Obligatoire : l’ID de votre inscription d’application                                                        |
| Locataire (tenant)                   | common     | `common` pour les comptes personnels et professionnels, `consumers` pour les personnels, ou votre entreprise |
| Synchroniser toutes les              | 15 minutes | 5, 15, 30 minutes ou 1 heure                                                                                 |
| Événements à venir                   | 12 mois    | 3, 6 ou 12 mois ; les 30 derniers jours sont gardés aussi                                                    |
| Vérifier l’inscription d’application | —          | Demande à Microsoft si l’application accepte les connexions, et dit ce qui ne va pas sinon                   |
| Synchroniser maintenant              | —          | Relit les calendriers de chaque utilisateur relié                                                            |

| Champ (chaque utilisateur, « Mes calendriers ») | Par défaut          | Rôle                                                                              |
| ----------------------------------------------- | ------------------- | --------------------------------------------------------------------------------- |
| Événements privés                               | Affichés « Occupé » | Affichés « Occupé » (sans titre ni lieu), affichés avec leurs détails, ou ignorés |
| Invitations refusées                            | Ignorées            | Ignorées ou affichées                                                             |

## Utilisation

1. L’administrateur : Intégrations → **Outlook Calendar** → Configuration : collez l’ID d’application, enregistrez, appuyez sur **Vérifier l’inscription d’application**.
2. Chaque utilisateur : la même page, bloc **Mes calendriers** → **Activer**.
3. Ses calendriers arrivent, privés ; chacun peut être désactivé ou partagé avec la maison depuis le même bloc.

### Relier votre compte

Gladys n’a pas encore de bouton « Connecter » par utilisateur : le code de connexion arrive donc dans **un calendrier temporaire**, dans votre liste « Mes calendriers » (et comme événement dans votre vue calendrier), nommé par exemple :

> Microsoft : tapez ABCD-EFGH sur microsoft.com/devicelogin

Ouvrez [microsoft.com/devicelogin](https://microsoft.com/devicelogin) sur n’importe quel ordinateur ou téléphone, tapez le code, connectez-vous avec votre compte Microsoft et acceptez les autorisations (lire vos calendriers, garder l’accès). En quelques secondes, le calendrier temporaire disparaît et vos calendriers Outlook apparaissent. Le code vaut environ 15 minutes et se renouvelle tout seul pendant une heure ; au-delà, enregistrez « Mes calendriers » à nouveau pour en obtenir un nouveau.

Pour relier un autre compte Microsoft, désactivez « Mes calendriers » puis réactivez-le. Si Microsoft n’accepte plus la connexion (mot de passe changé, accès retiré, 90 jours sans usage), le calendrier temporaire revient avec un nouveau code.

### Widget de tableau de bord

Le widget **Agendas Outlook** affiche combien d’utilisateurs sont reliés, les calendriers et événements synchronisés, et une ligne par utilisateur qui a activé « Mes calendriers » : synchronisé à quelle heure, en attente de connexion, ou en erreur. Seulement les prénoms : ni e-mail, ni événement. Un bouton synchronise tout de suite.

![Le widget Agendas Outlook : comptes reliés, calendriers et événements, l’état de chaque utilisateur et un bouton de synchronisation (illustration générée, pas une capture d’écran)](https://raw.githubusercontent.com/William-De71/gladys-outlook-calendar/main/docs/assets/widgets.fr.png)

### Scènes

Les événements eux-mêmes déclenchent des scènes avec le déclencheur Gladys **« Un événement de calendrier arrive »** (sur les calendriers partagés). L’intégration ajoute :

- **Déclencheur « Outlook : un utilisateur doit se connecter »** : un utilisateur a activé « Mes calendriers » (`new`) ou Microsoft n’accepte plus sa connexion (`revoked`), avec son prénom.
- **Déclencheur « Outlook : un utilisateur s’est connecté »** : avec son prénom et son compte Microsoft.
- **Déclencheur « Outlook : la synchro échoue »** : une fois par panne et par utilisateur (accès refusé, pas de boîte mail, trop de requêtes, Microsoft injoignable, erreur Microsoft).
- **Déclencheur « Outlook : la synchro refonctionne »** : avec les minutes de panne.
- **Action « Outlook : synchroniser maintenant »** : avant une scène qui regarde les événements du jour.

Exemples : quand un utilisateur doit se reconnecter (`revoked`), lui envoyer un message Telegram « Ton lien Outlook a expiré, regarde Mes calendriers » ; chaque matin à 7 h, synchroniser maintenant, puis lire les événements du jour à voix haute.

## Dépannage

- **« Microsoft refuse l’inscription d’application : AADSTS700016 »** : l’ID d’application est faux, ou l’application a été supprimée.
- **« AADSTS7000218 »** ou un secret client demandé : **Autoriser les flux clients publics** n’est pas sur Oui (étape 3).
- **« AADSTS50020 »** ou **« AADSTS9002331 »** : les types de comptes de l’application ne correspondent pas au locataire : choisissez « annuaire d’organisation et comptes Microsoft personnels », ou réglez le locataire sur `consumers` (personnels seulement) ou le domaine de votre entreprise (professionnels seulement).
- **Pas de calendrier temporaire après l’activation** : rechargez la page ; si l’administrateur n’a pas encore saisi l’ID d’application, le calendrier temporaire le dit.
- **« Code expiré, réenregistrez Mes calendriers »** : personne ne s’est connecté pendant une heure : enregistrez « Mes calendriers » pour un nouveau code.
- **Un compte professionnel est refusé** : votre entreprise exige peut-être qu’un administrateur approuve l’application (consentement administrateur) dans Microsoft Entra.
- **Un calendrier partagé n’a aucun événement** : la personne qui l’a partagé ne partage peut-être que les disponibilités ; un calendrier que Microsoft refuse de lire est ignoré.
- Journaux : Paramètres → Supervision, les journaux de l’intégration. Ne partagez jamais un code ni un jeton.
