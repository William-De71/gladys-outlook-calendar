// -----------------------------------------------------------------------------
// Texts of the integration, in English, French, German and Spanish (English
// fallback).
//
// The manifest carries its own labels; these are the texts the code writes at
// runtime: the connection status, the results of the buttons, the widget, and
// the texts users see in their own calendars (the sign-in calendar, private
// events), in the language of each user. German says "du", Spanish "usted".
// -----------------------------------------------------------------------------

const LANGUAGES = ['en', 'fr', 'de', 'es'];

const TEXTS = {
  // connection status and buttons (administrator)
  statusNoApp: {
    en: 'Enter the application (client) ID of your Microsoft Entra app registration: users cannot sign in before.',
    fr: 'Saisissez l’ID d’application (client) de votre inscription d’application Microsoft Entra : les utilisateurs ne peuvent pas se connecter avant.',
    de: 'Gib die Anwendungs-ID (Client) deiner App-Registrierung in Microsoft Entra ein: Vorher können sich die Benutzer nicht anmelden.',
    es: 'Introduzca el ID de aplicación (cliente) de su registro de aplicación en Microsoft Entra: los usuarios no pueden iniciar sesión antes.',
  },
  statusBadClientId: {
    en: 'The application (client) ID is not a valid ID: copy it again from the Overview page of the app registration.',
    fr: 'L’ID d’application (client) n’est pas un identifiant valide : recopiez-le depuis la page Vue d’ensemble de l’inscription d’application.',
    de: 'Die Anwendungs-ID (Client) ist keine gültige ID: Kopiere sie erneut von der Übersichtsseite der App-Registrierung.',
    es: 'El ID de aplicación (cliente) no es un identificador válido: cópielo de nuevo desde la página Información general del registro de aplicación.',
  },
  statusBadApp: {
    en: 'Microsoft refuses the app registration: {detail}',
    fr: 'Microsoft refuse l’inscription d’application : {detail}',
    de: 'Microsoft lehnt die App-Registrierung ab: {detail}',
    es: 'Microsoft rechaza el registro de aplicación: {detail}',
  },
  statusUnreachable: {
    en: 'Cannot reach login.microsoftonline.com: check the internet access of Gladys.',
    fr: 'Impossible de joindre login.microsoftonline.com : vérifiez l’accès Internet de Gladys.',
    de: 'login.microsoftonline.com ist nicht erreichbar: Prüfe den Internetzugang von Gladys.',
    es: 'No se puede acceder a login.microsoftonline.com: compruebe el acceso a Internet de Gladys.',
  },
  statusInitFailed: {
    en: 'Initialization failed, check the integration logs.',
    fr: 'L’initialisation a échoué, consultez les journaux de l’intégration.',
    de: 'Die Initialisierung ist fehlgeschlagen, prüfe die Protokolle der Integration.',
    es: 'La inicialización ha fallado, consulte los registros de la integración.',
  },
  checkOk: {
    en: 'The app registration is accepted: users can sign in from “My calendars”.',
    fr: 'L’inscription d’application est acceptée : les utilisateurs peuvent se connecter depuis « Mes calendriers ».',
    de: 'Die App-Registrierung wird akzeptiert: Die Benutzer können sich unter „Meine Kalender“ anmelden.',
    es: 'El registro de aplicación es aceptado: los usuarios pueden iniciar sesión desde «Mis calendarios».',
  },
  syncDone: {
    en: '{users} user(s): {calendars} calendar(s) and {events} event(s) synchronized.',
    fr: '{users} utilisateur(s) : {calendars} calendrier(s) et {events} événement(s) synchronisés.',
    de: '{users} Benutzer: {calendars} Kalender und {events} Ereignis(se) synchronisiert.',
    es: '{users} usuario(s): {calendars} calendario(s) y {events} evento(s) sincronizados.',
  },
  toastSync: {
    en: 'Synchronization started.',
    fr: 'Synchronisation lancée.',
    de: 'Synchronisierung gestartet.',
    es: 'Sincronización iniciada.',
  },

  // the sign-in calendar (each user, in their language)
  link_code: {
    en: 'Microsoft: type {code} at {url}',
    fr: 'Microsoft : tapez {code} sur {url}',
    de: 'Microsoft: gib {code} auf {url} ein',
    es: 'Microsoft: escriba {code} en {url}',
  },
  link_code_description: {
    en: 'To link your Outlook calendars, open {url} (computer or phone), type the code {code} and sign in with your Microsoft account. The code is valid until {time}; a new one follows on its own. This temporary calendar disappears once you are signed in.',
    fr: 'Pour relier vos calendriers Outlook, ouvrez {url} (ordinateur ou téléphone), tapez le code {code} et connectez-vous avec votre compte Microsoft. Le code est valable jusqu’à {time} ; un nouveau suit tout seul. Ce calendrier temporaire disparaît une fois la connexion faite.',
    de: 'Um deine Outlook-Kalender zu verknüpfen, öffne {url} (Computer oder Telefon), gib den Code {code} ein und melde dich mit deinem Microsoft-Konto an. Der Code gilt bis {time}; danach folgt von selbst ein neuer. Dieser vorübergehende Kalender verschwindet nach der Anmeldung.',
    es: 'Para vincular sus calendarios de Outlook, abra {url} (ordenador o teléfono), escriba el código {code} e inicie sesión con su cuenta de Microsoft. El código es válido hasta las {time}; después llega uno nuevo solo. Este calendario temporal desaparece una vez iniciada la sesión.',
  },
  link_no_app: {
    en: 'Microsoft: waiting for the administrator',
    fr: 'Microsoft : en attente de l’administrateur',
    de: 'Microsoft: Warten auf den Administrator',
    es: 'Microsoft: esperando al administrador',
  },
  link_no_app_description: {
    en: 'The administrator of Gladys must first enter the Microsoft application ID in the configuration of the integration. Then save “My calendars” again to get your sign-in code.',
    fr: 'L’administrateur de Gladys doit d’abord saisir l’ID d’application Microsoft dans la configuration de l’intégration. Enregistrez ensuite « Mes calendriers » à nouveau pour obtenir votre code de connexion.',
    de: 'Der Administrator von Gladys muss zuerst die Microsoft-Anwendungs-ID in der Konfiguration der Integration eingeben. Speichere dann „Meine Kalender“ erneut, um deinen Anmeldecode zu erhalten.',
    es: 'El administrador de Gladys debe introducir primero el ID de aplicación de Microsoft en la configuración de la integración. Después guarde de nuevo «Mis calendarios» para obtener su código de inicio de sesión.',
  },
  link_app: {
    en: 'Microsoft: the app registration refuses the sign-in',
    fr: 'Microsoft : l’inscription d’application refuse la connexion',
    de: 'Microsoft: Die App-Registrierung lehnt die Anmeldung ab',
    es: 'Microsoft: el registro de aplicación rechaza el inicio de sesión',
  },
  link_app_description: {
    en: 'Microsoft refused to give a sign-in code for the application set by the administrator (public client flows off, personal accounts not allowed…). Ask them to check it, then save “My calendars” again.',
    fr: 'Microsoft a refusé de donner un code de connexion pour l’application réglée par l’administrateur (flux client public désactivé, comptes personnels non autorisés…). Demandez-lui de la vérifier, puis enregistrez « Mes calendriers » à nouveau.',
    de: 'Microsoft hat für die vom Administrator eingestellte Anwendung keinen Anmeldecode ausgegeben (öffentliche Clientflows aus, persönliche Konten nicht erlaubt…). Bitte ihn, sie zu prüfen, und speichere dann „Meine Kalender“ erneut.',
    es: 'Microsoft se ha negado a dar un código de inicio de sesión para la aplicación configurada por el administrador (flujos de cliente público desactivados, cuentas personales no permitidas…). Pídale que la revise y luego guarde de nuevo «Mis calendarios».',
  },
  link_unavailable: {
    en: 'Microsoft: sign-in unavailable for now',
    fr: 'Microsoft : connexion indisponible pour l’instant',
    de: 'Microsoft: Anmeldung gerade nicht verfügbar',
    es: 'Microsoft: inicio de sesión no disponible por ahora',
  },
  link_unavailable_description: {
    en: 'Gladys could not reach Microsoft to get your sign-in code. Save “My calendars” again in a moment.',
    fr: 'Gladys n’a pas pu joindre Microsoft pour obtenir votre code de connexion. Enregistrez « Mes calendriers » à nouveau dans un instant.',
    de: 'Gladys konnte Microsoft nicht erreichen, um deinen Anmeldecode zu holen. Speichere „Meine Kalender“ gleich noch einmal.',
    es: 'Gladys no ha podido contactar con Microsoft para obtener su código de inicio de sesión. Guarde de nuevo «Mis calendarios» dentro de un momento.',
  },
  link_declined: {
    en: 'Microsoft: sign-in declined',
    fr: 'Microsoft : connexion refusée',
    de: 'Microsoft: Anmeldung abgelehnt',
    es: 'Microsoft: inicio de sesión rechazado',
  },
  link_declined_description: {
    en: 'The sign-in was declined on the Microsoft page. Save “My calendars” again to get a new code.',
    fr: 'La connexion a été refusée sur la page Microsoft. Enregistrez « Mes calendriers » à nouveau pour obtenir un nouveau code.',
    de: 'Die Anmeldung wurde auf der Microsoft-Seite abgelehnt. Speichere „Meine Kalender“ erneut, um einen neuen Code zu erhalten.',
    es: 'El inicio de sesión se ha rechazado en la página de Microsoft. Guarde de nuevo «Mis calendarios» para obtener un código nuevo.',
  },
  link_expired: {
    en: 'Microsoft: code expired, save My calendars again',
    fr: 'Microsoft : code expiré, réenregistrez Mes calendriers',
    de: 'Microsoft: Code abgelaufen, Meine Kalender neu speichern',
    es: 'Microsoft: código caducado, guarde Mis calendarios de nuevo',
  },
  link_expired_description: {
    en: 'No sign-in for an hour: the codes stopped. Save “My calendars” again to get a new code.',
    fr: 'Aucune connexion depuis une heure : les codes se sont arrêtés. Enregistrez « Mes calendriers » à nouveau pour obtenir un nouveau code.',
    de: 'Seit einer Stunde keine Anmeldung: Die Codes wurden angehalten. Speichere „Meine Kalender“ erneut, um einen neuen Code zu erhalten.',
    es: 'Ningún inicio de sesión desde hace una hora: los códigos se han detenido. Guarde de nuevo «Mis calendarios» para obtener un código nuevo.',
  },

  // calendars and events
  untitledCalendar: { en: 'Calendar', fr: 'Calendrier', de: 'Kalender', es: 'Calendario' },
  calendarOf: {
    en: 'Outlook calendar of {account}',
    fr: 'Calendrier Outlook de {account}',
    de: 'Outlook-Kalender von {account}',
    es: 'Calendario de Outlook de {account}',
  },
  calendarSharedBy: {
    en: 'Outlook calendar shared by {owner}',
    fr: 'Calendrier Outlook partagé par {owner}',
    de: 'Von {owner} geteilter Outlook-Kalender',
    es: 'Calendario de Outlook compartido por {owner}',
  },
  privateEvent: { en: 'Busy', fr: 'Occupé', de: 'Beschäftigt', es: 'Ocupado' },
  untitledEvent: { en: '(No title)', fr: '(Sans titre)', de: '(Ohne Titel)', es: '(Sin título)' },

  // widget
  widgetNoApp: {
    en: 'The administrator must first enter the Microsoft application ID in the configuration of the integration.',
    fr: 'L’administrateur doit d’abord saisir l’ID d’application Microsoft dans la configuration de l’intégration.',
    de: 'Der Administrator muss zuerst die Microsoft-Anwendungs-ID in der Konfiguration der Integration eingeben.',
    es: 'El administrador debe introducir primero el ID de aplicación de Microsoft en la configuración de la integración.',
  },
  widgetNoUser: {
    en: 'Nobody has enabled “My calendars” yet: each user does it on the Configuration page of the integration.',
    fr: 'Personne n’a encore activé « Mes calendriers » : chaque utilisateur le fait dans la page Configuration de l’intégration.',
    de: 'Noch niemand hat „Meine Kalender“ aktiviert: Jeder Benutzer tut das auf der Konfigurationsseite der Integration.',
    es: 'Nadie ha activado todavía «Mis calendarios»: cada usuario lo hace en la página Configuración de la integración.',
  },
  tileLinked: { en: 'Linked users', fr: 'Comptes reliés', de: 'Verknüpft', es: 'Vinculados' },
  tileCalendars: { en: 'Calendars', fr: 'Calendriers', de: 'Kalender', es: 'Calendarios' },
  tileEvents: { en: 'Events', fr: 'Événements', de: 'Ereignisse', es: 'Eventos' },
  rowLinking: {
    en: 'Waiting for sign-in',
    fr: 'En attente de connexion',
    de: 'Wartet auf Anmeldung',
    es: 'Esperando inicio de sesión',
  },
  rowNoApp: {
    en: 'Waiting for the app ID',
    fr: 'En attente de l’ID d’app',
    de: 'Wartet auf die App-ID',
    es: 'Esperando el ID de app',
  },
  rowLinkStopped: {
    en: 'Sign-in stopped',
    fr: 'Connexion arrêtée',
    de: 'Anmeldung angehalten',
    es: 'Inicio de sesión detenido',
  },
  rowError: {
    en: 'Error: {kind}',
    fr: 'Erreur : {kind}',
    de: 'Fehler: {kind}',
    es: 'Error: {kind}',
  },
  rowSynced: {
    en: 'Synced at {time}',
    fr: 'Synchro à {time}',
    de: 'Synchron. um {time}',
    es: 'Sincronizado a las {time}',
  },
  rowLinked: { en: 'Linked', fr: 'Relié', de: 'Verknüpft', es: 'Vinculado' },
  buttonSync: { en: 'Sync now', fr: 'Synchroniser', de: 'Jetzt synchron.', es: 'Sincronizar' },
  someone: { en: 'User', fr: 'Utilisateur', de: 'Benutzer', es: 'Usuario' },
  kind_auth: {
    en: 'sign-in refused',
    fr: 'connexion refusée',
    de: 'Anmeldung abgelehnt',
    es: 'sesión rechazada',
  },
  kind_app: { en: 'app refused', fr: 'app refusée', de: 'App abgelehnt', es: 'app rechazada' },
  kind_forbidden: {
    en: 'access denied',
    fr: 'accès refusé',
    de: 'Zugriff verweigert',
    es: 'acceso denegado',
  },
  kind_not_found: {
    en: 'no mailbox',
    fr: 'pas de boîte mail',
    de: 'kein Postfach',
    es: 'sin buzón',
  },
  kind_throttled: { en: 'throttled', fr: 'trop de requêtes', de: 'gedrosselt', es: 'limitado' },
  kind_server: {
    en: 'Microsoft error',
    fr: 'erreur Microsoft',
    de: 'Microsoft-Fehler',
    es: 'error de Microsoft',
  },
  kind_unreachable: {
    en: 'unreachable',
    fr: 'injoignable',
    de: 'nicht erreichbar',
    es: 'inaccesible',
  },
};

/**
 * @description Whether a value is a language of the integration.
 * @param {unknown} value - The value.
 * @returns {boolean} True for en, fr, de, es.
 */
function isLanguage(value) {
  return LANGUAGES.includes(value);
}

/**
 * @description A text in a language, with its values.
 * @param {string} key - A TEXTS key.
 * @param {string} language - The language (English when unknown).
 * @param {object} [values] - The values of the `{placeholders}`.
 * @returns {string} The text, the key itself when unknown.
 */
function t(key, language, values = {}) {
  const entry = TEXTS[key];
  if (!entry) {
    return key;
  }
  const template = entry[isLanguage(language) ? language : 'en'] ?? entry.en;
  return template.replace(/\{(\w+)\}/g, (whole, name) =>
    values[name] === undefined || values[name] === null ? '' : String(values[name]),
  );
}

/**
 * @description A text in the four languages.
 * @param {string|Function} key - A TEXTS key, or `(language) => text`.
 * @param {object|Function} [values] - The values, or `(language) => values`.
 * @returns {object} `{ en, fr, de, es }`.
 */
function all(key, values = {}) {
  return Object.fromEntries(
    LANGUAGES.map((language) => [
      language,
      typeof key === 'function'
        ? key(language)
        : t(key, language, typeof values === 'function' ? values(language) : values),
    ]),
  );
}

/**
 * @description A number in the format of a language.
 * @param {number} value - The number.
 * @param {string} language - The language.
 * @param {number} [decimals] - The decimals at most.
 * @returns {string} "12,5" in French.
 */
function formatNumber(value, language, decimals = 1) {
  if (!Number.isFinite(value)) {
    return '–';
  }
  return new Intl.NumberFormat(isLanguage(language) ? language : 'en', {
    maximumFractionDigits: decimals,
  }).format(value);
}

export { LANGUAGES, TEXTS, isLanguage, t, all, formatNumber };
