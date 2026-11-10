# Outlook Calendar

## Übersicht

Diese Integration holt deine **Outlook-Kalender** in Gladys — Outlook.com, Hotmail, Live oder Microsoft 365 — über die offizielle Microsoft-Graph-API, nur lesend. **Jeder Gladys-Benutzer verknüpft sein eigenes Microsoft-Konto**: Seine Kalender bleiben seine, privat, bis er sie mit dem Haushalt teilt. Die Ereignisse erscheinen dann im Kalender von Gladys, starten Szenen mit „Ein Ereignis steht bevor“, und der Assistent kann sie lesen.

Wiederkehrende Ereignisse, ganztägige Ereignisse, mit dir geteilte Kalender und Links zu Online-Besprechungen werden unterstützt. Private Ereignisse können als „Beschäftigt“ angezeigt oder weggelassen werden, und abgelehnte Einladungen werden standardmäßig weggelassen.

Nicht mit Microsoft verbunden. Noch nicht mit einem echten Microsoft-Konto getestet: Rückmeldungen sind willkommen.

## Voraussetzungen

- Gladys Assistant 5.3.0 oder neuer.
- Ein Microsoft-Konto mit Kalender: persönlich (Outlook.com, Hotmail, Live) oder geschäftlich (Microsoft 365).
- Einmalig, für das ganze Zuhause: eine kostenlose **App-Registrierung** in Microsoft Entra, angelegt vom Administrator von Gladys (jedes Microsoft-Konto kann eine anlegen).
- **Deine Daten**: Die Kalenderereignisse jedes Benutzers liest Gladys bei Microsoft; in Outlook wird nie etwas geschrieben.

### Die App-Registrierung anlegen

1. Gehe zu [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App-Registrierungen** → **Neue Registrierung** (ein persönliches Microsoft-Konto genügt).
2. Name: `Gladys`. Unterstützte Kontotypen: **Konten in einem beliebigen Organisationsverzeichnis und persönliche Microsoft-Konten**. Umleitungs-URI: leer lassen. **Registrieren**.
3. **Authentifizierung** → **Erweiterte Einstellungen** → **Öffentliche Clientflows zulassen**: **Ja** → **Speichern**.
4. **API-Berechtigungen** → **Berechtigung hinzufügen** → **Microsoft Graph** → **Delegierte Berechtigungen**: `Calendars.Read`, `User.Read`, `offline_access` → **Berechtigungen hinzufügen**.
5. Kopiere die **Anwendungs-ID (Client)** von der Seite **Übersicht**.

Kein geheimer Clientschlüssel, keine Umleitungsadresse: Die Benutzer melden sich mit einem Code an, von jedem Gerät aus.

## Konfiguration

| Feld oder Schaltfläche (Administrator) | Standard   | Bedeutung                                                                                            |
| -------------------------------------- | ---------- | ---------------------------------------------------------------------------------------------------- |
| Anwendungs-ID (Client)                 | —          | Pflichtfeld: die ID deiner App-Registrierung                                                         |
| Mandant (Tenant)                       | common     | `common` für persönliche und geschäftliche Konten, `consumers` nur für persönliche, oder deine Firma |
| Synchronisieren alle                   | 15 Minuten | 5, 15, 30 Minuten oder 1 Stunde                                                                      |
| Ereignisse voraus                      | 12 Monate  | 3, 6 oder 12 Monate; die letzten 30 Tage bleiben ebenfalls                                           |
| App-Registrierung prüfen               | —          | Fragt Microsoft, ob die App Anmeldungen annimmt, und sagt sonst, was falsch ist                      |
| Jetzt synchronisieren                  | —          | Liest die Kalender aller verknüpften Benutzer erneut                                                 |

| Feld (jeder Benutzer, „Meine Kalender“) | Standard                    | Bedeutung                                                                      |
| --------------------------------------- | --------------------------- | ------------------------------------------------------------------------------ |
| Private Ereignisse                      | Als „Beschäftigt“ angezeigt | Als „Beschäftigt“ (ohne Titel und Ort), mit Details angezeigt oder weggelassen |
| Abgelehnte Einladungen                  | Weggelassen                 | Weggelassen oder angezeigt                                                     |

## Verwendung

1. Der Administrator: Integrationen → **Outlook Calendar** → Konfiguration: Anwendungs-ID einfügen, speichern, **App-Registrierung prüfen** drücken.
2. Jeder Benutzer: dieselbe Seite, Block **Meine Kalender** → **Aktivieren**.
3. Seine Kalender kommen an, privat; jeder kann im selben Block ausgeschaltet oder mit dem Haushalt geteilt werden.

### Dein Konto verknüpfen

Gladys hat noch keine „Verbinden“-Schaltfläche pro Benutzer, deshalb kommt der Anmeldecode in **einem vorübergehenden Kalender**, in deiner Liste „Meine Kalender“ (und als Ereignis in deiner Kalenderansicht), zum Beispiel so benannt:

> Microsoft: gib ABCD-EFGH auf microsoft.com/devicelogin ein

Öffne [microsoft.com/devicelogin](https://microsoft.com/devicelogin) auf einem beliebigen Computer oder Telefon, gib den Code ein, melde dich mit deinem Microsoft-Konto an und akzeptiere die Berechtigungen (deine Kalender lesen, Zugriff behalten). Innerhalb weniger Sekunden verschwindet der vorübergehende Kalender und deine Outlook-Kalender erscheinen. Der Code gilt etwa 15 Minuten und wird eine Stunde lang von selbst erneuert; danach speichere „Meine Kalender“ erneut, um einen neuen zu erhalten.

Um ein anderes Microsoft-Konto zu verknüpfen, deaktiviere „Meine Kalender“ und aktiviere es wieder. Akzeptiert Microsoft die Anmeldung nicht mehr (Passwort geändert, Zugriff entzogen, 90 Tage ungenutzt), kommt der vorübergehende Kalender mit einem neuen Code zurück.

### Dashboard-Widget

Das Widget **Outlook-Kalender** zeigt, wie viele Benutzer verknüpft sind, die synchronisierten Kalender und Ereignisse und eine Zeile pro Benutzer, der „Meine Kalender“ aktiviert hat: synchronisiert um welche Uhrzeit, wartet auf Anmeldung oder Fehler. Nur Vornamen: keine E-Mail, kein Ereignis. Eine Schaltfläche synchronisiert sofort.

![Das Widget Outlook-Kalender: verknüpfte Benutzer, Kalender und Ereignisse, der Zustand jedes Benutzers und eine Sync-Taste (generierte Illustration, kein Screenshot)](https://raw.githubusercontent.com/William-De71/gladys-outlook-calendar/main/docs/assets/widgets.de.png)

### Szenen

Die Ereignisse selbst starten Szenen mit dem Gladys-Auslöser **„Ein Kalenderereignis steht bevor“** (bei geteilten Kalendern). Die Integration ergänzt:

- **Auslöser „Outlook: Ein Benutzer muss sich anmelden“**: Ein Benutzer hat „Meine Kalender“ aktiviert (`new`) oder Microsoft akzeptiert seine Anmeldung nicht mehr (`revoked`), mit seinem Vornamen.
- **Auslöser „Outlook: Ein Benutzer hat sich angemeldet“**: mit seinem Vornamen und seinem Microsoft-Konto.
- **Auslöser „Outlook: Synchronisierung schlägt fehl“**: einmal pro Ausfall und Benutzer (Zugriff verweigert, kein Postfach, gedrosselt, Microsoft nicht erreichbar, Microsoft-Fehler).
- **Auslöser „Outlook: Synchronisierung geht wieder“**: mit den Minuten des Ausfalls.
- **Aktion „Outlook: jetzt synchronisieren“**: vor einer Szene, die die heutigen Ereignisse prüft.

Beispiele: Wenn ein Benutzer sich erneut anmelden muss (`revoked`), ihm eine Telegram-Nachricht „Dein Outlook-Link ist abgelaufen, sieh in Meine Kalender nach“ schicken; jeden Morgen um 7:00 jetzt synchronisieren und dann die heutigen Ereignisse vorlesen.

## Fehlerbehebung

- **„Microsoft lehnt die App-Registrierung ab: AADSTS700016“**: Die Anwendungs-ID ist falsch, oder die App wurde gelöscht.
- **„AADSTS7000218“** oder ein geheimer Clientschlüssel wird verlangt: **Öffentliche Clientflows zulassen** steht nicht auf Ja (Schritt 3).
- **„AADSTS50020“** oder **„AADSTS9002331“**: Die Kontotypen der App passen nicht zum Mandanten: Wähle „beliebiges Organisationsverzeichnis und persönliche Microsoft-Konten“, oder stelle den Mandanten auf `consumers` (nur persönliche) oder die Domain deiner Firma (nur geschäftliche).
- **Kein vorübergehender Kalender nach dem Aktivieren**: Lade die Seite neu; hat der Administrator die Anwendungs-ID noch nicht eingegeben, sagt es der vorübergehende Kalender.
- **„Code abgelaufen, Meine Kalender neu speichern“**: Eine Stunde lang hat sich niemand angemeldet: Speichere „Meine Kalender“ für einen neuen Code.
- **Ein Geschäftskonto wird abgelehnt**: Deine Firma verlangt vielleicht, dass ein Administrator die App in Microsoft Entra genehmigt (Administratorzustimmung).
- **Ein geteilter Kalender hat keine Ereignisse**: Die Person, die ihn geteilt hat, teilt vielleicht nur Frei/Gebucht-Zeiten; ein Kalender, den Microsoft nicht lesen lässt, wird übersprungen.
- Protokolle: Einstellungen → Supervision, die Protokolle der Integration. Teile niemals einen Code oder ein Token.
