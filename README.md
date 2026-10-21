# Brim · Outlook

Plugin für [Brim](https://github.com/wheatley2056/brim), die Notch für Windows: verbindet **E-Mails, Termine und Aufgaben** aus Outlook mit der Notch.

| Wo | Was |
|---|---|
| **Tab „Jetzt“** | Alles in einer scrollbaren Liste: der anstehende Termin groß mit **Beitreten** (Teams, Zoom, Google Meet, Webex), dann **Heute** (bzw. **Morgen**, wenn heute nichts mehr ansteht), **Posteingang** und **Aufgaben**. Jeweils öffnen, neu anlegen, als gelesen bzw. erledigt markieren, Ordner in Outlook zeigen. |
| **Notch, kompakt** | Der nächste Termin erscheint 15 Minuten vorher mit Countdown („in 12 min“) und bleibt bis 5 Minuten nach Beginn. Zu Beginn kurz als HUD: „Teamrunde beginnt jetzt“. Sonst bleibt die kompakte Notch frei. |
| **HUD** | Neue E-Mail: „Anna Müller: Angebot Q4“. |
| **Widgets auf „Jetzt“** (Brim 1.2) | **Nächster Termin** (klein, mittel, groß – groß mit den folgenden), **Posteingang** (klein mit Anzahl, mittel, groß mit den neuesten) und **Aufgaben** (mittel, groß, zum Abhaken). Hinzufügen: „Jetzt“ lange drücken oder rechts anklicken, dann „+“. |

Brim zeigt im Tab „Jetzt“ die laufende Aktivität mit dem höchsten Rang. Das Plugin nimmt den höchsten, den Brim Plugins erlaubt (9), und steht damit vor Akku, Kamera und pausierter Musik. **Läuft Musik oder ein Timer, zeigt „Jetzt“ diese**, bis sie pausiert bzw. abgelaufen sind.

Läuft Outlook nicht, steht das im Tab „Jetzt“, mit **Outlook starten**.

## Voraussetzungen

- Windows 10/11 mit Brim
- **Klassisches Outlook für Windows** (Microsoft 365 / Outlook 2016 und neuer). Das Plugin spricht das Objektmodell von Outlook über COM an und hängt sich an die laufende Instanz; es startet Outlook nie von selbst und braucht kein Konto, keine App-Registrierung und kein Netzwerk.
- Das *neue* Outlook (`olk.exe`) bietet keine COM-Schnittstelle und wird nicht unterstützt.
- .NET 10 Desktop Runtime (bringt Brim mit)

## Installieren

Am einfachsten über die Plugin-Bibliothek ab Brim 2.3: Einstellungen → Plugins → Bibliothek → **Outlook** → Installieren.

Von Hand aus dem Quellcode:

```powershell
git clone --recurse-submodules https://github.com/wheatley2056/brim-outlook.git
cd brim-outlook
dotnet publish src/BrimOutlook -c Release -r win-x64 --self-contained false -o "$env:LOCALAPPDATA\Brim\Plugins\outlook"
```

Danach Brim neu starten, Notch aufklappen, Tab **Plugins**: rechts **Outlook** einschalten.

Das Plugin bittet um das Recht **„Benachrichtigungen bestimmter Apps lesen“** für Outlook. Damit gehen Outlooks eigene Windows-Benachrichtigungen (neue E-Mail, Erinnerungen) an das Plugin statt an Brims Vorschau und erscheinen als ein HUD statt doppelt. Dafür muss Brims Identität eingerichtet sein (Einstellungen → Kalender und Benachrichtigungen); ohne sie meldet das Plugin neue E-Mails selbst.

## Release

Tag `v<Version>` setzen, passend zu `version` in `src/BrimOutlook/plugin.json`. Die Action hängt dann `brim-outlook.zip` an das Release. Danach im Brim-Repository `./build/catalog/Add-Plugin.ps1 -Repository wheatley2056/brim-outlook -Tag v<Version>` ausführen und den Katalog per Pull Request aktualisieren.

Ausgaben stehen in `%LocalAppData%\Brim\logs\plugins\de.brim.outlook.log`.

## Einstellungen

In Brims `settings.json` unter `Plugins.de.brim.outlook.Values`:

| Schlüssel | Standard | Bedeutung |
|---|---|---|
| `refreshSeconds` | `30` | Outlook so oft abfragen (10–600 s). Der Countdown läuft unabhängig davon alle 15 s. |
| `meetingLeadMinutes` | `15` | So viele Minuten vor Beginn erscheint ein Termin in der Notch. |
| `maxItems` | `10` | Einträge pro Abschnitt (1–12; mehr passt nicht in Brims Grenze von 400 Bausteinen pro Nachricht). |
| `showUnreadInNotch` | `false` | Anzahl ungelesener E-Mails kompakt in der Notch zeigen, solange kein Termin ansteht. |
| `mailHud` | `true` | Neue E-Mails als HUD melden. |
| `meetingActivity` | `true` | Termin-Countdown in der Notch. Aus, wenn Brims eingebauter Kalender dieselben Termine schon zeigt (z. B. weil das Konto auch im Windows-Kalender steckt). |

## Was gelesen wird

- **Posteingang:** ungelesene Elemente, neueste zuerst (Absender, Betreff, Empfangszeit, Wichtigkeit).
- **Kalender:** Termine von jetzt bis morgen Abend im Standardkalender, Serien als einzelne Termine. Abgesagte und abgelehnte Termine fallen weg. Den Link zum Beitreten liefert bei Teams die Eigenschaft des Termins, sonst wird er in Ort und Text gesucht; Microsoft-Safe-Links werden ausgepackt.
- **Aufgaben:** nicht erledigte Aufgaben im Standard-Aufgabenordner, überfällige zuerst.

Nichts davon verlässt den Rechner; das Plugin schickt nur die angezeigten Texte an Brim. Aktions-IDs enthalten kurze Schlüssel statt Outlook-EntryIDs.

**Hinweis zu Sicherheitsabfragen:** Absender und Termintext gehören zu den geschützten Eigenschaften von Outlook. Mit aktuellem Virenschutz (auch Microsoft Defender) greift das Plugin ohne Nachfrage darauf zu. Ist kein Virenschutz aktiv, kann Outlook nachfragen; lehnst du ab, fehlen nur Absender und Beitreten-Links.

## Aufbau

```
src/BrimOutlook/
  Program.cs              Verbindung zu Brim (Brim.Sdk), Ereignisse verdrahten
  OutlookPlugin.cs        Abfrage-Schleife, Aktionen, neue E-Mails, schickt nur Änderungen
  Views.cs                Abschnitte und die Aktivität für „Jetzt“ als Widgets (reine Funktionen)
  Outlook/                Modell, Schnittstelle IOutlookSource, Links und Zeitangaben
  Com/                    Outlook über COM: eigener STA-Thread, Nachrichtenfilter für „Outlook ist beschäftigt“
external/brim             Brim als Submodul (für Brim.Sdk)
tests/BrimOutlook.Tests   xUnit; prüft u. a. alles Gesendete mit Brims WidgetValidator
```

Entwickeln und testen geht auch ohne Windows (`dotnet test`); nur `Com/` braucht Outlook.

## Lizenz

GPL-3.0-or-later, wie Brim und `Brim.Sdk`.
