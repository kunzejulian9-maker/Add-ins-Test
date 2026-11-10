# DoorLabs Outlook-Add-in (Testversion)

Fügt beim Erstellen eines Termins in Outlook den Knopf **«DoorLabs-Schild»** hinzu. Im Seitenfenster wählt man Gerät, Status und Farbe (nur die 6 Spectra-Farben). Das Add-in schreibt die Wahl als Code ins Feld **«Ort»**, z. B.:

```
Sitzungszimmer 2 [DoorLabs: Im Meeting | Blau | Büro Julian]
```

Ohne Server liest später das Schild diesen Code aus dem veröffentlichten Kalender (ICS-Link).

## Dateien
| Datei | Zweck |
|---|---|
| `manifest.xml` | Sagt Outlook, wo das Add-in liegt und wo der Knopf erscheint |
| `taskpane.html` | Das Seitenfenster (alles in einer Datei) |
| `commands.html` | Pflichtdatei für Outlook, sonst leer |
| `assets/` | Symbole 16 bis 128 px |

## Was ihr braucht
- Ein **Microsoft-Konto mit Outlook**: privates Outlook.com/Hotmail-Konto oder Microsoft 365. Die Zoho-Adressen @doorlabs.ch gehen nicht. Schulkonten blockieren eigene Add-ins oft.
- Ein **GitHub-Konto** (gratis) zum Hosten. Outlook lädt Add-ins nur über HTTPS.

## Schritt 1: Auf GitHub Pages stellen
1. Auf github.com neues **öffentliches** Repository mit dem Namen `doorlabs-outlook-addin` anlegen.
2. «Add file» → «Upload files» → `taskpane.html`, `commands.html` und den Ordner `assets` hineinziehen → «Commit changes».
3. «Settings» → «Pages» → Source «Deploy from a branch», Branch `main`, Ordner `/ (root)` → «Save».
4. 1–2 Minuten warten, dann `https://EUER-NAME.github.io/doorlabs-outlook-addin/taskpane.html` öffnen. Ihr müsst das Seitenfenster mit dem gelben Hinweis «Vorschau ohne Outlook» sehen.

## Schritt 2: Manifest anpassen
In `manifest.xml` mit Suchen/Ersetzen **`GITHUB-NAME`** durch euren GitHub-Namen ersetzen (9 Stellen). Diese Datei muss nicht auf GitHub, sie wird nur in Outlook hochgeladen.

## Schritt 3: In Outlook laden
1. Outlook im Browser öffnen (outlook.live.com bzw. outlook.office.com) oder das neue Outlook für Windows.
2. Im Browser `https://aka.ms/olksideload` öffnen → «Meine Add-Ins».
3. Unten «Benutzerdefinierte Add-Ins» → «Benutzerdefiniertes Add-In hinzufügen» → «Aus Datei hinzufügen» → `manifest.xml` wählen → bestätigen.

## Schritt 4: Testen
1. Kalender → **Neuer Termin**.
2. In der Leiste oben **«DoorLabs-Schild»** anklicken (steht er nicht da: unter «…» oder «Apps»).
3. Status und Farbe wählen → **Übernehmen** → im Feld «Ort» steht jetzt der Code.
4. Termin speichern. Termin wieder öffnen und Knopf drücken: Die Auswahl ist wieder da.

## Danach testen (ICS)
Kalender veröffentlichen (Einstellungen → Kalender → Freigegebene Kalender → «Kalender veröffentlichen»), ICS-Link im Browser öffnen und prüfen, ob in der Zeile `LOCATION:` der DoorLabs-Code steht. Das hängt von der gewählten Detailstufe ab.

## Bekannte Grenzen
- Der Code ist für alle sichtbar, die den Termin sehen (auch Eingeladene).
- Kein Zwang: Wer den Knopf nicht drückt, hat keinen DoorLabs-Code im Termin.
- Wenn das Manifest nicht angenommen wird: auf dem PC `npx office-addin-manifest validate manifest.xml` ausführen (braucht Node.js), Fehlermeldung an Claude geben.
