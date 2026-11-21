# GFS Studio

Eine klar gestaltete Desktop-Präsentationsvorlage für eine GFS. Das reduzierte Layout setzt auf Papierweiß, Schwarz und Rot statt auf typische KI-Verläufe.

Die App läuft komplett im Browser. Es ist kein Server nötig.

## Starten

Die App nutzt npm-Pakete und muss vor dem Start gebaut werden:

```bash
npm ci
npm run build
npm run preview
```

Danach `http://localhost:3000` im Browser öffnen. Der Ordner `dist/` enthält die fertige Seite und lässt sich auch mit jedem anderen statischen Webserver ausliefern.

## Bedienung

- **Texte bearbeiten:** Auf einen Text klicken. Änderungen werden lokal im Browser gespeichert.
- **Elemente:** Über das Menü **Elemente** Texte, Formen, Bilder oder Videos (YouTube/Vimeo-Link) hinzufügen. Elemente lassen sich ziehen, über die Liste auswählen und per Griff unten rechts in Breite und Höhe ändern.
- **Folien:** Über **KI & Folien → Folien** Layouts wie Titel, Kapiteltrenner, Zitat, zwei Spalten, Bild, Video oder Fazit hinzufügen. Die Vorlage **Leer** eignet sich für eigene Inhalte.
- **Design und Übergänge:** Unter **Design** sechs Farbwelten, unter **Übergang** pro Folie einen Übergang.
- **Präsentieren:** **F** für Vollbild, **P** für den Präsentationsmodus, **Esc** beendet ihn.
- **Projekt sichern / laden:** Eine JSON-Datei mit allen Texten, Elementen und Einstellungen.

## Export und Import

**Export** (Menü **Datei exportieren**):

- **PPTX** und **ODP/OTP:** Der Folienhintergrund (Formen, Dekorationen, Bilder) wird als exaktes Bild übernommen. Texte sind echte Textfelder mit derselben Position, Größe, Farbe, Schrift und denselben Zeilenumbrüchen wie im Editor.
- **PDF:** Jede Folie wird als vollständiges, exaktes Bild exportiert.

Grenzen des Exports:

- Formen und Dekorationen sind im Hintergrundbild und nicht einzeln bearbeitbar.
- Gedrehte Texte (z. B. die Haftnotizen der Vorlage) bleiben ebenfalls im Hintergrundbild.
- Videos und 3D-Modelle werden durch beschriftete Platzhalter ersetzt. Animationen und Übergänge werden nicht übernommen.
- Die Schriften der Vorlage sind Arial, Georgia und Courier. Auf Systemen ohne diese Schriften weicht die Darstellung entsprechend ab.

**Import** (Menü **Datei öffnen**):

- **PPTX und ODP:** Text, Formen und Bilder werden als bearbeitbare Elemente übernommen. Gruppen, Verläufe, gedrehte Elemente und Formate wie EMF/WMF werden übersprungen.
- **PDF:** Jede Seite wird als Bild in eine Folie eingefügt. Der Text bleibt dabei im Bild und ist nicht bearbeitbar.
- Beim Import werden die Zusatzfolien durch die Datei ersetzt. Die Startfolien bleiben erhalten. Vorher erscheint eine Bestätigung.

Die Browser-Speicherung umfasst etwa 5 MB. Sehr große Präsentationen mit vielen Bildern lassen sich nicht vollständig im Browser speichern. Der Export zeigt in diesem Fall eine Meldung. Sichere dann die Projektdatei.

## KI-Assistent (OpenRouter)

Unter **KI-Einstellungen** trägst du deinen OpenRouter-Schlüssel ein. Die Anfragen gehen direkt aus dem Browser an OpenRouter. Es ist kein Server nötig.

- Der Schlüssel bleibt lokal im Browser (wahlweise nur für die aktuelle Sitzung) und wird nicht in Projektdateien exportiert.
- **Schlüssel niemals in den Quellcode oder ein öffentliches Repository schreiben.**
- Der Chat berät, ohne ungefragt Folien oder Texte einzufügen. Mit **Werkzeuge erlauben** kann er eine Folienübersicht abrufen, Design oder Übergang ändern oder Wikipedia durchsuchen.
- **Info-Check** markiert Quellenlücken, prüft aber keine Fakten live.

## Veröffentlichen

Die Vorlage braucht zum Betrieb nur statische Dateien aus `dist/`.

**Render (Web Service):**

- **Build Command:** `npm ci && npm run build`
- **Start Command:** `npm start` (liefert `dist/` auf dem von Render gesetzten `$PORT` aus)
- **Node.js:** Version 20 oder neuer

**Render (Static Site)** oder **Cloudflare Pages, Netlify, GitHub Pages:**

- **Build Command:** `npm ci && npm run build`
- **Publish Directory:** `dist`

Alle Build-Werkzeuge stehen in `dependencies`, damit `npm ci` auch im Production-Modus alles installiert.

Für die 3D-Ansicht wird beim Öffnen eine Bibliothek aus dem Google-CDN geladen. Dafür ist eine Internetverbindung nötig.

## Entwicklung

```bash
npm test              # Unit-Tests (Vitest)
npm run build         # Typprüfung und Build nach dist/
```

Die Export- und Import-Formate liegen in `presentation-io.js` ohne DOM-Abhängigkeit und sind dort getestet (`tests/presentation-io.test.ts`).

## Lizenz

Siehe `LICENSE`.
