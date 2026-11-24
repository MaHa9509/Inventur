# Büro-Inventur Technik

Offline-Web-App (PWA) für die Inventur der Bürotechnik auf dem Handy. Erfasste Geräte lassen sich als formatierte Excel-Datei (.xlsx) exportieren und von mehreren Handys zusammenführen.

- Läuft auf iPhone (Safari) und Android (Chrome), nach dem ersten Laden **ohne Internet**.
- Die Daten bleiben auf dem jeweiligen Handy (IndexedDB im Browser). Es gibt keinen Server und keine Cloud.
- IDs werden automatisch vergeben: `<Kürzel des Handys>-<Nummer>`, z. B. `A-0001`, `B-0001`.
- Die Räume aus dem Raumplan sind schon eingetragen (4070–4300 und „Lager / Sonstiges“). Die Teams heißen zunächst neutral „Team 1“ bis „Team 4“ (Farben wie im Plan).

---

## 1. Bereitstellen (einmalig)

Die App besteht nur aus statischen Dateien. Sie muss **einmal über HTTPS** erreichbar sein. Danach speichert jedes Handy sie für den Offline-Betrieb.

**Variante A: Netlify Drop (am einfachsten, kostenlos)**
1. <https://app.netlify.com/drop> öffnen (kostenloses Konto).
2. Den Ordner `inventur-app` per Drag & Drop in das Browserfenster ziehen.
3. Netlify zeigt eine Adresse wie `https://xyz.netlify.app`. Diesen Link an alle schicken, die erfassen.

**Variante B: GitHub Pages**
1. Neues Repository anlegen und den Inhalt des Ordners `inventur-app` hochladen.
2. *Settings → Pages → Branch: main / root* aktivieren.
3. Die App ist dann unter `https://<name>.github.io/<repo>/` erreichbar.

**Variante C: Interner Webserver / Intranet**
Den Ordner auf einen beliebigen Webserver mit **HTTPS** legen (IIS, Apache, nginx …). Eine Serverlogik ist nicht nötig.

> Wenn man `index.html` einfach als Datei auf dem Handy öffnet, funktioniert die App **nicht** zuverlässig. Der Offline-Modus und das dauerhafte Speichern brauchen HTTPS.

## 2. Auf dem Handy installieren

1. Link im Browser öffnen (iPhone: **Safari**, Android: **Chrome**).
2. Auf den Home-Bildschirm legen:
   - **iPhone:** Teilen-Symbol → „Zum Home-Bildschirm“
   - **Android:** Menü ⋮ → „App installieren“ bzw. „Zum Startbildschirm hinzufügen“
3. Ab jetzt die App **über das Symbol auf dem Home-Bildschirm** starten.

> Wichtig für iPhones: Safari kann die Daten von Websites löschen, die 7 Tage nicht benutzt wurden. Für Apps auf dem Home-Bildschirm gilt das nicht. Die App also unbedingt über das Symbol nutzen und regelmäßig exportieren.

## 3. Einrichtung pro Handy: Kürzel vergeben!

Beim ersten Start fragt die App:
- **Kürzel** (1–3 Buchstaben), z. B. `A`, `B`, `C`
- **Name** der erfassenden Person

Danach einmalig unter **Einstellungen → Teams** die Teamnamen eintragen (Namen antippen zum Umbenennen). Die Farben entsprechen dem Raumplan:

| Team | Farbe | Räume |
|---|---|---|
| Team 1 | orange | 4180, 4220, 4240, 4110 |
| Team 2 | hellgrün | 4190, 4170, 4160 |
| Team 3 | hellblau | 4260–4300 |
| Team 4 | lila | 4080, 4070 |

Tipp (vor Beginn der Erfassung): Ein Handy fertig einrichten, Backup erstellen und auf den anderen Handys bei der Einrichtung „Backup wiederherstellen“ wählen. Danach dort nur noch **Kürzel und Name ändern** (Einstellungen → Kürzel / Name ändern), sonst haben zwei Handys dasselbe Kürzel!

**Jedes Handy braucht ein anderes Kürzel.** Am besten legt man vorher eine kleine Liste an, z. B.:

| Kürzel | Person | Bereich |
|---|---|---|
| A | … | Flur oben 4180–4240 |
| B | … | Flur oben 4260–4300 |
| C | … | Unten 4070–4170 |

## 4. Ablauf der Inventur

1. Im Reiter **Räume** einen Raum antippen und **„+ Gerät in diesem Raum erfassen“** wählen.
2. Gerätetyp, Hersteller, Modell und Seriennummer eingeben. Das Team wird aus dem Raum vorgeschlagen.
3. **„Speichern & nächstes Gerät im selben Raum“**: Die App zeigt die neue ID groß an. Diese ID auf das Gerät schreiben oder kleben. Danach ist das Formular für das nächste Gerät vorbereitet, Raum, Arbeitsplatz, Person und Team sind schon eingetragen.
4. Wenn ein Raum komplett ist: Raum antippen und **„Raum als fertig erfasst markieren“**. Im Grundriss erscheint dann ein grüner Haken.
5. Taucht eine Seriennummer schon auf, warnt die App sofort. So fallen doppelt erfasste Geräte auf.

Gelöschte IDs werden nie erneut vergeben.

## 5. Export nach Excel

Reiter **Excel** → **„Als Excel exportieren“** → **Teilen** (Mail, Teams, AirDrop, OneDrive …) oder **Herunterladen**.

Die Datei `Inventur_<Kürzel>_<Datum>.xlsx` enthält:
- **Inventar:** alle Geräte, mit Filtern und fixierter Kopfzeile; Datumswerte sind echte Excel-Daten
- **Übersicht:** Anzahl nach Gerätetyp, Raum, Team und Status sowie der Fortschritt je Raum. Das sind Formeln, die sich aktualisieren, wenn man in Excel etwas ändert.
- **Räume:** Team-Zuordnung, Gerätezahl, „Fertig erfasst“
- **Prüfen** (nur wenn nötig): Seriennummern, die unter mehreren IDs vorkommen

Die App erinnert daran, wenn länger als einen Tag nicht exportiert wurde.

## 6. Zusammenführen mehrerer Handys

1. Jedes Handy exportiert seine Excel-Datei und schickt sie an eine Person, z. B. per Mail.
2. Diese Person öffnet auf **einem** Handy oder **am PC im Browser** (gleicher Link) den Reiter **Excel** → **„Dateien auswählen …“** und wählt alle Dateien auf einmal aus.
3. Eine Vorschau zeigt: *X neu, Y aktualisiert, Z unverändert*. Außerdem warnt sie vor doppelten Seriennummern.
4. **„Zusammenführen“** antippen und danach **„Als Excel exportieren“**. Das ergibt `Inventur_Gesamt_<Datum>.xlsx`.

Regeln: Gibt es dieselbe ID mehrmals, gewinnt die neuere Änderung. Importierte Geräte behalten ihre ID. Man kann beliebig oft importieren, dabei entstehen keine Duplikate. Fertig-Markierungen der Räume werden mit übernommen.

## 7. Datensicherung

Unter **Einstellungen → Datensicherung**:
- **Backup erstellen:** eine JSON-Datei mit allen Geräten und Einstellungen
- **Backup wiederherstellen:** zum Beispiel nach einem Handywechsel. Das geht auch direkt auf dem Startbildschirm der Einrichtung.
- **Alle Daten löschen:** mit doppelter Bestätigung. Der ID-Zähler bleibt erhalten.

## 8. Anpassen

- Räume, Teams (mit Farbe) und Gerätetypen lassen sich in den **Einstellungen** ändern.
- Die Standardwerte stehen oben in `app.js` (`DEFAULT_ROOMS`, `DEFAULT_TEAMS`, `DEFAULT_TYPES`). Sie gelten nur für Handys, die neu eingerichtet werden.
- **Nach Änderungen an den Dateien** in `service-worker.js` die `VERSION` erhöhen (z. B. `1.0.1`) und neu hochladen. Die Handys holen sich die neue Version beim nächsten Öffnen mit Internet und zeigen einen Hinweis.

## Dateien

| Datei | Zweck |
|---|---|
| `index.html` | Oberfläche |
| `app.js` | gesamte Logik (Erfassung, IDs, Export, Import) |
| `styles.css` | Gestaltung |
| `service-worker.js` | Offline-Cache |
| `manifest.json`, `icons/` | Installation auf dem Home-Bildschirm |
| `vendor/exceljs.min.js` | Bibliothek für Excel-Dateien (ExcelJS 4.4.0, MIT-Lizenz) |
