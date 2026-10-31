# PoC-Ergebnis: OneDrive/Excel-Anbindung für Pocket Budget

Stand: 2026-10-31. Getestet mit privatem Microsoft-Konto und Dummy-Datei `PocketBudget-PoC/test.xlsx`.

## Bewertung

Die Architektur trägt für **private Microsoft-Konten**: Web-App auf Vercel (Functions), Anmeldung per OAuth 2.0 Authorization Code Flow (vertraulicher Client), Dateizugriff über Microsoft Graph per Download/Upload der `.xlsx`.

Die Excel-REST-API (Tabellen, Zellen, Sessions) ist laut Microsoft für private OneDrive-Konten **nicht** unterstützt. Der Weg über Download/Upload ist deshalb die gewählte Lösung. Die Datei wird im Backend mit ExcelJS bearbeitet und komplett neu hochgeladen. Das ist für eine Datei mit einfachen Spalten (ID, Ausgabe, Ort, Betrag) ohne Diagramme geeignet. Bei komplexeren Dateien (Diagramme, Pivot, Makros) wäre es riskant.

## Abnahmekriterien

| Nr. | Kriterium | Ergebnis | Nachweis |
|---|---|---|---|
| 1 | Microsoft-Anmeldung | erfüllt | manuell getestet |
| 2 | Richtige Datei auswählen | erfüllt | manuell getestet (eigener Datei-Browser über das Backend) |
| 3 | Datei auslesen | erfüllt | manuell getestet |
| 4 | Neuer Datensatz in Originaldatei | erfüllt | manuell getestet |
| 5 | Änderung in Excel sichtbar | erfüllt (Rückmeldung der Testerin: PoC erfolgreich) | manuell |
| 6 | Änderung aus Excel in der App sichtbar | erfüllt (wie 5) | manuell |
| 7 | Parallele Änderungen nicht unbemerkt überschrieben | **teilweise belegt** | Automatische Tests (9, bestanden) mit simuliertem OneDrive. Bei geöffneter Datei lehnte OneDrive den Upload ab (Fehler „Datei schließen"), nach dem Schließen klappte das Speichern. Die Meldung „KONFLIKT" bei paralleler Änderung derselben Zeile wurde an echtem OneDrive **nicht** bestätigt. |
| 8 | Vollständige Trennung der Benutzerdaten | per Code-Prüfung belegt, siehe Einschränkung unten | Benutzer nur aus Server-Session; `driveId`/`itemId` nur aus dem Server-Store; Datei-Auswahl wird mit dem Token des Benutzers geprüft; Smoketest: alle Datenzugriffe ohne Login → 401. Kein Test mit zweitem echtem Benutzer. |
| 9 | Berechtigungen dokumentiert | erfüllt | siehe unten |

## Berechtigungen (Microsoft Graph, delegiert)

`openid`, `profile`, `offline_access`, `User.Read`, `Files.ReadWrite`.

- Kein `.All`, keine Application-Berechtigungen.
- `Files.ReadWrite` erlaubt der App Lese- und Schreibzugriff auf das gesamte OneDrive des angemeldeten Benutzers. Eine Beschränkung auf eine einzelne Datei (`Files.SelectedOperations.Selected`) ist laut Dokumentation nicht belegt machbar. Die Begrenzung auf die gewählte Datei erfolgt nur durch die App-Logik.

## Wichtige Einschränkungen

1. **Betreiberzugriff:** Wer das Backend betreibt (Datenbank plus `TOKEN_ENC_KEY`), kann technisch die Refresh Tokens entschlüsseln und damit auf die verbundenen OneDrives zugreifen. Die Trennung der Benutzer gilt innerhalb der App, nicht gegenüber dem Betreiber. Die Anforderung „Entwicklerin darf keinen Zugriff auf Daniels Finanzdaten haben" ist mit **einer gemeinsamen** Instanz nicht technisch garantierbar. Eine technische Trennung gäbe es nur, wenn Daniel eine eigene Instanz betreibt (eigenes Vercel-Projekt, eigene App-Registrierung, eigener Schlüssel) oder wenn alles im Browser läuft (keine Tokens im Backend, dafür tägliche Neuanmeldung).

   **Entscheidung (2026-10-31): Variante A.** Gemeinsame Instanz, betrieben von der Entwicklerin (eigener Code, Git, Vercel, App-Registrierung). Daniel verknüpft sein eigenes OneDrive über die Website. Die Entwicklerin sieht seine Daten nicht über die App, könnte sie als Betreiberin aber technisch erreichen. Die Anforderung ist damit organisatorisch (Vertrauen), nicht technisch erfüllt. Maßnahmen:
   - Daniel wird vorab darüber informiert, dass die Betreiberin technischen Zugriff hätte.
   - Es werden keine Dateiinhalte gespeichert oder protokolliert; gespeichert werden nur das verschlüsselte Refresh Token und die Datei-ID.
   - Zwei-Faktor-Anmeldung bei GitHub, Vercel, Neon und Microsoft.
   - Daniel kann den Zugriff jederzeit unter https://account.live.com/consent/Manage widerrufen.
2. **Daniels Konto:** Daniel nutzt ein privates Microsoft-Konto. Die Zustimmung durch Firmen-Admins entfällt damit. Mit seinem Konto wurde noch nicht getestet. Der erste Login mit dem Zustimmungsdialog ist ausstehend.
3. **Sperre bei geöffneter Datei:** Ist die Datei in Excel geöffnet, schlägt das Speichern fehl (Fehler 423, Meldung an den Benutzer). Die App muss das verständlich anzeigen.
4. **Konflikte auf Dateiebene:** Jede Änderung an der Datei macht einen Upload mit altem eTag ungültig. Die App wiederholt dann auf frischen Daten (Anhängen) oder meldet einen Konflikt (Ändern derselben Zeile).
5. **Refresh Tokens:** Laufen bei Nichtbenutzung nach etwa 90 Tagen ab und können widerrufen werden. Die App verlangt dann eine erneute Anmeldung.
6. **Client-Secret** der App-Registrierung läuft nach 6 Monaten ab und muss erneuert werden.
7. **Nicht getestet:** Vercel-Deployment, Neon-Postgres als Token-Speicher, Microsoft-Dateiauswahldialog (Picker), Excel Desktop als Co-Autor, Handy-Browser.

## Offene Punkte vor dem Produktivbetrieb

- Daniel über den Betreiberzugriff (Einschränkung 1) informieren und Maßnahmen umsetzen.
- Test mit Daniels Konto und Konflikttest an echtem OneDrive (zwei Browser/Excel parallel).
- Vercel und Neon einrichten und testen.
- Entscheidung, ob eine ID-Spalte in Daniels Datei ergänzt wird (nötig für sicheres Ändern und Löschen).
