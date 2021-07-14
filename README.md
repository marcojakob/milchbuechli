# Milchbüechli-Supertool

Das Milchbüechli-Supertool ist eine einfache Einnahmen- und Ausgabenrechnung für die Buchhaltung einer Einzelunternehmung in der Schweiz.

Dies ist die einfachste Form der Buchhaltung für Einzelunternehmen mit weniger als CHF 100'000 Umsatz pro Jahr (erlaubt wäre bis CHF 500'000, wird aber komplizierter wegen der Mehrwertsteuer). Diese Form der Buchhaltung erleichtert den Einstieg in eine selbständige oder unternehmerische Tätigkeit.

Wenn die Belege sauber abgelegt und benannt sind, so ist das Erstellen der **Buchhaltung am Jahresende nur noch ein Knopfdruck**!

Das Tool funktioniert in Excel und auch in Google Sheets.

## Belege 

Das Milchbüechli-Supertool hilft, eine klare Struktur in die Ablage von Belegen zu bringen. In den Dateinamen der Belege sind alle Informationen für die Buchhaltung bereits enthalten (Datum, Beschreibung, Betrag, etc.).

### Einnahmen

Im Unterordner `Einnahmen` werden alle Belege (Rechnungen) für Einnahmen abgelegt.

Beispieldateiname: `2021-02-23 KUNDE A - Workshop CHF1100.pdf`

1. Das Rechnungsdatum (in umgekehrter Reihenfolge Jahr-Monat-Tag wegen der Sortierung).
2. Ein freier Text zur Beschreibung.
3. Der Betrag.

Sobald wir den Betrag erhalten haben, können wir optional das Datum der Überweisung notieren. Dann sieht der ganze Dateiname so aus: `2021-02-23 KUNDE A - Workshop CHF1100 BEZ2021-03-05.pdf`

### Ausgaben

Im Unterordner `Ausgaben` werden alle Belege für Ausgaben abgelegt.

Beispieldateiname: `2021-04-04 Arbeitsplatz Effinger CHF270 KAT5.pdf`

1. Das Rechnungsdatum (in umgekehrte Reihenfolge Jahr-Monat-Tag wegen der Sortierung).
2. Ein freier Text zur Beschreibung.
3. Der Betrag.
4. Die Ausgabekategorie (siehe Excel-Datei im Tabellenblatt Einstellungen).

## Anleitung (Excel)

Die Excel-Datei hat zwei hinterlegte Makro-Funktionen. Eine ist für die Einnahmen und eine für die Ausgaben. Im Tabellenblatt Einnahmen bzw. Ausgaben hat es entsprechende Knöpfe, um die Daten aus den Dateinamen auszulesen.

Damit dies funktioniert, braucht es einen Ordner für die Einnamen und einen für die Ausgaben. Dort drin müssen die Belege und Rechnungen mit ganz bestimmten Dateinamen abgespeichert sein (siehe Beispielordner oder Einnahmen/Ausgaben unten).

## Anleitung (Google Sheets)

Das Milchbüechli-Supertool kann auch in Google Sheets verwendet werden. So wird das Google Sheet eingerichtet:

1. Excel-Datei auf Google Drive hochladen.
2. Die Unterordner `Einnahmen` und `Ausgaben` erstellen.
3. Die Excel-Datei mit Google Sheets öffnen. Zuerst wird sie im Excel-Format geöffnet.
4. Nach Google Sheets umwandeln: Im Menu *Datei* den Punkt *Als Google-Tabelle speichern* auswählen.
5. Nun gibt es zwei Dateien. Die Excel-Datei (Endung .xlsm) kann gelöscht werden.
6. In der Google Sheets Datei Menu *Tools*, *Skripteditor* öffnen.
7. Im Skripteditor den ganzen Inhalt der Datei [`GoogleSheetsScript.gs`](GoogleSheetsScript.gs) hineinkopieren.
8. Speichern, Skripteditor schliessen und Google Sheet neu laden.

Nun sollte in Google Sheets ein neues Menu `🥛 Milchbüechli` erscheinen. Unter diesem Menu können die Einnahmen und Ausgaben eingelesen werden. Allenfalls muss bei der ersten Ausführung noch die Berechtigung erteilt werden. 

## Lizenz

Erstellt mit ✨ von www.jakobservices.ch

Das Milchbüechli-Supertool darf frei verwendet und weiterentwickelt werden. Dieses Werk ist - sofern nicht anders angegeben - lizenziert unter [Creative Commons Attribution-ShareAlike 4.0 International](https://creativecommons.org/licenses/by-sa/4.0/) (CC BY-SA 4.0).

### Was bedeutet die Lizenz konkret?

Diese Lizenz drückt zwei Aspekte aus, die uns wichtig sind:

- **Attribution**: Wertschätzung und Dank ausdrücken für die Leute, von denen wir etwas empfangen haben.
- **ShareAlike**: Das, was wir grosszügig von anderen empfangen haben, geben wir im gleichen Sinn wieder weiter.

Du darfst die Unterlagen also verwenden, kopieren und weiterentwickeln. Dabei musst du einfach die Autoren nennen, die die Unterlagen erstellt haben (Attribution). Und wenn du die Unterlagen weiterentwickelst, dann musst du diese wieder unter den gleichen Bedingungen anderen zur Verfügung stellen (ShareAlike).

### Quellenangabe

So etwa könnte eine Quellenangabe aussehen:

📌 Milchbüechli-Supertool von [Jakob Services](https://www.jakobservices.ch), Lizenz: [CC BY-SA 4.0](https://creativecommons.org/licenses/by-sa/4.0/)

## Spende

🍣 Wenn dir das Milchbüechli-Supertool weitergeholfen hat, darfst du mir gerne mit dem Donate-Knopf ein Mittagessen spendieren :-). Herzlichen Dank!

[![Spende via PayPal](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=info@jakobservices.ch&item_name=Milchbuechli-Supertool&currency_code=CHF)