# Export-MultiGroupTreeHtmlWithDuplicates

## Beschreibung

Dieses PowerShell-Skript analysiert rekursiv die Mitgliedschaften mehrerer Active Directory Gruppen.  
Es erzeugt einen übersichtlichen HTML-Report, der

- die verschachtelten Gruppenstrukturen,
- doppelte Nutzermitgliedschaften (also Benutzer, die in mehreren Gruppen vorkommen),
- sowie Zykluserkennungen in der Gruppenverschachtelung

grafisch darstellt und hervorhebt.

## Voraussetzungen

- PowerShell mit installiertem ActiveDirectory-Modul  
- Angemessene Berechtigungen, um AD-Gruppen und deren Mitglieder auszulesen

## Verwendung

```powershell
Export-MultiGroupTreeHtmlWithDuplicates -StartGroupNames "GruppeA", "GruppeB", "GruppeC" -HtmlOutputPath "C:\Pfad\Zu\Report.html"
```

## Parameter

| Parameter         | Beschreibung                                               | Typ        | Pflicht  | Standardwert                      |
|-------------------|------------------------------------------------------------|------------|----------|----------------------------------|
| `StartGroupNames` | Array von Active Directory Gruppennamen als Ausgangspunkt  | `string[]` | Ja       | -                                |
| `HtmlOutputPath`  | Speicherort und Dateiname des generierten HTML-Reports    | `string`   | Nein     | `. \MultiGroupTreeWithDuplicates.html` |


## Funktionsweise

- Für jede Startgruppe wird die Gruppenmitgliedschaft rekursiv ausgelesen.  
- Gruppen werden als Baumstruktur dargestellt.  
- Zyklen in der Gruppenverschachtelung werden erkannt und farblich markiert.  
- Benutzer, die mehrfach in verschiedenen Gruppen auftauchen, werden im Bericht farblich hervorgehoben und am Ende des Reports gelistet.  


