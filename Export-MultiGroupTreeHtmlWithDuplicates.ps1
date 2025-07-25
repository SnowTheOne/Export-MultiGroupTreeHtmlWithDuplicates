<#
.PSScriptInfo

.VERSION
2.0.0

.TAGS
ActiveDirectory, AD, Groups, GroupMembership, Recursive, DuplicateMembers, HTMLReport

.DESCRIPTION
Erstellt einen HTML-Report mit den Mitgliedschaften mehrerer AD-Gruppen, zeigt verschachtelte Gruppen, doppelte Nutzer und erkennt Zyklen in der Gruppenstruktur.

.EXTERNALMODULEDEPENDENCIES
ActiveDirectory

.REQUIREDSCRIPTS
Keine

.EXTERNALSCRIPTDEPENDENCIES
Keine

.RELEASENOTES
Version 2.0.0
- Unterstützt mehrere Start-Gruppen
- Erkennt und markiert doppelte Nutzermitgliedschaften
- Erzeugt übersichtlichen HTML-Report mit Visualisierung der AD-Gruppenstruktur
#>

<#
.DESCRIPTION
Ermittelt rekursiv die Mitglieder mehrerer Active Directory-Gruppen und erstellt einen HTML-Bericht,
der doppelte Nutzermitgliedschaften und verschachtelte Gruppenstrukturen inklusive Zyklen hervorhebt.

.PARAMETER StartGroupNames
Ein Array von AD-Gruppennamen als Strings, die als Ausgangspunkt für die Analyse dienen.

.EXAMPLE
Export-MultiGroupTreeHtmlWithDuplicates -StartGroupNames @("GruppeA", "GruppeB", "GruppeC")
Erzeugt einen HTML-Report für die Gruppen "GruppeA", "GruppeB" und "GruppeC".

.NOTES
Benötigt das ActiveDirectory PowerShell-Modul und angemessene Rechte, um Gruppenmitgliedschaften auszulesen.
#>

function Export-MultiGroupTreeHtmlWithDuplicates {
    param (
        [string[]]$StartGroupNames,
        [string]$HtmlOutputPath = ".\MultiGroupTreeWithDuplicates.html"
    )

    Import-Module ActiveDirectory -ErrorAction Stop

    $htmlParts = New-Object System.Collections.Generic.List[string]
    $htmlParts.Add("<html><head><meta charset='utf-8'><style>body {font-family: Calibri;} ul {margin-left: 1em;} .cyc {color: red;} .dup {color: orange; font-weight: bold;} </style></head><body>")
    $htmlParts.Add("<h1>AD Gruppenstrukturen mit doppelten Nutzern</h1>")

    # HashSets zur Duplikaterkennung
    $globalUserSet = @{}
    $duplicateUsers = @{}

    foreach ($StartGroupName in $StartGroupNames) {
        $visitedGroups = @{}

        $htmlParts.Add("<h2>AD Gruppenstruktur: $StartGroupName</h2>")
        $htmlParts.Add("<ul>")

        function Add-ToHtml {
            param([string]$groupDN)

            if ($visitedGroups.ContainsKey($groupDN)) {
                $groupName = $visitedGroups[$groupDN]
                $htmlParts.Add("<li><span class='cyc'>$groupName (Zyklus erkannt)</span></li>")
                return
            }

            try {
                $group = Get-ADGroup -Identity $groupDN -Properties Name -ErrorAction Stop
            } catch {
                $htmlParts.Add("<li><i>Gruppe $groupDN nicht gefunden</i></li>")
                return
            }

            $groupName = $group.Name
            $visitedGroups[$groupDN] = $groupName

            $htmlParts.Add("<li>$groupName<ul>")

            $members = Get-ADGroupMember -Identity $groupDN -ErrorAction SilentlyContinue
            if (-not $members) {
                $htmlParts.Add("<li><i>Keine Mitglieder</i></li>")
            } else {
                foreach ($m in $members) {
                    if ($m.objectClass -eq 'group') {
                        Add-ToHtml -groupDN $m.DistinguishedName
                    } else {
                        # Duplikate prüfen
                        $userKey = $m.SamAccountName
                        if ($globalUserSet.ContainsKey($userKey)) {
                            $duplicateUsers[$userKey] = $m.Name
                            $htmlParts.Add("<li class='dup'>$($m.Name) (bereits gezeigt)</li>")
                        } else {
                            $globalUserSet[$userKey] = $true
                            $htmlParts.Add("<li>$($m.Name)</li>")
                        }
                    }
                }
            }

            $htmlParts.Add("</ul></li>")
        }

        try {
            $startGroup = Get-ADGroup -Identity $StartGroupName -Properties DistinguishedName
            Add-ToHtml -groupDN $startGroup.DistinguishedName
        } catch {
            $htmlParts.Add("<li><i>Startgruppe $StartGroupName nicht gefunden</i></li>")
        }

        $htmlParts.Add("</ul>")
    }

    # Liste der doppelten Nutzer
    if ($duplicateUsers.Count -gt 0) {
        $htmlParts.Add("<h2>Doppelte Nutzer in der Struktur</h2>")
        $htmlParts.Add("<ul>")
        foreach ($dupUser in $duplicateUsers.GetEnumerator() | Sort-Object Name) {
            $htmlParts.Add("<li>$($dupUser.Value) (SamAccountName: $($dupUser.Key))</li>")
        }
        $htmlParts.Add("</ul>")
    } else {
        $htmlParts.Add("<p>Keine doppelten Nutzer gefunden.</p>")
    }

    $htmlParts.Add("</body></html>")

    $htmlOutput = $htmlParts.ToArray() -join "`r`n"
    Set-Content -Path $HtmlOutputPath -Value $htmlOutput -Encoding UTF8

    Write-Host "✅ HTML-Datei mit doppelten Nutzern erzeugt: $HtmlOutputPath"
    Start-Process $HtmlOutputPath
}
