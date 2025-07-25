<#
.PSScriptInfo

.VERSION
2.0.2

.TAGS
ActiveDirectory, AD, Groups, GroupMembership, Recursive, DuplicateMembers, HTMLReport

.DESCRIPTION
Erstellt einen HTML-Report mit den Mitgliedschaften mehrerer AD-Gruppen, zeigt verschachtelte Gruppen, doppelte Nutzer (pro Baum) und erkennt Zyklen in der Gruppenstruktur.

.EXTERNALMODULEDEPENDENCIES
ActiveDirectory

.REQUIREDSCRIPTS
Keine

.EXTERNALSCRIPTDEPENDENCIES
Keine

.RELEASENOTES
Version 2.0.2
- SamAccountName wird hinter jedem Benutzernamen angezeigt
- Duplikaterkennung je Gruppenbaum bleibt bestehen
#>

<#
.DESCRIPTION
Ermittelt rekursiv die Mitglieder mehrerer Active Directory-Gruppen und erstellt einen HTML-Bericht,
der doppelte Nutzermitgliedschaften (pro Gruppenbaum) und verschachtelte Gruppenstrukturen inklusive Zyklen hervorhebt.

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
    $htmlParts.Add("<h1>AD Gruppenstrukturen mit doppelten Nutzern (pro Baum)</h1>")

    foreach ($StartGroupName in $StartGroupNames) {
        $visitedGroups = @{}
        $globalUserSet = @{}
        $duplicateUsers = @{}

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
                        $userKey = $m.SamAccountName
                        $displayName = "$($m.Name) ($userKey)"

                        if ($globalUserSet.ContainsKey($userKey)) {
                            $duplicateUsers[$userKey] = $displayName
                            $htmlParts.Add("<li class='dup'>$displayName (bereits gezeigt)</li>")
                        } else {
                            $globalUserSet[$userKey] = $true
                            $htmlParts.Add("<li>$displayName</li>")
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

        # Optional: Doppelte Nutzer pro Baum auflisten
        if ($duplicateUsers.Count -gt 0) {
            $htmlParts.Add("<h3>Doppelte Nutzer in $StartGroupName</h3><ul>")
            foreach ($dupUser in $duplicateUsers.GetEnumerator() | Sort-Object Value) {
                $htmlParts.Add("<li>$($dupUser.Value)</li>")
            }
            $htmlParts.Add("</ul>")
        } else {
            $htmlParts.Add("<p>Keine doppelten Nutzer in $StartGroupName gefunden.</p>")
        }
    }

    $htmlParts.Add("</body></html>")

    $htmlOutput = $htmlParts -join "`r`n"
    Set-Content -Path $HtmlOutputPath -Value $htmlOutput -Encoding UTF8

    Write-Host "✅ HTML-Datei mit SamAccountName-Hinweis erzeugt: $HtmlOutputPath"
    Start-Process $HtmlOutputPath
}
