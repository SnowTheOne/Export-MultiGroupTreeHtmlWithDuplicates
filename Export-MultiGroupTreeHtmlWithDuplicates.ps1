<#
.PSScriptInfo

.VERSION
2.0.2

.TAGS
ActiveDirectory, AD, Groups, GroupMembership, Recursive, DuplicateMembers, HTMLReport

.DESCRIPTION
Generates  an easy to read HTML-Report which shows graphically an highlighted the nested group structure, duplicate user memberships (users who are members of multiple nested groups) and duplicate nested groups within these groups.

.EXTERNALMODULEDEPENDENCIES
ActiveDirectory

.REQUIREDSCRIPTS
None

.EXTERNALSCRIPTDEPENDENCIES
None

.PARAMETER StartGroupNames
Array of Active Directory group names as starting point.

.PARAMETER HtmlOutputPat
Where to save the HTML Report and file name of the HTML-Report

.EXAMPLE
Export-MultiGroupTreeHtmlWithDuplicates -StartGroupNames "GroupA", "GroupB", "GroupC")
Generates a HTML-Report for the Groups "GroupA", "GroupB" and "GroupC".

.NOTES
Needs the ActiveDirectory PowerShell-Module and appropriate permissions to browse group memberships.
#>

function Export-MultiGroupTreeHtmlWithDuplicates {
    param (
        [string[]]$StartGroupNames,
        [string]$HtmlOutputPath = ".\MultiGroupTreeWithDuplicates.html"
    )

    Import-Module ActiveDirectory -ErrorAction Stop

    $htmlParts = New-Object System.Collections.Generic.List[string]
    $htmlParts.Add("<html><head><meta charset='utf-8'><style>body {font-family: Calibri;} ul {margin-left: 1em;} .cyc {color: red;} .dup {color: orange; font-weight: bold;} </style></head><body>")
    $htmlParts.Add("<h1>AD group structure with duplicate users per Tree</h1>")

    foreach ($StartGroupName in $StartGroupNames) {
        $visitedGroups = @{}
        $globalUserSet = @{}
        $duplicateUsers = @{}

        $htmlParts.Add("<h2>AD group structure: $StartGroupName</h2>")
        $htmlParts.Add("<ul>")

        function Add-ToHtml {
            param([string]$groupDN)

            if ($visitedGroups.ContainsKey($groupDN)) {
                $groupName = $visitedGroups[$groupDN]
                $htmlParts.Add("<li><span class='cyc'>$groupName (duplicate detected)</span></li>")
                return
            }

            try {
                $group = Get-ADGroup -Identity $groupDN -Properties Name -ErrorAction Stop
            } catch {
                $htmlParts.Add("<li><i>Group $groupDN not found</i></li>")
                return
            }

            $groupName = $group.Name
            $visitedGroups[$groupDN] = $groupName

            $htmlParts.Add("<li>$groupName<ul>")

            $members = Get-ADGroupMember -Identity $groupDN -ErrorAction SilentlyContinue
            if (-not $members) {
                $htmlParts.Add("<li><i>No members</i></li>")
            } else {
                foreach ($m in $members) {
                    if ($m.objectClass -eq 'group') {
                        Add-ToHtml -groupDN $m.DistinguishedName
                    } else {
                        $userKey = $m.SamAccountName
                        $displayName = "$($m.Name) ($userKey)"

                        if ($globalUserSet.ContainsKey($userKey)) {
                            $duplicateUsers[$userKey] = $displayName
                            $htmlParts.Add("<li class='dup'>$displayName (duplicate found)</li>")
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
            $htmlParts.Add("<li><i>Group $StartGroupName not found</i></li>")
        }

        $htmlParts.Add("</ul>")

        if ($duplicateUsers.Count -gt 0) {
            $htmlParts.Add("<h3>Duplicate users found in group $StartGroupName</h3><ul>")
            foreach ($dupUser in $duplicateUsers.GetEnumerator() | Sort-Object Value) {
                $htmlParts.Add("<li>$($dupUser.Value)</li>")
            }
            $htmlParts.Add("</ul>")
        } else {
            $htmlParts.Add("<p>No duplicate users found in group $StartGroupName.</p>")
        }
    }

    $htmlParts.Add("</body></html>")

    $htmlOutput = $htmlParts -join "`r`n"
    Set-Content -Path $HtmlOutputPath -Value $htmlOutput -Encoding UTF8

    Write-Host "✅ HTML-Report generated: $HtmlOutputPath"
    Start-Process $HtmlOutputPath
}
