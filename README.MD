# Export-MultiGroupTreeHtmlWithDuplicates

## Description

This PowerShell-Script recursively analyzes the group memberships of multiple Active-Directory groups. 
It generates an easy to read HTML-Report which shows graphically an highlighted

- the nested group structure,
- duplicate user memberships (users who are members of multiple nested groups),
- duplicate nested groups within these groups

## Prerequisites

- PowerShell with installed ActiveDirectory-Module   
- Appropriate Permission to browse AD-Groups and members

## Usage

```powershell
Export-MultiGroupTreeHtmlWithDuplicates -StartGroupNames "GroupA", "GroupB", "GroupC" -HtmlOutputPath "C:\path\to\Report.html"
```

## Parameter

| Parameter         | Description                                               | Type        | Mandatory  | Default                      |
|-------------------|------------------------------------------------------------|------------|----------|----------------------------------|
| `StartGroupNames` | Array of Active Directory group names as starting point  | `string[]` | Yes       | -                                |
| `HtmlOutputPath`  | Where to save the HTML Report and file name of the HTML-Report    | `string`   | No     | `. \MultiGroupTreeWithDuplicates.html` |


## Function

- Members of each start group are browsed recursively.  
- Groups and members are shown as Tree-Structure.  
- Duplicate nested groups or members are highlighted in color.  


