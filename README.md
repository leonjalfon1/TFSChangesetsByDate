# TFSChangesetsByDate
Powershell script to retrieve changesets by dates and generate a simple summary report file

---

## Example Uses

 - Find all the changesets created between two dates
 - Find total of changesets between two dates
 - Find total of changesets by month in the last year
 - Find who did changes in specific path between two dates

---
 
## Requisites

 - Powershell version 2.0 or later
 - TF.exe tool (usually located in C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE)
 - More info about TF in: https://www.visualstudio.com/en-us/docs/tfvc/use-team-foundation-version-control-commands
 
---
 
## Parameters

### Mandatory

#### Path
 - Path in source control (directory or file)
 - Example: "$/TeamProject/Branch/Directory"
#### CollectionUrl
 - TFS Collection Url
 - Example: "http://tfsserver:8080/tfs/PTU_Collection"
#### Dates
 - With the following format {StartDay/StartMonth/StartYear-EndDay/EndMonth/EndYear}
 - You can set several dates separated by ';'
 - Example: "1/2/2016-1/4/2016;1/4/2016-1/6/2016;1/6/2016-1/8/2016"
 
### Optional

#### TFPath
 - Path where TF.exe is stored
 - By default TF.exe the tool searched in the path "C:\Program Files (x86)\{VS 17,15,13}\Common7\IDE"
 - For Example: "C:\Program Files\TFVC"
#### File
 - Specify the file where the summary will be created (the default is "Desktop\ChangesetsByDates_{datetime}")
 - Example: "C:\Users\User\Desktop\TFSChangesetsByDates\Summary.txt"

---

## Usage

- Mandatory
```
.\TFSChangesetsByDates.ps1 -Path "$/TP/Branch/Directory" -CollectionUrl "http://tfsserver:8080/tfs/Collection" -Dates "1/2/2016-1/4/2016;1/4/2016-1/6/2016;1/6/2016-1/8/2016"
```
- All
```
.\TFSChangesetsByDates.ps1 -Path "$/TP/Branch/Directory" -CollectionUrl "http://tfsserver:8080/tfs/Collection" -Dates "1/2/2016-1/4/2016;1/4/2016-1/6/2016;1/6/2016-1/8/2016" -TFPath "C:\Program Files\TFVC" -File "C:\Users\User\Desktop\TFSChangesetsByDates\Summary.txt"
```
 
---
 
## Summary Report

 - Show the total changesets between each interval
 - Show the total changesets in all the intervals
 - Show the changesets details in each interval
 - See an example in the [SummaryReportExample](SummaryReportExample.txt)
 
---

## Contributing

 - Please feel free to contribute, suggest ideas or open issues
