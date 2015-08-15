# PowerShell_SQLite_Report
Create a report with data from a CSV file and Web Service

## The Problem
One product has data about servers, but limited reporting functionality. 

Another product has different data about servers, but custom reporting is done by the reporting group. A web service is available for querying and searching the web service however... 

## The Goal
Combine data from two different sources (a CSV and Web Service) into a report that shows what servers are in both systems and which servers are in one inventory system but not the other. 

## The Solution
Obtaining the data from both sources in PowerShell was so easy, why wouldn't comparing the data and matching up records be equally easy? 

I was wrong, obtaining and transforming the data in PowerShell was easy, but matching up records was a slow painstaking process in PowerShell. 

Many times I found myself thinking "this would be so easy if it was a database! I could just do a SQL JOIN..." 


So I looked into putting data into a database. MS-SQL and MS Access all took some setup, and then I found PSSQLite. 
(https://github.com/RamblingCookieMonster/PSSQLite)

The PSSQLite module was perfect. I could quickly and easily create a database, and even though SQLite doesn't implement every join type that is easy to work around, espically when you can combine PowerShell result objects with a plus sign. ($combinedResult = $SQLQueryResult1 + $SQLQueryResult2)

Watching the first run of the PSSQLite version of the script was just magic. Or something. 

Things have been sanitized so this version of the script si mainly an example. 