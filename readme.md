#Export User and Roles from the Sitecore Domain Manager
Allow administrators to export users and roles to Excel directly from the Domain Manager. 

Created for simplying user rights validation. 
A matching import functionality would be the obvious next move, but is not planned at the moment.

1508 / Design in Love with Technology / http://1508.dk

## Required dependencies
Sitecore.Client
Sitecore.Kernel
SpreadsheetGear2012.Core	(Licens required)

## Installation 
There is no nuget package, instead 2 screenshots display the required changes. 
Please note that the code is based on SpreadsheetGear2012.Core, a commercial product (for the best handling of special characters). 
It can be recoded to another excel generator engine or straight csv. 

## Possible improvements 
* A matching import functionality.

Feel free to contribute!