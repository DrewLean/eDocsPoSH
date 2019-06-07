# eDocsPoSH
Scripts used for managing eDocs DM/RM

importExample.ps1: 
Code changes required by line
6 to your Libraries
13 to API module location
14 to CSV location
22/23 to username and password of eDOCS account to use for bulk imports
43-135 if your extensions and appIDs are named differently
162 to your profile form name.

This will import based on a csv document. This is highly customised to our environment and is meant as an example only.

eDOCSAPI.psm1
Code changed required by line
195 - Logon param for server sets
266 & 269 - SQL Details
