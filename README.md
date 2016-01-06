# Extract-AD-users-from-a-distribution-Group
Script which extract users belong to a distribution group in ActiveDirectory


I'm a fanboy Linux. Never coded on windows. 
But I'm still "openminded" so I answer "yes" to a company for coding that. An so discovering PowerShell  
PowerShell is very easy way and as well limited 

So what's do my program : 
It simply extract users from a distribution group, and show the differences (compare) since the last extract. 

Prerequisite : 
- Active directory 
- Computer within a domain active directory 
 

Manual : 
- It's quite simple (I like simplicity) you just need to set the var "$Currentdir" where is currently located the script.
-  Set in a raw format the distribution group impacted inside the file : config.ini
-  Execute the script : ADExtractUsers_DistributionGroup.ps1

Result : 
- Create a dir : extract_csv  
- Put inside the full extract csv  e.g : FULL_VTPD_20150409.csv		contain : 'SamAccountName,Groupe,Enabled'   
- Put inside the incremental extract csv since the last extract e.g : INC_VTPD_20150409.csv   contain : 'SamAccountName,Groupe,Enabled'   
- Of course you have a log file : log.txt 

Comment : 
- This script is clever, all cases have been thought. 
- If you got powershell >2 (Which I guess not) you don't need to use : function-export.ps1. (But anyway it will be ignore)
    - > Why : Load the function "-append" from command  'export-csv' is  available from powershell v3.0
- In the output file you can set alias for renaming the group  e.g VTOMPROD_Consultation -> Consult

Enjoy. 



