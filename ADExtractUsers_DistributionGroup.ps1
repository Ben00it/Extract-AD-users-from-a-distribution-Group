# ***********************************************************************************************
# *                                                                                             *
# * Date : 01/04/2014  						                                                    *
# * Author : Ben00it                                                                            *
# *                                                                                             *
# * Script which extract users belong to a distribution group in ActiveDirectory                *
# *                                                                                             *
# *     Input: File config.ini contain distribution group                                       *
# *            e.g  "GG_ROL_VTOMPROD_TestOrdonnancement"  								        *
# *                                                                                             *
# *    Output: Full Extract.CSV contain : 'SamAccountName,Groupe,Enabled'                       *
# *            e.g : FULL_VTPD_20150409.csv														*
# *            Incremental extract show the diff since the last Full Extract in date            * 
# *            e.g : INC_VTPD_20150409.csv                                                      *
# *                                                                                             *
# ***********************************************************************************************

	#---------------------------------------------------------- 
	# Declare Variables
	#----------------------------------------------------------
	# General Vars
	# Function prompt { "$pwd\" }
	# Only this var must be set, define the path for schedule task exemple : "
	$Currentdir = "C:\Exploit\Extract Group\"
 
	# Modules Vars  
	$append = "function-export.ps1"
	$ModuleExportCsvAppend = "$Currentdir$append"
      
	# Init File 
	$Configini = "config.ini"
	$ConfigFile = "$Currentdir$Configini"
 
	# Log File
	$Log = "log.txt"
	$LogFile = "$Currentdir$Log"
   
	# Full Export CSV (Date + File) 
	$Dircsv = "extract_csv\"
	$Date = Get-Date -Format  yyyyMMdd
	$CsvFileTemp = "$Currentdir$dircsv" + "FULL_VTPD_" + $Date + ".tmp" # eg : FULL_VTPD_20150324.tmp
	$CsvFile = "$Currentdir$dircsv" + "FULL_VTPD_" + $Date + ".csv" # eg : FULL_VTPD_20150324.csv
    
	# Incrémental Export CSV
	$Lastcsv = "$Currentdir$dircsv" + "FULL_*"
	$TempReverseInc = "$Currentdir$dircsv" + "INC_VTPD_" + $Date + ".tmp" # eg : FULL_VTPD_20150324.tmp
	$CsvFileInc =  "$Currentdir$dircsv" + "INC_VTPD_" + $Date + ".csv" # eg : INC_VTPD_20150324.txt
   
	#---------------------------------------------------------- 
	# Declare & Load Functions
	#----------------------------------------------------------     

	# Function Logs e.g : log "string" will push data in $LogFile and will screen it on the output
	function Log ($string) {  
		Write-Host $string
    "$(get-date -Format 'yyyy-MM-dd" "HH:mm:ss') -> $($string)" | Out-File $LogFile -Append -Encoding unicode
	}

	# Config.ini check
	function Initfile {
		If (Test-Path $ConfigFile) { # Test if config.ini exists
			if((Get-Content $ConfigFile) -eq $Null) { # Test if config.ini is empty
				echo ""
				log "[INFO] Fichier '$Configini' vide, Merci de renseigner le contenu du fichier '$Configini' (Groupe AD de type CN, exemple: `"GG_ROL_VTOMPROD_TestOrdonnancement`")"
				Exit 1 
			}
		}
		Else { # If config.ini doesn't exist, then create 
			echo ""
			log "[ERREUR] Fichier '$Configini' non présent"
			New-Item $ConfigFile -type file
			log "[INFO] Création du fichier '$Configini': OK"
			log "[INFO] Merci de renseigner le contenu du fichier '$Configini' (Groupe AD de type CN, exemple: `"GG_ROL_VTOMPROD_TestOrdonnancement`")"
			echo "" 
			Exit 1
		}
	}
  
    function Checkfolder {
		# Check if 'Extract_csv' folder exists 
		If (Test-Path $Currentdir$dircsv){
			# Folder already exists, we continu 
		}
		Else{ #folder need to be create
			New-Item -ItemType Directory -path "$Currentdir$dircsv"
			echo ""
			log "[INFO] Création du répertoire : $dircsv"
		}
	}
   
	# Check If a full CSV of the same date already exist, if yes, then delete
	function DeleteCsvFile { If ( Test-Path $CsvFile ){	Remove-Item $CsvFile } }
	
	# Delete CSV Full Temp File
	function DeleteCsvFileTemp { If ( Test-Path $CsvFileTemp ){ Remove-Item $CsvFileTemp } }
	
	# Check If a incremental CSV of the same date already exist, if yes, then delete
	function DeleteCsvFileInc { If ( Test-Path $CsvFileInc ){ Remove-Item $CsvFileInc } }
	
	# Delete CSV Incremental Temp File 
	function DeleteTempReverseInc {	If ( Test-Path $TempReverseInc ){	Remove-Item $TempReverseInc} }
   
	# Function Alias group e.g : 'GG_ROL_VTOMPROD_TestOrdonnancement' -> 'Test'
	function Aliasgroup{
		if   ($GrepAdMembers -like "*VTOMPROD_TestOrdonnancement*"){
			$Alias = "Test" 
			return $Alias
		}
        elseif ($GrepAdMembers -like "*VTOMPROD_Consultation*"){
			$Alias = "Consult"
			return $Alias
		}
        elseif ($GrepAdMembers -like "*VTOMPROD_Ordonnancement*"){
			$Alias = "Ordo" 
			return $Alias
		}
        elseif ($GrepAdMembers -like "*VTOMPROD_Suivi&Relance*"){
			$Alias = "Suivi"
			return $Alias
		}
        elseif ($GrepAdMembers -like "*Qualiparc*"){
			$Alias = "Qualiparc"
			return $Alias
		}
        else {
			$Alias = "Groupe sans alias = $GrepAdMembers"
			return $Alias
		}
	} 
   
	# Function Full Export CSV
	function FullExport {
		$GetContentFile = Get-Content $ConfigFile # Load Distribution Group from init file
		$Group = "Groupe"
		foreach ($UneLigne in $GetContentFile) {  
			$GrepAdMembers = Get-ADObject -filter 'objectclass -eq "group"' -properties *  | Select-String "$UneLigne" # Grep of distribution group in all AD, then extract full absolute path
			if ($GrepAdMembers) { # If grep isn't empty, Distribution group exist, we continu 
				$Alias = Aliasgroup # Executing function 'Aliasgroup' for loading variable. 
				$RequestProfil = Get-ADGroupMember -Id "$GrepAdMembers" -Recursive | select @{n='SamAccountName';e={$_.SamAccountName}}, @{n=$Group;e={$Alias}} # Main request for extract
				$RequestProfil | Export-CSV -path $CsvFileTemp -Append -Encoding unicode -NoTypeInformation # Extract-csv 
			}
			else { Log "[WARNING] Le groupe $UneLigne n'est pas identié dans le fichier $Configini"	}
		}
		# Add in the CSV the colum 'Enabled' for each Users into the group previously define
		Import-Csv $CsvFileTemp | foreach-object {
				$SamName = $_.SamAccountName
				$RequestUser = Get-ADUser $SamName |Select-Object -property Enabled
				$ResultFilter = $RequestUser  | % { $_."Enabled" }
				$ResultUser = $_ | select *,@{n='Enabled';e={$ResultFilter}} | Export-csv $CsvFile -Append -Encoding unicode -NoTypeInformation # Extract-csv 			
		} 
	# Cleaning Temp File
	DeleteCsvFileTemp			
	}
	
	# Function Export Incremental CSV
	function IncrementalExport {
		## First incremental only check if user are added or deleted				
		$ImportPreviousCSV = Import-csv $Previouscsv  # Load last CSV in date 
		$ImportCurrentCSV = Import-csv $CsvFile 
		
		# Compare 2 arrays in two CSV "SamAccountName & Groupe" if users have been added or deleted
		$CompareUsers = Compare-Object -ReferenceObject $ImportPreviousCSV -DifferenceObject $ImportCurrentCSV -Property SamAccountName,Groupe -PassThru
		
		# If var $Compare is not empty, we're exporting, otherwise we're doing nothing
		if ( $CompareUsers ) {
			$CompareUsers | Export-CSV -path $CsvFileInc -NoTypeInformation -Encoding unicode  
			# File format : replace side indicator from 'Compare-Object'
			( Get-Content $CsvFileInc ) | Foreach-Object {			
				$_ -replace '=>', 'Ajout' -replace '<=', 'Suppression'
				} | out-file $CsvFileInc
		}
		
		## Second Compare focus only if users accounts are enabled(unlock) or disabled(lock)  
		# First we execute Reverse compare  
		$CompareReverse = Compare-Object -ReferenceObject $ImportCurrentCSV -ExcludeDifferent $ImportPreviousCSV -Property SamAccountName,Groupe -PassThru -IncludeEqual
		If ( $CompareReverse ) {
		$CompareReverse | Export-CSV -path $TempReverseInc -NoTypeInformation -Encoding unicode
		# Import again reverse compare for checking users enabled or disabled 
		$TempFileInc = Import-csv $TempReverseInc
		}
		# Then we run the difference between PreviousCSV and TempReverseInc for looking for users Enabled or Disabled
		If ( $TempFileInc ) {
		$CompareEnabled = Compare-Object -ReferenceObject $ImportPreviousCSV -DifferenceObject $TempFileInc -Property SamAccountName,Groupe, Enabled -PassThru -IncludeEqual | Where-Object { $_.SideIndicator -eq '=>' }
		}
		# If var $CompareReverse is not empty, we're exporting, otherwise we're doing nothing
		if ( $CompareEnabled ) {
			$CompareEnabled | Export-CSV -path $CsvFileInc -Append -NoTypeInformation -Encoding unicode
			$TempReplaceINC = Import-csv $CsvFileInc
			foreach( $test in $TempReplaceINC ) { 
			$field3 = $test.Enabled
					if ( $field3 -match "False" ) { $test.SideIndicator = $test.SideIndicator.replace('=>','Utilisateur Désactivé') }
					if ( $field3 -match "True" ) { $test.SideIndicator = $test.SideIndicator.replace('=>','Utilisateur Activé')	}
			}  
			$TempReplaceINC | export-csv $CsvFileInc  -NoTypeInformation -Encoding unicode
		}
		
		# If one or more of those vars is empty, print : 'No Diff' otherwise : 'Export Success'
		If ( $CompareUsers -Or $CompareEnabled ) { Log "[OK] Incr Export CSV Successfull !  INC_VTPD_$date.csv" }
		Else { Echo "[INFO] Aucune différence remarqué, donc pas de fichier Incrémental" } 
			
		# Cleaning Temp File
		DeleteTempReverseInc
	}
   
	#---------------------------------------------------------- 
	# Loading Modules
	#----------------------------------------------------------  
	# Module Export-CSV "-Append"
	# Load the function "-append" from command  'export-csv' which is not available inside powershell v2.0
	# Check if powershell >3, then load module append located in the script : $FunctionExportCsv
	if (3 -gt $host.version.major) { 
		If (Test-Path $ModuleExportCsvAppend){
			# The file is here, we load
			. $ModuleExportCsvAppend 
        }
		Else {
			Log "[ERROR] Be carefull can't load module '-append' in the script $append -> https://dmitrysotnikov.wordpress.com/2010/01/19/export-csv-append/"
			echo "" 
			Exit 1 
        }
	}
   
	# Loading Active Directory Module
	Try { get-module -listavailable | Where-Object {$_.name -like "ActiveDirectory*"} | import-module  -ErrorAction Stop } 
	# Chargement du module ActiveDirectory
    Catch { 
		Log "[ERROR] ActiveDirectory Module couldn't be loaded. Script will stop!" 
		Exit 1 
    } 
	
	###########################################################
	#---------------------------------------------------------- 
	# MAIN  ---    Extract Full / Incrémental 
	#---------------------------------------------------------- 
	###########################################################
   
	#---------------------------------------------------------- 
	# Screen Initialisation
	#---------------------------------------------------------- 
	cls # Clean screen   
	Write-Host "This program will extract the group AD you previously set in the file : $ConfigFile" -foregroundcolor RED
	Get-Date -format "dd-MMM-yyyy HH:mm"
	echo ""
	
	#_1 Check init File 
	try 	{ Initfile }
	catch	{ Log "[ERROR] Problem during initfile $Configini " }
	
	#_2 Check CSV Folder
	try 	{ Checkfolder }
	catch	{ Log "[ERROR] Problem during check $dircsv" }			
	
	#_3 Delete CSV Full & Inc of the same day 
	try 	{ DeleteCsvFile }
	catch	{ Log "[ERROR] Problem during delete :  $CsvFile " }	
	try 	{ DeleteCsvFileInc }
	catch	{ Log "[ERROR] Problem during delete : $CsvFileInc " }	
	
	#_4 Load Last CSV present in date for the compare-object (to execute before full export) 
	try		{ $Previouscsv = Get-ChildItem -path $Lastcsv | Sort-Object LastAccessTime -Descending | Select-Object -First 1
					If (!$Previouscsv){
						$Previouscsv = $CsvFile	
					}
	}
	catch	{ log "[WARNING] Problem during loading the last recent CSV : "}
	
	#_5 Launch Full Export
	try		{ FullExport 
				echo ""
				Log "[OK] Full Export CSV Successfull ! FULL_VTPD_$date.csv"
	}
	Catch	{
				Log "[ERROR] Erreur lors de l'export Full CSV"
				echo ""
	}
	
	#_6 Launch Incrémental if a previous full extract is present (at least day-1)
	try		{ 
	If ( $Previouscsv -eq $CsvFile ) {
							Log "[INFO] Impossible to compare file with only one file Full /QUIT "
							Exit 1 
				}
				Else { IncrementalExport }	
	 }
	Catch	{ Log "[ERROR] Problème durant l'export incrémental" }
	
	# --- End of Program  --- #
