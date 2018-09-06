
# -----------------------------------------------------------------------------
# Script: Get-Metadata
# Version: 2.1
# Author: Tin Aung Phyoe (taphyoe@gmail.com)
# Created Date: 05/09/2018
# Keywords: Metadata, Storage, Files
# Credits: Antonio Gatti
# Last Modified: 06/09/2018 11:16 AM
# comments: to get file metadata and folder size
# Gets all the metadata and returns a CSV file (with | separation)
# 
# PARAMETERS:
# -folder: folder to be scanned
# -depth: set how deep you want the script to go and get the file metadata. After this level it will only return the folder size. (default value is 5)
# -maFileRows: set the max number of rows you want in the CSV. It will generate another file every time the script reach the limit.
#
# -----------------------------------------------------------------------------


	param(
	[string[]]$folder,
	[int]$depth = 5,
	[int]$maxFileRows = 300000
	)
	
	#File Increment Number Eg. result20180905_1.csv
    $fileCount = 1
	$cd = (Get-Item -Path ".\").FullName
	$r = $cd + "\result"+ ((Get-Date).tostring("yyyyMMdd")) + "_" +$fileCount +".csv"
	#File/Folder properties that we are interested in. Check Headers for detail.
	$fileproperties = 0,1,2,3,4,5,6,10,11,19,165,166,191,192,193,195
	$folderCounts =0;
	$rowCounts = 0
	$mainFolders = [System.Collections.ArrayList]@()
	$foldersNotFound = [System.Collections.ArrayList]@()
	$folders = [System.Collections.ArrayList]@()
	

	Write-Output "Start Time $((Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt'))"
	Write-Output "Folder Level|Name|Item Type|Size (MB)|Type|Date modified|Date created|Date accessed|Attributes|Owner|Kind|Rating|File extension|Filename|Folder name|Folder path|Folder|Path|Status" >> $r 
	Write-Output "Retrieving Sub Folders..."
	
	#Scanning sub folders and folder permission error for last level
	foreach($sfolder in $folder){
	
		$sfolderTrimmed = $sfolder.TrimEnd('\') #remove the last \ in the path
		
		
		if(Test-Path -Path $sfolderTrimmed){#Check if folder exist

			$folders+=$sfolderTrimmed;
			$mainFolders+=$sfolderTrimmed;

			 if($depth -eq 0){
			 $folders += (gci  -Recurse $sfolderTrimmed  -ErrorAction SilentlyContinue -ErrorVariable gciErrorsMain | ?{ $_.PSIsContainer }).FullName
			 }
			 else{
			  if($depth -gt 1){
				$folders += (gci -Depth ($depth-2) -Recurse $sfolderTrimmed  -ErrorAction SilentlyContinue -ErrorVariable gciErrorsMain | ?{ $_.PSIsContainer }).FullName
			  }
			  else{
				#$folder += (gci $sfolder  -ErrorAction SilentlyContinue -ErrorVariable gciErrorsMain | ?{ $_.PSIsContainer }).FullName
			  }
			  # Scanning one more level for permission error
			  $scanForError = gci -Depth ($depth) -Recurse $sfolderTrimmed  -ErrorAction SilentlyContinue -ErrorVariable gciErrorsMain | ?{ $_.PSIsContainer }
			 }
		 }
		 else{
			
			[void]$foldersNotFound.Add(''+$sfolderTrimmed +' is not found!')
		 }
	}
	
	

	foreach($sFolder in $folders){
	
	
		$folderCounts +=1;
		Write-Output "Checking files in $sFolder $folderCounts/$($folders.Count)"
		
		$folderLevel = 0;
		
		#Define which level are we at 
		if($mainFolders.Contains($sFolder)) {
			$folderLevel = 1
		}
		else{
			$mainfolderlevel = ($mainFolders[0].ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
			$subfolderlevel = ($sFolder.ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
			$folderLevel =  ($subfolderlevel - $mainfolderlevel)+1;
		}#End Define which level are we at 
		
		
		$a = 0
		$objShell = New-Object -ComObject Shell.Application
		$objFolder = $objShell.namespace($sFolder)
		$currentFolderFileCounts =0
		$currentFolderTotalFileCounts =$objFolder.items().Count
		
		#Get all files in a folder
		foreach ($File in $objFolder.items()){
			
			#Powershell Progress Status
			$currentFolderFileCounts+=1
			$pc = [math]::Round(($currentFolderFileCounts/$currentFolderTotalFileCounts*100),2)
			$ac = "Percentage Complete: " + $pc + "% | " + "Checking File " + $currentFolderFileCounts + " of " + $currentFolderTotalFileCounts
			
			Write-Progress -Activity $ac  -PercentComplete ($currentFolderFileCounts/$currentFolderTotalFileCounts*100)
			#End Powershell Progress Status
			
			$isWriteToFile = 1;
			
			
			$hash= ' ' +$folderLevel +'|';
			
			#To decide create new file 
			if($rowCounts -gt ($maxFileRows-2)){
				$rowCounts=0
				$fileCount+=1
				$r = $cd + "\result"+ ((Get-Date).tostring("yyyyMMdd")) + "_" +$fileCount +".csv"
				Write-Output "Folder Level|Name|Item Type|Size (MB)|Type|Date modified|Date created|Date accessed|Attributes|Owner|Kind|Rating|File extension|Filename|Folder name|Folder path|Folder|Path|Status" >> $r 
			}#End To decide create new file 
			
			Try{
				if($objFolder.getDetailsOf($File, 165) -or ($depth) -eq $folderLevel ){ #Check if it is file or we are at last folder level
					$isWriteToFile =1;
					foreach ($a in $fileproperties){
						if($a -eq 1){  # Property 1 is Size, For size we have to get differently for file and folder 
						
							if($objFolder.getDetailsOf($File, 165)){
								$hash += "File|"
								$hash += ([math]::Round((Get-Item $objFolder.getDetailsOf($File, 195) -ErrorAction SilentlyContinue -ErrorVariable gciErrorsFileSize  ).length/1MB,4))  
							}
							else{
								$hash += "Folder|"
								$hash+="{0:N2}" -f ((Get-ChildItem $File.Path -Recurse -ErrorAction SilentlyContinue -ErrorVariable gciErrorsLastLevel | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB)
							}
							
							$hash += " |"
						}# End Property 1 is Size, For size we have to get differently for file and folder 
						else{
							if($objFolder.getDetailsOf($File, $a)){ #All other properties
									 $hash += $objFolder.getDetailsOf($File, $a) + '|'
							}
							else{
								$hash += ' '+ '|'
							}
						}
					} #end for $fileproperties
					$hash +='Ok'
				}
				else{
					$isWriteToFile = 0; #If not file or last folder level , don't write out flag to 0
				}
			}#End Try
			Catch{
				$hash =  ' ' +$folderLevel +'|'+$File.Path+"|||||||||||||||||" +  $_.Exception.Message
			 }#End Catch
			 
			 if($isWriteToFile -eq 1) {#If not file or last folder level , don't write to file
				$rowCounts +=1;
				$hash | Out-File $r -Append
			 }#End If not file or last folder level , don't write to file
			 
			 
			 $a=0
		} #end foreach $file
	} #end foreach $folder
  
  #Error Folders scanning first level folder
  foreach ($errorRecord in $gciErrorsMain){
	"0|Error|||||||||||||||||" + ($errorRecord -split '\n')[0]  | Out-File $r -Append
  }
  
   #Error Folders scanning last level folder
   foreach ($errorRecord in $gciErrorsLastLevel){
	"0|Error|||||||||||||||||" + ($errorRecord -split '\n')[0]  | Out-File $r -Append
  }
  
   #Error files/folders while getting information
   foreach ($errorRecord in $gciErrorsFileSize){
	"0|Error|||||||||||||||||" + ($errorRecord -split '\n')[0]  | Out-File $r -Append
	}
	
   #Error folders not found
   foreach ($errorRecord in $foldersNotFound){
	"0|Error|||||||||||||||||" + ($errorRecord -split '\n')[0]  | Out-File $r -Append
  }
    
  Write-Output "End Time $((Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt'))"
  