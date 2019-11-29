#Explosion Shared mailboxes groups's and setting directly permission to  users
#Martinelli Nicolas 03/10/2017

Write-host "---STARTING SCRIPT SHARED MAILBOX ---" -foreground green
###############################--LOADING MODULES--###############################
Import-Module ActiveDirectory
#Chek if the file already exist, if it exist I rename it
$ExistFile = Test-Path C:\Temp\PermissionExport.csv
if ($ExistFile -eq "True"){
			rename-item "C:\Temp\PermissionExport.csv" "C:\Temp\PermissionExport_OLD.csv"}
####################################################################################
Write-host "1_STARTING EXTRACTING DATA FROM EXCHANGE" -foreground green
###############################--ESTRAZIONE DATI DA EXCHANGE--###############################
$OutFile = "C:\Temp\PermissionExport.csv"

"DisplayName" + ";" + "Alias" + ";" + "FullAccess" + ";" + "SendAs" | Out-File $OutFile -Force
Write-Host "EXTRACTING IN PROGRESS,PLEASE WAIT..."
$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Select Identity, Alias, DisplayName, DistinguishedName
ForEach ($Mailbox in $Mailboxes) {
	$SendAs = Get-ADPermission $Mailbox.DistinguishedName | ? {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY\SELF" -and !$_.IsInherited} | % {$_.User}
	$FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq "FullAccess" -and !$_.IsInherited} | % {$_.User}
	$Mailbox.DisplayName + ";" + $Mailbox.Alias + ";" + $FullAccess + ";" + $SendAs | Out-File $OutFile -Append
}
##########################################################################################	
Write-host "__FINISH EXTRACTING DATA FROM EXCHANGE" -foreground green
Write-host "2_START CREATING FILE GUIDE" -foreground green
###############################--CREAZIONE FILE DATI .CSV--###############################
#Ripulisco il file dati per poterlo elaborare
#Clean data file's for next elaboration (in my case i've an entry to remove before next step)
(Get-Content $OutFile) -replace "`XXX\\", '' | Set-Content "D:\E2K10\work\PermessiShared.csv"

$ProcessingFile = "Path\PermessiShared.csv"
#Importing data file and variables initialization
$Elenco = Import-csv -Delimiter ";" $ProcessingFile
$Shared = $Elenco | Select alias
$Perm_SendAS = $Elenco |  Select SendAs
$El_SendAS = $Elenco | Select alias, SendAs

$Perm_FullaAccess = $Elenco |  Select FullAccess
$EL_FullAccess = $Elenco | Select alias, FullAccess
##########################################################################################	
Write-host "__FINISH CREATE FILE GUIDE" -foreground green
Write-host "3_STARTING PROCESSING SEND AS PHASE" -foreground green
###############################--SECTION SENDAS--###############################
foreach ($i in $EL_SendAs){
		Write-host "Processing Mailbox: "$i.alias -foreground yellow
		If ($i.SendAs -ne $null -and $i.SendAS -notlike ""){
			$AllInfSendAs = Get-ADGroupMember $i.SendAs | select Objectclass,samaccountname
			$Recursive_GroupSendAS = $AllInfSendAs| Where-Object {$_.objectClass -eq 'group'}
			$User_SendAs = $AllInfSendAs | Where-object {$_.objectClass -eq 'user'}
			$U_SendAS_SAM = $User_SendAs |select samaccountname
			$U_SendAS_SAM
			foreach($USA in $User_SendAs){
				Add-ADPermission -Identity $i.alias -User $USA.samaccountname -AccessRights ExtendedRight -ExtendedRights "Send As"
				Write-host "Add-ADPermission -Identity" $i.alias "-User" $USA.samaccountname "-AccessRights ExtendedRight -ExtendedRights Send As"
				}
			if ($Recursive_GroupSendAS -ne $null)
			{
				$SAM_SA = $Recursive_GroupSendAS.samaccountname
				Write-host "Explosion group $SAM_SA"
				$UserRec_SendAs = Get-ADGroupMember $Recursive_GroupSendAS.samaccountname | select samaccountname
				foreach($URSA in $UserRec_SendAs){
					Add-ADPermission -Identity $i.alias -User $URSA.samaccountname -AccessRights ExtendedRight -ExtendedRights "Send As"
					Write-host "Add-ADPermission -Identity" $i.alias "-User" $URSA.samaccountname "-AccessRights ExtendedRight -ExtendedRights Send As"
				}
			} 
			else{Write-host "No Group like members" -foreground magenta}
		$Recursive_GroupSendAS = $null
		$U_SendAS_SAM = $null
		}
		else{ Write-Host "No members with SEND AS permission for this shared mailbox"}
	}
################################################################################
Write-host "__FINISH SEND AS PHASE" -foreground green
Write-host "4_STARTING PROCESSING FULL ACCESS PHASE" -foreground green
###############################--SECTION FULL ACCESS--###############################		
foreach ($i in $EL_FullAccess){
		Write-host "Processing Mailbox: "$i.alias -foreground yellow
		If ($i.FullAccess -ne $null -and $i.FullAccess -notlike ""){
			$AllInfFullAccess = Get-ADGroupMember $i.FullAccess | select Objectclass,samaccountname
			$Recursive_GroupFullAccess = $AllInfFullAccess| Where-Object {$_.objectClass -eq 'group'}
			$User_FullAccess = $AllInfFullAccess | Where-object {$_.objectClass -eq 'user'}
			$U_FullAccess_SAM = $User_FullAccess |select samaccountname
			$U_FullAccess_SAM
			foreach($UFA in $User_FullAccess){
				Add-MailboxPermission -Identity $i.alias -User $UFA.samaccountname -AccessRights FullAccess -InheritanceType All
				Write-host "Add-MailboxPermission -Identity" $i.alias "-User" $UFA.samaccountname "-AccessRights FullAccess -InheritanceType All"
				}
			if ($Recursive_GroupFullAccess -ne $null)
				{
				$SAM_FA = $Recursive_GroupFullAccess.samaccountname
				Write-host "Explosion group $SAM_FA"
				$UserRec_FullAccess = Get-ADGroupMember $Recursive_GroupFullAccess.samaccountname | select samaccountname
				foreach($URFA in $UserRec_FullAccess){
					Add-MailboxPermission -Identity $i.alias -User $URFA.samaccountname -AccessRights FullAccess -InheritanceType All
					Write-host "Add-MailboxPermission -Identity" $i.alias "-User" $URFA.samaccountname "-AccessRights FullAccess -InheritanceType All"
					}
				} 
			else {Write-host "No group like members" -foreground magenta}
			#Set variable to null
			$Recursive_GroupFullAccess = $null
			$U_FullAccess_SAM = $null
		}
		else { Write-Host "No members with FULL ACCESS permissions for this shared mailbox"}
	}	
################################################################################
Write-host "__FINISH FULL ACCESS PHASE" -foreground green
Write-host "---END SCRIPT SHARED MAILBOX---"
