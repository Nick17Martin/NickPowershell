#Export information users's from AD in .csv
#Nicolas Martinelli

#Set you Organizational Unit
$OrganizationalUnit = "Insert you OU"
#Set Output Path
$output = "insert your output path and file name.csv"
#Initialize Array
$ObjectOutput =@()
$Prop =@()
#Extract Information from Active Directory
Write-host "Collecting information...please wait" - foreground "Yellow"
$List = Get-ADUser -Filter * -Properties DisplayName,samaccountname,department,employeeType,CanonicalName,PasswordLastSet,Enabled,accountExpires,AccountExpirationDate,whenCreated,whenChanged,LastLogonDate,extensionAttribute3,memberof -searchbase $OrganizationalUnit -searchscope 2

Foreach ($i in $List){
	$MO = $i.MemberOf
	$GroupMembership = ($MO | % { (Get-ADGroup $_).Name; }) -join ',';
#Set properties
	$info = @{
				Gruppo = $GroupMembership
				DisplayName = $i.DisplayName
				SamAccountName = $i.samaccountname
				Department = $i.department
				employeeType = $i.employeeType
				CanonicalName = $i.CanonicalName
				PasswordLastSet = $i.PasswordLastSet
				Enabled = $i.Enabled
				accountExpires = $i.accountExpires
				AccountExpirationDate = $i.AccountExpirationDate
				whenCreated = $i.whenCreated
				whenChanged = $i.whenChanged
				LastLogonDate = $i.LastLogonDate
				extensionAttribute3 = $i.extensionAttribute3
		}
		$ObjectOutput += New-Object -TypeName PSobject -Property $info
	}
Write-host "Exporting information in .csv file" -foreground "Yellow"
$ObjectOutput | Export-csv FinalTest.csv -Delimiter ";" -NoTypeInformation 
