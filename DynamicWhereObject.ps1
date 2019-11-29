#Initialize variable
$var1 = "Mario"
$var2 = "Rossi"
#Define array
$WhereArray = @()
#Check if variable is empty
if ($var1 -ne $null) {$WhereArray+='$_.var1 -like $var1'}
if ($var2 -ne $null) {$WhereArray+='$_.var1 -like $var2'}
#Build array in a string that concatenate every object with "-and"
$WhereString = $WhereArray -Join " -and "
#Build scriptblock with final string
$WhereBlock = [scriptblock]::Create( $WhereString )
#How to use in your code
? {Invoke-Command $WhereBlock}
#if a variable "$a" valorized with a .csv file data
$a | ? {Invoke-Command $WhereBlock}
