$ClusterVM = Get-Cluster | Get-VM
$Output =@()

foreach($i in $ClusterVM){
    $Clu = $i | select @{Name='Cluster';Expression={$_.VMHost.Parent}}
    $VM = $i.name
	#Get disk information
    $HD = get-vm $VM | get-harddisk | select CapacityGB
	#Insert all value in one variable
    $HDGB = ($HD | % { ($_).CapacityGB; }) -join ',';
    $Information =@{
            DisplayName = $i.name
            Disk_GB = $HDGB
            Cluster= $Clu.Cluster
            CPU= $i.NumCpu
            Memory= $i.MemoryGB
}
$Output += New-Object -TypeName PSobject -Property $Information 
}
$Output | Select DisplayName,Disk_GB,CPU,Memory,Cluster | Export-Csv OutputFile.csv -Delimiter ";" -NoTypeInformation 
