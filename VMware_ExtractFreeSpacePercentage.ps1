#Extract the free space in percentage
Get-Datastore | Select @{N="DataStoreName";E={$_.Name}},@{N="Percentage Free Space(%)";E={[math]::Round(($_.FreeSpaceGB)/($_.CapacityGB)*100,2)}} | Where {$_."Percentage(<20%)" -le 20}
