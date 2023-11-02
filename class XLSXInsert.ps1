class XLSXInsert {
    [XLSX]$XLSXobj
    [Array]$CMDBServerInfoInserts
    [Array]$CMDBDatastoreInfoInserts

# Constructor (Client Only)

XLSXInsert([XLSX]$XLSXobj) {
    $this.XLSXobj = $XLSXobj
    [Array]$this.CMDBServerInfoInserts = @()
    [Array]$this.CMDBDatastoreInfoInserts = @()
}

# Method: Create General Information Inserts
[void] createCMDBServerInfoInserts([PSCustomObject]$VM) {
    Try{
        [String]$Hosted      	= ($VM.host -split "\.")[0]
        [String]$Type        	= 'Virtual'
        [String]$OS             = $this._GetRealOS($VM)      
        [pscustomobject]$NameProps  = Get-PropertiesFromServerName -ServerName $VM.vm
        [String]$Status           = ''
        [pscustomobject]$DatastoreProps  = Get-PropertiesFromDatastoreName -DatastoreName $VM.path
        [Array]$IPs         	= $this.XLSXobj.networkinfo | where-object {$_.VM -eq $VM.vm} | Select-Object -expandproperty 'IP Address'
        If($IPS){
            [String]$IPs = Get-IPAddressString -IPs $IPS
        }
        Else{
            $IPS = ''
        }
        [Int]$CPUCount       	= $this.XLSXobj.vcpuinfo     | where-object {$_.VM -eq $VM.vm} | Select-Object -expandproperty 'CPUs' | Select-Object -first 1
        [Int]$MemoryCount    	= $this.XLSXobj.meminfo      | where-object {$_.VM -eq $VM.vm} | Select-Object -expandproperty 'Size MB' | Select-Object -first 1
        $DeleteStatement     	= "Delete from [dbo].[ServerInfo] where name = '$($VM.VM)';`r`n"    
        $InsertHeader        	= "INSERT INTO [dbo].[ServerInfo] VALUES`r`n"
        $GenericInsert       	= "('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}');" -f $VM.VM, $NameProps.Location, $IPs, $Hosted, $NameProps.Function, $DatastoreProps.Replication, $Type, $OS, $NameProps.Role, $DatastoreProps.Datastore, $CPUCount, $MemoryCount, $NameProps.Environment, $Status
        $GenericInsert       	= $GenericInsert -replace "''","NULL"
        $CompleteInsert      	= $DeleteStatement + $InsertHeader + $GenericInsert
        $this.CMDBServerInfoInserts += $CompleteInsert
    }
    Catch {
        Write-Error $_.Exception.Message
    }
}

# Method: Determine Real OS Value
[String] _GetRealOS($VM) {
    If($VM.'OS according to the VMware Tools' -eq $VM.'OS according to the configuration file'){
        Return $VM.'OS according to the VMware Tools'
    }
    elseif($VM.'OS according to the VMware Tools' -ne '' -and $VM.'OS according to the VMware Tools'-ne $Null) {
        Return $VM.'OS according to the VMware Tools'
    }
    Else{
        Return $VM.'OS according to the configuration file'
    }
}

# Method: Create Datastore Information Inserts
[void] createCMDBDatastoreInfoInserts([PSCustomObject]$Datastore) {
    Try{
        [Int]$NumHosts          = $Datastore.'# VMs'
        [Int]$NumVMs            = $Datastore.'# Hosts'	
        [Int]$FreePercent       = $Datastore.'Free %'
        $DeleteStatement     	= "Delete from [dbo].[DatastoreInfo] where name = '$($Datastore.Name)';`r`n"    
        $InsertHeader        	= "INSERT INTO [dbo].[DatastoreInfo] VALUES`r`n"
        $GenericInsert       	= "('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}');" -f $Datastore.Name, $NumVMs, $NumHosts, $Datastore.Hosts, $Datastore.'Provisioned MB', $Datastore.'Capacity MB', $Datastore.'In Use MB', $Datastore.'Free MB'
        $GenericInsert       	= $GenericInsert -replace "''","NULL"
        $CompleteInsert      	= $DeleteStatement + $InsertHeader + $GenericInsert
        $this.CMDBDatastoreInfoInserts += $CompleteInsert
    }
    Catch{
        Write-Error $_.Exception.Message
    }
}

}