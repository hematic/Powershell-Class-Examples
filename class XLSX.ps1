class XLSX {
    [String]$FilePath
    [PSCustomObject]$Info
    [PSCustomObject]$vCPUInfo
    [PSCustomObject]$MemInfo
    [PSCustomObject]$DiskInfo
    [PSCustomObject]$PartitionInfo
    [PSCustomObject]$NetworkInfo
    [PSCustomObject]$SnapShotInfo
    [PSCustomObject]$VMWareToolsInfo
    [PSCustomObject]$ClusterInfo
    [PSCustomObject]$ResourcePoolInfo
    [PSCustomObject]$HostInfo
    [PSCustomObject]$HBAInfo
    [PSCustomObject]$NICInfo
    [PSCustomObject]$DataStoreInfo

    # Constructor (Client Only)

    XLSX([String]$FilePath) {
        $this.FilePath = $FilePath
        $this._parseInfo()
    }

    # Method: Parse General Information (Constructor Step 1)
    hidden [void] _parseInfo() {
        Try{
            $this.Info = Import-Excel -Path $this.FilePath -WorkSheetname "vInfo" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error importing Excel file."
            Write-Error $_.Exception.Message

        }
    }

    # Method: Parse Vcpu Information
    [Void] parsevCPUInfo() {
        Try{
            $this.vCPUInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vcpu" -ErrorAction Stop

        }
        Catch [System.Exception] {
            Write-Host "Error parsing vCPU worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse Memory Information
    [Void] parseMemInfo() {
        Try{
            $this.MemInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vMemory" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vMemory worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse Disk Information
    [Void] parseDiskInfo() {
        Try{
            $this.DiskInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vDisk" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vdisk worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse Partition Information
    [Void] parsePartitionInfo() {
        Try{
            $this.PartitionInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vPartition" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vPartition worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse Network Information
    [Void] parseNetworkInfo() {
        Try{
            $this.NetworkInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vNetwork" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vNetwork worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse SnapShot Information
    [Void] parseSnapShotInfo() {
        Try{
            $this.SnapShotInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vSnapshot" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vSnapshot worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse VMTools Information
    [Void] parseVMToolsInfo() {
        Try{
            $this.VMWareToolsInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vTools" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vTools worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse Cluster Information
    [Void] parseClusterInfo() {
        Try{
            $this.ClusterInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vCluster" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vCluster worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse Resource Pool Information
    [Void] parseResourcePoolInfo() {
        Try{
            $this.ResourcePoolInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vrp" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vResourcePool worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse Host Information
    [Void] parseHostInfo() {
        Try{
            $this.HostInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vHost" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vHost worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse HBA Information
    [Void] parseHBAInfo() {
        Try{
            $this.HBAInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vHBA" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vHBA worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse NIC Information
    [Void] parseNICInfo() {
        Try{
            $this.NICInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vNic" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vnic worksheet."
            Write-Error $_.Exception.Message
        }
    }

    # Method: Parse Datastore Information
    [Void] parseDatastoreInfo() {
        Try{
            $this.DataStoreInfo = Import-Excel -Path $this.FilePath -WorkSheetname "vDatastore" -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Host "Error parsing vdatastore worksheet."
            Write-Error $_.Exception.Message
        }
    }
}