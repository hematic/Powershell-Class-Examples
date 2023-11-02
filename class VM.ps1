class VM {
	
    [String]$VMName;
    [String]$ID
    [String]$PowerState;
    [String]$Notes;
    [String]$OS;
    [String]$ToolsVersion;
    [Array]$NetAdapters;
    [Array]$VMDisks;
    [Array]$IPs;
    [Array]$HardDisks;
    [PSCustomObject]$Resources;
    [PSCustomObject]$VMHost;
    $VMObject;
	
    VM([String]$VMName) {
        $this._getVM($VMName)
    }
	
    # Method: Get VM Information
    hidden [void] _getVM([String]$VMName) {
        $VM = $NUll
        Try {
            $VM = Get-VM $VMName -ErrorAction Stop
            $View = Get-View -id $VM.id -ErrorAction Stop
        }
        Catch {
			
            switch -wildcard ($_.exception.message) {
                '*You are not currently connected to any servers*' {
                    Throw "No VCenter Connection is available."
                }
                '*was not found using the specified filter(s)*' {
                    Throw "No VM of that name exists on the VCenter server you are connected to."
                }
                '*is not connected*' {
                    Throw "The VCenter Connection has expired or broken."
                }
                Default {
                    Throw "Something Went Wrong : $($_.exception.message)"
                }
            }
        }
		
        $this.VMName = $VM.Name
        $this.ID = $VM.ID
        $this.PowerState = $VM.Powerstate
        $this.OS = $VM.guest.OSFullName
        $this.Notes = $VM.Notes
        $this.Resources = New-Object -TypeName PSCustomObject -Property @{
            'Number of CPUs'   = $VM.NumCpu
            'Cores Per Socket' = $VM.CoresPerSocket
            'Memory (MB)'      = $VM.MemoryMB
        }
        $this.IPs = $VM.guest.IPAddress
        $this.VMdisks = $VM.guest.Disks
        $this.ToolsVersion = $VM.guest.ToolsVersion
        $this.VMobject = $VM
        $this._getNetworkAdapters()
        $this._getHostInformation($VM.VMHost)
        $this._getHardDiskInfo()
    }
	
    # Method: Get VM Host Information
    hidden [void] _getHostInformation([String]$VMHostName) {
		
        Try {
            $HostInfo = Get-VMHost -Name $VMHostName -ErrorAction Stop
        }
        Catch {
			
            switch -wildcard ($_.exception.message) {
                '*You are not currently connected to any servers*' {
                    Write-Error "No VCenter Connection is available."
                }
                '*was not found using the specified filter(s)*' {
                    Write-Error "No Host of that name exists on the VCenter server you are connected to."
                }
                Default {
                    Write-Error "$($_.exception.message)"
                }
            }
            return;
        }
		
        $this.VMHost = New-object -TypeName PSCustomObject -Property @{
            'Name'                  = $HostInfo.name
            'NumCPU'                = $HostInfo.NumCpu
            'Total Memory (MB)'     = $HostInfo.MemoryTotalMB
            'Used Memory (MB)'      = $HostInfo.MemoryUsageMB
            'Processor Type'        = $HostInfo.ProcessorType
            'HyperThreading Active' = $HostInfo.HyperthreadingActive
            'Parent'                = $HostInfo.Parent
        }
    }
	
    # Method: Get hardDisk Information
    hidden [void] _getHardDiskInfo() {
		
        Try {
            $Disks = Get-HardDisk -VM $this.VMobject -ErrorAction Stop
        }
        Catch {
			
            switch -wildcard ($_.exception.message) {
                '*You are not currently connected to any servers*' {
                    Write-Error "No VCenter Connection is available."
                }
                '*Cannot convert*' {
                    Write-Error "The object passed to Get-HardDisk was not a proper VM Object of type 'VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine'"
                }
                Default {
                    Write-Error "$($_.exception.message)"
                }
            }
            return;
        }
		
        [Array]$DiskArray = @()
        Foreach ($Disk in $Disks) {
            $DiskArray += new-object -TypeName PSCustomObject -Property @{
                'StorageFormat' = $Disk.StorageFormat
                'Persistence'   = $Disk.Persistence
                'Filename'      = $Disk.Filename
                'CapacityKB'    = $Disk.CapacityKB
                'CapacityGB'    = $Disk.CapacityGB
            }
        }
        $this.HardDisks = $DiskArray
    }
	
    # Method: Get NetworkAdapters
    hidden [void] _getNetworkAdapters() {
		
        Try {
            $this.NetAdapters = Get-NetworkAdapter -VM $this.VMobject
        }
        Catch {
            Throw "Unable to get network adapaters : $($_.exception.message)"
        }
    }
	
    # Method: Restart VM
    [String] Restart() {
		
        Try {
            Restart-VM -VM $this.VMName -Confirm:$false -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "VM Restart initialized."
            Return $Null
        }
        Catch {
            Throw "Unable to Restart VM : $($_.exception.message)"
        }
    }
	
    # Method: ShutDown VM
    [String] ShutDown() {
		
        Try {
            Stop-VM -VM $this.VMname -Confirm:$False -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "VM Shutdown initialized."
            Return $Null
        }
        Catch {
            Throw "Unable to ShutDown VM : $($_.exception.message)"
        }
    }
	
    # Method: Restart Guest OS
    [String] RestartGuest() {
		
        Try {
            $this.VMObject | Restart-VMGuest -Confirm:$False -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "VM Restart initialized."
            Return $Null
        }
        Catch {
            switch -wildcard ($_.exception.message) {
                '*The attempted operation cannot be performed in the current state (Powered off)*' {
                    Throw "Unable to Restart Guest OS : The VM is in a Powered Off State."
                }
                Default {
                    Throw "Unable to Restart Guest OS : $($_.exception.message)"
                }
            }
            Throw "Unable to Restart Guest OS : $($_.exception.message)"
        }
    }
	
    # Method: ShutDown Guest OS
    [String] ShutDownGuest() {
		
        Try {
            $this.vmobject | ShutDown-VMGuest -Confirm:$False -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "VM Guest OS Shutdown initialized."
            Return $Null
        }
        Catch {
            switch -wildcard ($_.exception.message) {
                '*The attempted operation cannot be performed in the current state (Powered off)*' {
                    Throw "Unable to Shutdown Guest OS : Guest OS is already powered off."
                }
                Default {
                    Throw "Unable to Shutdown Guest OS : $($_.exception.message)"
                }
            }
            Throw "Unable to Shutdown Guest OS : $($_.exception.message)"
        }
    }
	
    # Method: Start VM
    [String] Start() {
		
        Try {
            Start-VM -VM $this.VMName -Confirm:$false -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "VM initialized."
            Return $Null
        }
        Catch {
            switch -wildcard ($_.exception.message) {
                '*The attempted operation cannot be performed in the current state (Powered on)*' {
                    Throw "Unable to Start VM : The VM is already on."
                }
                Default {
                    Throw "Unable to Start VM : $($_.exception.message)"
                }
            }
            Throw "Unable to Start VM: $($_.exception.message)"
        }
    }
	
    # Method: Set VM CPu Count
    [String] SetCPUCount([Int]$CPUCount) {
		
        Try {
            $this.VMObject | Set-VM -NumCpu $CPUCount -Confirm:$false -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "Set VM CPU Count to $CPUCount."
            Return $Null
        }
        Catch {
            switch -wildcard ($_.exception.message) {
                '*CPU hot plug is not supported for this virtual machine*' {
                    Throw "Unable to Set VM CPU Count : CPU hot plug is not supported for this virtual machine."
                }
                Default {
                    Throw "Unable to Set VM CPU Count : $($_.exception.message)"
                }
            }
            Throw "Unable to Set VM CPU Count : $($_.exception.message)"
        }
    }
	
    # Method: Set VM VRAM Amount
    [String] SetRAMinGB([Int]$RamInGB) {
		
        Try {
            $this.VMObject | Set-VM -MemoryGB $RamInGB -Confirm:$false -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "Set VM memory to $RamInGB."
            Return $Null
        }
        Catch {
            Throw "Unable to Set VM Ram : $($_.exception.message)"
        }
    }
	
    # Method: Add New Disk
    [String] AddDisk([Int]$CapacityGB, [String]$StorageFormat) {
		
        Try {
            $this.VMObject | New-HardDisk -CapacityGB $CapacityGB -StorageFormat $StorageFormat
            Write-RichText -LogType Success -LogMsg "Added new Disk."
            Return $Null
        }
        Catch {
            Throw "Unable to Add New Disk : $($_.exception.message)"
        }
    }
	
    # Method: Update VMWare Tools
    [String] UpdateTools() {
		
        Try {
            $this.VMObject | Update-Tools -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "VMWare tools Updated."
            Return $Null
        }
        Catch {
            Throw "Unable to update VMWare tools : $($_.exception.message)"
        }
    }
	
    # Method: Set Description
    [String] SetDescription([String]$Description) {
		
        Try {
            $this.VMObject| Set-annotation -CustomAttribute Description -Value $Description -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "VM Description Set"
            Return $Null
        }
        Catch {
            Throw "Unable to set VM Description : $($_.exception.message)"
        }
    }
	
    # Method: Set Notes
    [String] SetNotes([String]$Notes) {
		
        Try {
            $this.VMObject | Set-VM -Notes $Notes -Confirm:$False -ErrorAction Stop
            Write-RichText -LogType Success -LogMsg "VM notes Set"
            Return $Null
        }
        Catch {
            Throw "Unable to set VM Notes : $($_.exception.message)"
        }
    }
	
}
#Test for tfs 7