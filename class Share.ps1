class Share {
    [PSCredential]$Credential
    [String]$ShareType;
    [String]$ShareLocation;
    [String]$ClientNumber;
    [String]$Matternumber;
    [String]$MatterName
    [String]$ClientName;
    [String]$DFSpath;
    [String]$ShortcutPath;
    [String]$ClientFolderPath
    [String]$MatterFolderPath
    [String]$ClientSecuritygroup
    [String]$MatterSecurityGroup
	
    # Constructor (Client Only)
	
    Share([String]$Sharetype, [String]$ShareLocation, [String]$ClientNumber) {
        $this.credential = $Global:Credential
        $this.ShareType = $ShareType
        $this.ShareLocation = $ShareLocation
        $this._validateClientNumber([String]$ClientNumber)
        $this._queryClientName([String]$ClientNumber)
        $this._getPaths()
        $this._buildClientFolderPath()
        $this._setClientSecurityGroup()
    }
	
    # Constructor (Client and Matter)
    Share([String]$Sharetype, [String]$ShareLocation, [String]$ClientNumber, [String]$Matternumber) {
        $this.credential = $Global:Credential
        $this.ShareType = $ShareType
        $this.ShareLocation = $ShareLocation
        $this._validateClientNumber([String]$ClientNumber)
        $this._queryClientName([String]$ClientNumber)
        $this._queryMatterName([String]$MatterNumber)
        $this._getPaths()
        $this._buildClientFolderPath()
        $this._buildMatterFolderPath()
        $this._setClientSecurityGroup()
        $this._setMatterSecurityGroup()
    }
	
    # Method: Validate Client Number (Constructor Step 1)
    hidden [void] _validateClientNumber([String]$ClientNumber) {
        IF ([String]::IsNullOrWhiteSpace($ClientNumber)) {
            Write-Error "Client Number cannot be blank."
        }
		
        ElseIf (($ClientNumber -as [Int]) -isnot [Int]) {
            Write-Error "Client Number is not all digits."
        }
    }
	
    # Method: Query Client Name (Constructor Step 1)
    hidden [void] _queryClientName([String]$ClientNumber) {
        $Query = $NUll
        Try {
            $Query = Get-ClientName -Credential $this.Credential -ClientNumber $ClientNumber
        }
        Catch {
            Write-Error $_
        }
		
        If ($Query -eq $False) {
            Write-Error "Client number is not in the database."
        }
        Else {
            $this.ClientNumber = $ClientNumber
            $this.clientname = Format-Name -ClientName $Query
        }
    }
	
    # Method: Query Matter Name
    hidden [void] _queryMatterName([String]$Matternumber) {
        $Query = $NUll
        Try {
            $Query = Get-MatterName -Credential $this.Credential -ClientNumber $this.ClientNumber -MatterNumber $Matternumber
        }
        Catch {
            Write-Error $_
        }
		
        If ($Query -eq $False) {
            Write-Error "Matter number is not in the database."
        }
        ElseIf ($Query -eq $Null -or $Query -eq '') {
            Write-Error "Matter number is not in the database."
        }
        Else {
            $this.MatterNumber = $MatterNumber
            $this.MatterName = $Query
        }
    }
	
    # Method: Get Paths
    hidden [void] _getPaths() {
        If ($this.ShareLocation -in ('AM1', 'AP1', 'EM1')) {
            Try {
                $Paths = Get-DatacenterPaths -ShareLocation $this.ShareLocation -ShareType $this.Sharetype
                $this.dfspath = $Paths.dfspath
                $this.shortcutpath = ($Paths.shortcutpath + '\' + $this.ClientName)
            }
            Catch {
                Write-Error $_
            }
        }
        Else {
            Try {
                $Paths = Get-LocalOfficePaths -ShareLocation $this.ShareLocation -ShareType $this.Sharetype
                $this.dfspath = $Paths.dfspath
                $this.shortcutpath = ($Paths.shortcutpath + "\" + $this.ClientName)
            }
            Catch {
                Write-Error $_
            }
        }
		
    }
	
    # Method: Build Client Folder Path
    hidden [void] _buildClientFolderPath() {
        $Folder = ($this.ClientNumber + '_' + $this.ClientName).TrimEnd("_")
        $this.ClientFolderPath = $this.DFSpath + '\' + $Folder
    }
	
    # Method: Build Matter Folder Path
    hidden [void] _buildMatterFolderPath() {
        $Folder = ($this.MatterNumber + '_' + $this.MatterName).TrimEnd("_")
        $this.MatterFolderPath = $this.ClientFolderPath + '\' + $Folder
    }
	
    hidden [Void] _setClientSecurityGroup() {
        Try {
            $CSGName = New-ClientSecurityGroupName -Location $this.sharelocation -ClientNumber $this.clientnumber
            $this.ClientSecurityGroup = $CSGName
        }
        Catch {
            Write-Error $_
        }
    }
	
    hidden [Void] _setMatterSecurityGroup() {
        Try {
            $MSGName = New-MatterSecurityGroupName -Location $this.sharelocation -ClientNumber $this.clientnumber -MatterNumber $this.Matternumber
            $this.MatterSecurityGroup = $MSGName
        }
        Catch {
            Write-Error $_
        }
    }
	
    # Method: Test Client Folder Path
    [Bool] testClientFolderPath() {
        If (Test-Path -Path $this.ClientFolderPath) {
            Return $True
        }
        Else {
            Return $False
        }
    }
	
    # Method: New Client Folder path
    [Void] newClientFolderPath() {
        Try {
            New-Item -ItemType Directory -Path $this.ClientFolderPath -Force -ErrorAction Stop
        }
        Catch [System.Exception] {
            Write-Error $_
        }
    }
	
    # Method: Test Matter Folder Path
    [Bool] testMatterFolderPath() {
        If (Test-Path -Path $this.MatterFolderPath) {
            Return $True
        }
        Else {
            Return $False
        }
    }
	
    # Method: New Matter Folder path
    [Void] newMatterFolderPath() {
        Try {
            New-Item -ItemType Directory -Path $this.MatterFolderPath -Force -ErrorAction Stop
        }
        Catch {
            Write-Error $_
        }
    }
	
    # Method: New Shortcut
    [Void] newShortcut() {
        Try {
            $FullShortcutPath = $this.Shortcutpath + '.lnk'
            $Shell = New-Object -ComObject ("WScript.Shell")
            $ShortCut = $Shell.CreateShortcut($FullShortcutPath)
            $ShortCut.TargetPath = $this.ClientFolderPath
            $ShortCut.WindowStyle = 1;
            $ShortCut.Save()
        }
        Catch [System.Exception] {
            Write-Error $_
        }
    }
	
}