class Mailbox {
	
    [String]$SamAccountName;
    [String]$Servername;
    [String]$Database;
    [String]$Guid;
    [String]$EmailAddresses;
    [Int]$DeletedItemcount;
    [Int]$ItemCount;
    [String]$TotalDeletedItemSize;
    [String]$TotalItemSize;
    [String]$LegacyExchangeDN;
    [String]$Displayname;
    [Bool]$HiddenFromGAL;
    [String]$office;
    [String]$PrimarySMTPAddress

	
	Mailbox([String]$SAMAccountName)
	{
		$this._getMailbox($SAMAccountName)
	}
	
	# Method: Get Mailbox Information
	hidden [void] _getMailbox([String]$SAMAccountName)
	{
		$Mailbox = $NUll
		Try{
			$Mailbox = Get-Mailbox $SAMAccountName -ErrorAction Stop
		}
		Catch{
			Throw "No Mailbox matches that SAMAccountName"
		}
		
        [String]$this.Database           = $Mailbox.Database
        [String]$this.SamAccountName     = $Mailbox.SamAccountName
        [String]$this.Guid               = $mailbox.Exchangeguid
        [String]$this.servername         = $Mailbox.Servername
        [String]$this.emailaddresses     = $Mailbox.emailaddresses
        [String]$this.LegacyExchangeDN   = $Mailbox.legacyexchangedn
        [Bool]$this.HiddenFromGAL        = $Mailbox.HiddenFromAddressListsEnabled
        [String]$this.PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress

        $this._getmailboxstatistics()
        $this._getaduser()
	}
	
	# Method: Get Mailbox Statistics
	hidden [void] _getmailboxstatistics()
	{
		Try
		{
			$Stats = Get-MailboxStatistics -Identity $this.SamAccountName -ErrorAction Stop
			
            $this.DeletedItemcount     = $Stats.deleteditemcount;
            $this.ItemCount            = $Stats.ItemCount;
            $this.TotalDeletedItemSize = $Stats.TotalDeletedItemSize;
            $this.TotalItemSize        = $Stats.TotalItemSize;

		}
		Catch
		{
			Throw "Unable to get mailbox statistics : $($_.exception.message)"
		}
	}

	# Method: Get Mailbox Statistics
	hidden [void] _getAdUser()
	{
		Try
		{
			$User = Get-Aduser -Identity $this.SamAccountName -Properties * -ErrorAction Stop
			
            $this.Office      = $User.office;
            $this.displayname = $User.displayname


		}
		Catch
		{
			Throw "Unable to get mailbox statistics : $($_.exception.message)"
		}
	}

	# Method: Add Mailbox Permission
	[String] AddPermission($User,$Rights)
	{
		Try
		{
            $Splat = @{
                Identity = $this.SamAccountName
                User = $User
                AccessRights = $Rights
                InheritanceType = 'All'
                ErrorAction = 'Stop'

            }
			Add-MailboxPermission @Splat
			Return $Null
		}
		Catch
		{
			Throw "Unable to Add Permission : $($_.exception.message)"
		}
	}

	# Method: Set Out of Office
	[String] EnableOOF($InternalMessage,$ExternalMessage)
	{
		Try
		{
            $Splat = @{
                Identity = $this.SamAccountName
                AutoReplyState = 'enabled'
                ExternalAudience = 'all'
                InternalMessage = $InternalMessage
                ExternalMessage = $ExternalMessage
                ErrorAction = 'Stop'

            }
			Set-MailboxAutoReplyConfiguration @Splat -Confirm:$False 
			Return $Null
		}
		Catch
		{
			Throw "Unable to Set Out of Office : $($_.exception.message)"
		}
	}

    # Method: Hide From GAL
    [String] HideFromGAL()
    {
    	Try
		{
			Set-Mailbox -HiddenFromAddressListsEnabled $true -Identity $this.SAMAccountName -ErrorAction Stop
			Return $Null
		}
		Catch
		{
			Throw "Unable to Hide From GAL : $($_.exception.message)"
		}
    }

    # Method: Hide From GAL
    [String] SetAllowedSender($sender)
    {
    	Try
		{
			Set-Mailbox -Identity $this.SAMAccountName -AcceptMessagesOnlyFrom $Sender -ErrorAction Stop
			Return $Null
		}
		Catch
		{
			Throw "Unable to Hide From GAL : $($_.exception.message)"
		}
    }
}