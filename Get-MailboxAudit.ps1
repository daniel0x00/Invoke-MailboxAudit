function Get-MailboxAudit {

    <#
    .SYNOPSIS
        Returns the delegated permissions of a given mailbox in an Office 365 account. Required Admin account at tenant level.

    .PARAMETER Credential
        The secure credential object.
        
        PSCredential. Mandatory.
    #>   

    [CmdletBinding()]
    [OutputType([psobject])]
    param(    
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$false)]
        [System.Management.Automation.PSCredential] $Credential,

        [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]
        [string] $Identity,

        [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$false)]
        [Switch] $Proxy

        #TODO: receive a pipeline array to proccess only given accounts. 
    )

    begin {

        ## Remove all existing Powershell sessions 
        Get-PSSession | Remove-PSSession | Out-Null 

        ## Connect to O365:
        Write-Verbose "[Get-MailboxAudit] Connecting to the tenant using provided credentials..."

        if ($Proxy) { 
            Write-Verbose "[Get-MailboxAudit] You supplied a proxy switch. Make sure the proxy setting is configured in Internet Explorer configuration and that you can browser the web."

            $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -SessionOption $proxyOptions -ErrorAction Stop
        }
        else { $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -ErrorAction Stop }
        
        if ($null -ne $session) { 
            Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
            Write-Verbose "[Get-MailboxAudit] Connected."
        }
        else { Write-Error "[Get-MailboxAudit] Can't connect. Check if you're behind a proxy and if so, configure the proxy settings in Internet Explorer first." }
        
    }

    process {
        # Retreiving emails of the tenant:
        Write-Verbose "[Get-MailboxAudit] Retreiving target emails from the tenant..."
        
        $users = if ($Identity) { Get-Mailbox -Identity $Identity } else { Get-Mailbox -ResultSize unlimited | Where-Object {$_.PrimarySmtpAddress -notmatch '{'} }

        $users_count = ($users | Measure-Object).Count
        Write-Verbose "[Get-MailboxAudit] Users received. Quantity: $users_count."

        # Listing out Mailbox level permissions
        Write-Verbose "[Get-MailboxAudit] Listing out Mailbox level permissions..."
        $users | ForEach-Object {
            # Grab the user primary email:
            $user_primarysmtpaddress = $_.PrimarySmtpAddress
            $user_emaildomain = [string](([regex]::Match($user_primarysmtpaddress, "(?<domain>@(\w+\.)+(\w+)?$)")).groups["domain"].value)

            #
            # Grab the mailbox permissions:
            #
            $mailboxpermissions = Get-MailboxPermission -ResultSize unlimited -Identity $user_primarysmtpaddress | Where-Object {$_.isinherited -ne $true -and $_.user -notmatch 'SELF$' -and $_.isvalid -eq $true -and $_.deny -eq $false} 
            # Output:
            $mailboxpermissions | Select-Object @{n='User';e={$user_primarysmtpaddress}},@{n='GrantedUser';e={$_.user}},@{n='AccessType';e={'MailboxLevelReadAndManage'}},@{n='Permission';e={$_.accessrights}},@{n='Details';e={''}}

            #
            # Grab the SendAs permissions:
            #
            $recipientpermission = Get-RecipientPermission -ResultSize unlimited -Identity $user_primarysmtpaddress | Where-Object {$_.AccessControlType -eq 'allow' -and $_.trustee -notmatch 'SELF$'}
            # Output:
            $recipientpermission | Select-Object @{n='User';e={$user_primarysmtpaddress}},@{n='GrantedUser';e={$_.trustee}},@{n='AccessType';e={'MailboxLevelSendAs'}},@{n='Permission';e={$_.accessrights}},@{n='Details';e={''}}
        
            #
            # Grab the SendOnBehalf permissions:
            #
            if ($_.GrantSendOnBehalfTo -ne $null) {
                # Output:
                $_ | Select-Object @{n='User';e={$user_primarysmtpaddress}},@{n='GrantedUser';e={$_.GrantSendOnBehalfTo -join ', '}},@{n='AccessType';e={'MailboxLevelSendOnBehalf'}},@{n='Permission';e={'Granted'}},@{n='Details';e={''}}
            }

            #
            # Grab the Folder permissions:
            #
            ## Inbox:
            $folder = $user_primarysmtpaddress + ':\Inbox'
            $folderpermission = Get-MailboxFolderPermission -Identity $folder | Where-Object {$_.AccessRights -ne 'none' -and $_.IsValid -eq $true}
            # Output:
            $folderpermission | Select-Object @{n='User';e={$user_primarysmtpaddress}},@{n='GrantedUser';e={$_.User -join ', '}},@{n='AccessType';e={'FolderLevel:Inbox'}},@{n='Permission';e={$_.AccessRights}},@{n='Details';e={$_.SharingPermissionFlags -join ', '}}
            
            ## Calendar:
            $folder = $user_primarysmtpaddress + ':\Calendar'
            $folderpermission = Get-MailboxFolderPermission -Identity $folder | Where-Object {$_.AccessRights -ne 'none' -and $_.IsValid -eq $true -and $_.AccessRights -ne 'AvailabilityOnly'}
            # Output:
            $folderpermission | Select-Object @{n='User';e={$user_primarysmtpaddress}},@{n='GrantedUser';e={$_.User -join ', '}},@{n='AccessType';e={'FolderLevel:Calendar'}},@{n='Permission';e={$_.AccessRights}},@{n='Details';e={$_.SharingPermissionFlags -join ', '}}
            
            
            #
            # Grab the Folder forwarding rules outside the organization:
            #
            $forwardrules = Get-InboxRule -Mailbox $user_primarysmtpaddress -BypassScopeCheck | Where-Object {$_.enabled -eq $true -and $_.isvalid -eq $true -and ($_.ForwardAsAttachmentTo -ne $null -or $_.ForwardTo -ne $null) -and ($_.ForwardAsAttachmentTo -notmatch $user_emaildomain -or $_.ForwardTo -notmatch $user_emaildomain)}
            $forwardrules | Select-Object @{n='User';e={$user_primarysmtpaddress}},@{n='GrantedUser';e={$_.ForwardTo -join ',' -join $_.ForwardAsAttachmentTo}},@{n='AccessType';e={'ForwardRule'}},@{n='Permission';e={'Enabled'}},@{n='Details';e={$_.Description}}
            
        }
    }
    end { }
}