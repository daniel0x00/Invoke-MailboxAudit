function Get-MailboxAudit {
    <#

    #>   
    [CmdletBinding()]
    [OutputType([psobject])]
    param(
        [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]
        [Alias("mail")] 
        [string] $Identity,

        [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$false)]
        [System.Management.Automation.PSCredential] $Credential,

        [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$false)]
        [Switch] $Proxy,

        [Parameter(Position=3, Mandatory=$false, ValueFromPipeline=$false)]
        [Switch] $MFA
    )

    begin {

        ## Connect to O365:
        Write-Verbose "[Get-MailboxAudit] Connecting to the tenant using provided credentials..."

        if ($MFA) {
            $tenantadmincommand = Get-Command -Name Get-MailboxPermission -EA 0
            if (-not($tenantadmincommand)) { Write-Error 'ERROR: you specified to be MFA. You have to run the coomand "$proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig; Connect-EXOPSSession -UserPrincipalName <upn> -PSSessionOption $proxyOptions", under a "Microsoft Exchange Online Remote PowerShell Module" console.' -ErrorAction Stop }
            else { Write-Verbose "[Get-MailboxAudit] Connected." }
        }
        else {
            if ($Proxy) { 
                Write-Verbose "[Get-MailboxAudit] You supplied a proxy switch. Make sure the proxy setting is configured in Internet Explorer configuration and that you can browser the web."
    
                $proxyOptions = New-PSSessionOption -ProxyAccessType IEConfig
                $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -SessionOption $proxyOptions -ErrorAction Stop
            }
            else { $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -ErrorAction Stop }
            
            if ($null -ne $session) { 
                Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
                Write-Verbose "[Get-MailboxAudit] Connected."
    
                $tenantadmincommand = Get-Command -Name Get-MailboxPermission -EA 0
                if (-not($tenantadmincommand)) { Write-Error "ERROR: couldn't find all cmdlets from the PSSession connection. Normally due to not enough rights at the tenant. Try a different account." -ErrorAction Stop }
            }
            else { Write-Error "[Get-MailboxAudit] Can't connect. Check if you're behind a proxy and if so, configure the proxy settings in Internet Explorer first." }
        }
        
    }

    process {
        # Retreiving emails of the tenant:
        if ($Identity) { Write-Verbose "[Get-MailboxAudit] Listing permissions for user: $Identity" }
        else { Write-Verbose "[Get-MailboxAudit] Retreiving target emails from the tenant..." }
        
        $users = if ($Identity) { Get-Mailbox -Identity $Identity -ErrorAction silentlycontinue } else { Get-Mailbox -ResultSize unlimited | Where-Object {$_.PrimarySmtpAddress -notmatch '{'} }

        # Listing out Mailbox level permissions
        $users | ForEach-Object {
            # Grab the user primary email:
            $user_primarysmtpaddress = $_.PrimarySmtpAddress
            $user_emaildomain = [string](([regex]::Match($user_primarysmtpaddress, "(?<domain>@(\w+\.)+(\w+)?$)")).groups["domain"].value)

            #
            # Grab the Mailbox permissions:
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
    end { 
        Get-PSSession | Remove-PSSession | Out-Null 
        Write-Verbose "[Get-MailboxAudit] Completed."
    }
}