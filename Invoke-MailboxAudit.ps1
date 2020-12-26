#Requires -Version 5
function Invoke-MailboxAudit {
    <#
        .SYNOPSIS 
            Grab all Exchange permissions of a user. MailBox level, SendAs, SendOnBehalf, Folder (Top of Information Store, Inbox, Sent Items, Calendar, etc) and more.
            Tested against O365 Exchange Online.

        .PARAMETER Identity
            String. Required. Pipeline enabled. 
            The mailbox you want to audit. It accepts wildcards, like: a* 

        .NOTES
            Version:        4.7
            Author:         Daniel Ferreira
            Creation Date:  
                    v1.0:   5th  Nov. 2018
                    v2.0:   9th  Nov. 2018
                    v3.0:   10th Nov. 2018
                    v4.0:   14th Nov. 2018
                    v4.5:   16th Nov. 2018
                    v4.6:   23rd Nov. 2018
                    v4.7:   15th Dec. 2018
            Purpose:        Combine all Exchange permissions given by a user to a delegate in a single script. 
                            Possible work with Exchange Server 2010+, but untested and missing some parameters like -DomainController on cmdlets. 

        .EXAMPLE
            List all the permissions for 'user2' and 'user5' (folder permissions only in Inbox folder)

            PS C:\> $c = Get-Credential
            PS C:\> 'user2','user5' | Invoke-MailboxAudit -Credential $c -Verbose | Format-Table -AutoSize

        .EXAMPLE
            List all permissions but the Mailbox and Forwading rules for imported users, targeting folders Inbox, Calendar and Sent Items; regardless of the culture of the mailbox.
            In a MFA scenario, run this under a "Microsoft Exchange Online Remote PowerShell Module" console:

            PS C:\> $c = Get-Credential
            PS C:\> Import-Csv .\Users.csv | Invoke-MailboxAudit -Credential $c -MFA -Proxy -SkipMailboxPermission -SkipForwardingRules -Folder Inbox,Calendar,SentItems -Verbose
        
        .EXAMPLE
            List all the permissions for users with a mailbox that starts by: a*

            PS C:\> $c = Get-Credential
            PS C:\> Invoke-MailboxAudit -Credential $c -Identity a* -SkipMailboxPermission -Verbose 
        
        .EXAMPLE
            List all the permissions for all users in the tenant, for the Inbox and Sent Items folders, skipping all child user-created folders as well as forwarding rules:

            PS C:\> $c = Get-Credential
            PS C:\> 97..(97+25) | select @{n='Identity';e={[char]$_+'*'}} | Invoke-MailboxAudit -Credential $c -Proxy -SkipMailboxPermission -SkipUserCreatedFolder -SkipForwardingRule -SkipSendAsPermission -Folder Inbox,SentItems -Verbose 
        
        .OUTPUTS
            'user2','user5','mike' | Invoke-MailboxAudit -Credential $c -Verbose | Format-Table -AutoSize

            User                          GrantedUser                                    AccessType                                        Permission                            Details
            ----                          -----------                                    ----------                                        ----------                            -------
            user2@cditest.onmicrosoft.com Default                                        Folder:Inbox (Inbox)                              Owner
            user5@cditest.onmicrosoft.com Default                                        Folder:Top of Information Store                   ReadItems, FolderOwner, FolderVisible
            user5@cditest.onmicrosoft.com Default                                        Folder:subIBX - rare& \characters (User Created) DeleteOwnedItems
            user5@cditest.onmicrosoft.com "outside@domain.com" [SMTP:outside@domain.com] ForwardRule                                       Enabled                               If the message:...
            user5                         mike@cditest.onmicrosoft.com                   MailboxPermission                                 FullAccess
            user5                         user2@cditest.onmicrosoft.com                  MailboxPermission                                 FullAccess
            mike@cditest.onmicrosoft.com  user5, user2                                   SendOnBehalf                                      Granted
            mike                          user2@cditest.onmicrosoft.com                  MailboxSendAs                                     SendAs
            mike                          user5@cditest.onmicrosoft.com                  MailboxSendAs                                     SendAs
            mike@cditest.onmicrosoft.com  user2                                          Folder:Inbox (Inbox)                              Editor
        
    #>   
    [CmdletBinding()]
    [OutputType([psobject])]
    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias('mail','email','userprincipalname','user','mailbox')]
        [string] $Identity, 

        [Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$false)]
        [System.Management.Automation.PSCredential] $Credential,

        [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$false)]
        [switch] $Proxy,

        [Parameter(Position=3, Mandatory=$false, ValueFromPipeline=$false)]
        [switch] $MFA,

        [Parameter(Position=4, Mandatory=$false, ValueFromPipeline=$false)]
        [ValidateSet('Inbox','Calendar','SentItems','Tasks','Notes','Personal','Drafts','ConversationHistory','JunkEmail','Archive')]
        [string[]] $Folder='Inbox',

        [Parameter(Position=5, Mandatory=$false, ValueFromPipeline=$false)]
        [string] $Filter, # format: ((CustomAttribute1 -eq "X") -or (CustomAttribute1 -eq "Y"))

        [Parameter(Position=6, Mandatory=$false, ValueFromPipeline=$false)]
        [switch] $SkipMailboxPermission,

        [Parameter(Position=7, Mandatory=$false, ValueFromPipeline=$false)]
        [switch] $SkipSendAsPermission,

        [Parameter(Position=8, Mandatory=$false, ValueFromPipeline=$false)]
        [switch] $SkipFolderPermission,

        [Parameter(Position=9, Mandatory=$false, ValueFromPipeline=$false)]
        [switch] $SkipForwardingRule,

        [Parameter(Position=10, Mandatory=$false, ValueFromPipeline=$false)]
        [switch] $SkipUserCreatedFolder,

        [Parameter(Position=11, Mandatory=$false, ValueFromPipeline=$false)]
        [switch] $OutputErrors,

        [Parameter(Position=12, Mandatory=$false, ValueFromPipeline=$false)]
        [string] $OutputErrorsFile=$('Invoke-MailboxAuditErrors_'+ (Get-Date -Format d).Replace('/','-') +'.txt'),

        [Parameter(Position=13, Mandatory=$false, ValueFromPipeline=$false)]
        [string] $ConnectionUri='https://outlook.office365.com/powershell-liveid/', # For Exchange on-premises, use: http://<ServerFQDN>/PowerShell/

        [Parameter(Position=14, Mandatory=$false, ValueFromPipeline=$false)]
        [ValidateSet('Default', 'Basic', 'Negotiate', 'NegotiateWithImplicitCredential', 'Credssp', 'Digest', 'Kerberos')]
        [string] $Authentication='Basic',

        [Parameter(Position=15, Mandatory=$false, ValueFromPipeline=$false)]
        [switch] $LoadFunctions
    )

    begin {

        ##
        ### Global vars:
        ##
        $global:watcher = [Diagnostics.Stopwatch]::StartNew()
        $global:users = [System.Collections.ArrayList]@()
        $global:users_count = 0
        $global:session = $null

        # Fix filter parameter is there is any:
        if ($Filter) { $Filter = "-and $Filter" }

        ## Load write-error helper function:
        function Write-Error ($message) {
            [Console]::ForegroundColor = 'red'
            [Console]::BackgroundColor = 'black'
            [Console]::Error.WriteLine("$message")
            [Console]::ResetColor()
        }

        function getConnection { (Get-PSSession).where({$_.ConfigurationName -eq 'Microsoft.Exchange' -and $_.Availability -eq 'Available' -and $_.State -eq 'Opened'}) }
        
        function setConnection {
            # Verbose:
            Write-Verbose "[Invoke-MailboxAudit] Connecting to Exchange..."

            # Configure the proxy options:
            $proxyOptions = if ($Proxy) { New-PSSessionOption -ProxyAccessType IEConfig -IdleTimeout -1 } else { New-PSSessionOption -IdleTimeout -1 }

            # Configure the connection, depending of if -MFA is enabled or not:
            if ($MFA) { Connect-EXOPSSession -UserPrincipalName "$($credential.username)" -PSSessionOption $proxyOptions -ErrorAction Stop }
            else { 
                $localsession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $Credential -Authentication $Authentication -AllowRedirection -SessionOption $proxyOptions -ErrorAction Stop
                Set-Variable -Name session -Value $localsession -Scope Global
            }
            
            # Check if the user decided to load up the cmdlets from the PSRemoting session. Only useful for interactive sessions.
            if (($LoadFunctions) -and ($null -ne $global:session) -and (-not($MFA))) { Import-PSSession $global:session -DisableNameChecking -AllowClobber | Out-Null }
        }
        
        function testConnection {
            if (-not(getConnection)) { setConnection }
            if (getConnection) { Set-Variable -Name session -Value (getConnection) -Scope Global }
            else { throw 'Impossible to connect.' }
        }
        
        function getFolderPermission {
            [CmdletBinding()]
            [OutputType([psobject])]
            param(
                [Parameter(Mandatory = $true, Position = 0)]
                [string] $Identity,

                [Parameter(Mandatory = $false, Position = 1)]
                [string[]] $Folder='Top of Information Store',

                [Parameter(Mandatory = $false, Position = 2)]
                [switch] $Top,

                [Parameter(Mandatory = $false, Position = 3)]
                [string[]] $Exclude
            )
            process {
                ($Folder).foreach({
                    $proccessingfolder = $_

                    if (-not($Top)) { $folderQualifiedName = $Identity + ":\" + $proccessingfolder }
                    else { $folderQualifiedName = $Identity }

                    try {
                        #
                        # If -Top folder, means we're aiming root 'Top of information Store'
                        # Root folder can have different langages per culture.
                        if ($Top) {
                            # Verbose
                            Write-Verbose "[Invoke-MailboxAudit] Listing Folder permissions of $Identity, folder name: $proccessingfolder"

                            # Request and output the folder permission:
                            Invoke-Command -Session $global:session -ErrorVariable ErrorOut -ErrorAction silentlycontinue -ScriptBlock { Get-MailboxFolderPermission -Identity $using:folderQualifiedName } | Where-Object {$_.AccessRights -ne 'none' -and $_.IsValid -eq $true -and $_.user -notin $Exclude} | Select-Object @{n='User';e={$Identity}},@{n='GrantedUser';e={$_.User -join ', '}},@{n='AccessType';e={"Folder:$proccessingfolder"}},@{n='Permission';e={$_.AccessRights -join ', '}},@{n='Details';e={$_.SharingPermissionFlags -join ', '}}
                            
                            if ($ErrorOut) { Write-Error "[Invoke-MailboxAudit] Error requesting Top Folder permission of $Identity --> $ErrorOut"; $ErrorOut += " --> $Identity" }
                            if (($OutputErrors) -and ($ErrorOut)) { $ErrorOut | Out-File -FilePath $OutputErrorsFile -Append }
                        }
                        #
                        # If not -AllFolders and not -Top, means targeting a particular folder.
                        else {
                            
                            # Retrive the folder names. 
                            # NOTE: due to unability to use Get-MailboxFolder to list folders of a given user, we have to use Get-MailboxFolderStatistics. This cmdlet is very expensive as it has to calculate other statistics about the identity. Nevertheless by using PSRemoting the cost of retreiving the data to the client is mitigated a bit.
                            # Due to different cultures, the regular 'Inbox', 'Calendar', etc. name folders can be different. 
                            # The cmdlet Get-MailboxFolderPermission will be the one listing the appropiate names by utilizing the parameter -FolderScope
                            # As a consecuence, this cmdlet will also return all the child folders of the targeted folder and therefore all will be analized for permissions later on. 

                            # Output example for an account with one Inbox subfolder with special characters:
                                # FolderId       : LgAAAADOqRZpOnJQRKwst1x6iKvIAQBUzkROi9CYSrgM7380JH/DAAAAAAEMAAAB
                                # FolderPath     : /Inbox
                                # FolderType     : Inbox

                                # FolderId       : LgAAAADOqRZpOnJQRKwst1x6iKvIAQBUzkROi9CYSrgM7380JH/DAAAAGA2jAAAB
                                # FolderPath     : /subIBX - rare \characters
                                # FolderType     : User Created

                            # Verbose
                            Write-Verbose "[Invoke-MailboxAudit] Listing Folder names of $Identity, folder scope: $proccessingfolder"

                            $folders = Invoke-Command   -Session $global:session -ErrorVariable ErrorOut -ErrorAction silentlycontinue `
                                                        -ScriptBlock { Get-MailboxFolderStatistics -Identity $using:Identity -FolderScope $using:proccessingfolder | Select-Object FolderId, FolderPath, FolderType }
                            
                            # Check if we've recived folders:
                            if (-not($folders)) {
                                if ($ErrorOut) { Write-Error "[Invoke-MailboxAudit] Error requesting Folder Listing of $Identity --> $ErrorOut"; $ErrorOut += " --> $Identity" }
                                if (($OutputErrors) -and ($ErrorOut)) { $ErrorOut | Out-File -FilePath $OutputErrorsFile -Append }
                            }
                            else {
                                $folders | Where-Object {$_.FolderType -ne 'BirthdayCalendar' -and $_.FolderPath -notmatch 'holidays'} | ForEach-Object {
                                    # Create the FolderId:
                                    $folderid = $_.FolderId
                                    $folderpath = $_.FolderPath.Remove(0,1) + ' (' + $_.FolderType + ')'
                                    $folderQualifiedName = $Identity + ":" + $folderid

                                    # Check flag SkipUserCreatedFolder:
                                    if (($SkipUserCreatedFolder) -and ($_.FolderType -eq 'User Created')) { Write-Verbose "[Invoke-MailboxAudit] Skipping Folder permissions of $Identity, user-created folder name: $folderpath" }
                                    else {
                                        # Verbose
                                        Write-Verbose "[Invoke-MailboxAudit] Listing Folder permissions of $Identity, folder name: $folderpath"

                                        # Request and output the folder permission:
                                        Invoke-Command -Session $global:session -ErrorVariable ErrorOut -ErrorAction silentlycontinue -ScriptBlock { Get-MailboxFolderPermission -Identity $using:folderQualifiedName } | Where-Object {$_.AccessRights -ne 'none' -and $_.IsValid -eq $true -and $_.user -notin $Exclude} | Select-Object @{n='User';e={$Identity}},@{n='GrantedUser';e={$_.User -join ', '}},@{n='AccessType';e={"Folder:$folderpath"}},@{n='Permission';e={$_.AccessRights -join ', '}},@{n='Details';e={$_.SharingPermissionFlags -join ', '}} | Where-Object {$_.AccessType -notmatch 'Calendar' -and $_.Permission -ne 'AvailabilityOnly'}
                                    }
                                }
                            }
                        }
                    }
                    catch { 
                        Write-Error "[Invoke-MailboxAudit] Can't run getFolderPermission helper function on $Identity, folder $proccessingfolder. Exception: $_"
                        if ($OutputErrors) { $_ | Out-File -FilePath $OutputErrorsFile -Append }
                    }
                })
            }
        }

        ## Connect to Exchange:
        testConnection
    }

    process {
        $global:users.Clear()

        # Retreiving all mailboxes of the given identity:
        Write-Verbose "[Invoke-MailboxAudit] Retreiving mailboxes with identity $Identity"

        #
        # Retreiving all targeted mailboxes an 'OnBehalf' permission.
        # Get-Mailbox cmdlet do accept wildcards and also advanced filtering.
        # This is not an expensive command.
        # If aiming thousands of users, better to call Invoke-MailboxAudit by passing in the pipeline the letters of the alphabet in order, like: a*
        # 
        $mailboxfilter = [scriptblock]::create('((UserPrincipalName -like "'+$Identity+'") -and (IsMailboxEnabled -eq $true) -and (IsInactiveMailbox -eq $false)) '+$Filter)
        $mailboxes = Invoke-Command -Session $global:session -ErrorVariable ErrorOut -ErrorAction silentlycontinue -ScriptBlock { Get-Mailbox -Filter $using:mailboxfilter -ResultSize unlimited | Select-Object PrimarySmtpAddress,GrantSendOnBehalfTo,UserPrincipalName,DisplayName,Name } 

        # Calculate how many mailboxes did the Get-Mailbox cmdlet got. If the user piped in this cmdlet, then it will always be 1. 
        if ($null -ne $mailboxes) {
            $global:users_count = ($mailboxes | Measure-Object).count
            if ($global:users_count -eq 1) { $null = $global:users.Add($mailboxes) }
            elseif ($global:users_count -gt 1) { 
                $null = $global:users.AddRange($mailboxes)
                Write-Verbose "[Invoke-MailboxAudit] Retrieved mailboxes: $global:users_count"
            }
        }
        else { if (($OutputErrors) -and ($ErrorOut)) { $ErrorOut | Out-File -FilePath $OutputErrorsFile -Append } }


        #
        # Grab the SendOnBehalf permissions
        # This permissions comes with the execution of Get-Mailbox, so we just filter on who has the appropiate property configured
        # This is not an expensive permission.
        #
        Write-Verbose "[Invoke-MailboxAudit] Listing SendOnBehalf permissions of $Identity"
        # Query and output the permission
        ($global:users).where({$_.GrantSendOnBehalfTo -ne $null}) | Select-Object @{n='User';e={$_.PrimarySmtpAddress}},@{n='GrantedUser';e={$_.GrantSendOnBehalfTo -join ', '}},@{n='AccessType';e={'SendOnBehalf'}},@{n='Permission';e={'Granted'}},@{n='Details';e={''}}


        #
        # Retreiving 'SendAs' permission. This permission allows to totally impersonate a user (different than SendOnBehalf)
        # Get-RecipientPermission cmdlet do accept wildcards.
        # This is not an expensive command.
        # If aiming thousands of users, better to call Invoke-MailboxAudit by passing in the pipeline the letters of the alphabet in order, like: a*
        # 
        if (-not($SkipSendAsPermission)) {
            # Verbose message:
            Write-Verbose "[Invoke-MailboxAudit] Listing SendAs permissions of $Identity"

            # Request and output the permissions:
            Invoke-Command -Session $global:session -ErrorVariable ErrorOut -ErrorAction silentlycontinue -ScriptBlock { Get-RecipientPermission -ResultSize unlimited -Identity $using:Identity } | Where-Object {$_.AccessControlType -eq 'allow' -and $_.trustee -notmatch 'SELF$'} | Select-Object @{n='User';e={$_.identity}},@{n='GrantedUser';e={$_.trustee -join ', '}},@{n='AccessType';e={'MailboxSendAs'}},@{n='Permission';e={$_.accessrights -join ', '}},@{n='Details';e={''}}
            if ($ErrorOut) { Write-Error "[Invoke-MailboxAudit] Error requesting SendAs permissions of $Identity --> $ErrorOut"; $ErrorOut += " --> $Identity" }
            if (($OutputErrors) -and ($ErrorOut)) { $ErrorOut | Out-File -FilePath $OutputErrorsFile -Append }
        }


        #
        # Grab the Mailbox permissions:
        # Get-MailboxPermission cmdlet do accept wildcards, but is not fast at all. 
        # Mailbox Permissions are configured at tenant level in the O365 EXO Admin Portal (in case of O365). The user is not aware of this permissions (permission not visible in Outlook)
        #
        # WARNING : Get-MailboxPermission is a VERY EXPENSIVE cmdlet even when using it with wildcards. It is much recommended to skip it if you're targeting all your mailboxes, and rather running Invoke-MailboxAudit targeting only Mailbox Permissions against your most critical mailboxes.
        # To skip Get-MailboxPermission cmdlet in Invoke-MailboxAudit, use the switch -SkipMailboxPermission.
        #
        # This is an expensive operation while targeting thousands of users, as by default will return even domain group and self access to the inbox. 
        # If aiming thousands of users, better to run Invoke-MailboxAudit exclusively for this permission.
        #
        if (-not($SkipMailboxPermission)) { 
            Write-Verbose "[Invoke-MailboxAudit] Listing Mailbox permissions of $Identity"

            # Request and output the permissions:
            Invoke-Command -Session $global:session -ErrorVariable ErrorOut -ErrorAction silentlycontinue -ScriptBlock { Get-MailboxPermission -ResultSize unlimited -Identity $using:Identity } | Where-Object {$_.isinherited -eq $false -and $_.user -notmatch 'SELF$' -and $_.isvalid -eq $true -and $_.deny -eq $false -and $_.Identity -notmatch 'DiscoverySearchMailbox' -and $_.Identity -notmatch 'AggregateGroupMailbox'} | Select-Object @{n='User';e={$_.Identity}},@{n='GrantedUser';e={$_.user -join ', '}},@{n='AccessType';e={'MailboxPermission'}},@{n='Permission';e={$_.accessrights -join ', '}},@{n='Details';e={''}}
            if ($ErrorOut) { Write-Error "[Invoke-MailboxAudit] Error requesting Mailbox permissions of $Identity --> $ErrorOut"; $ErrorOut += " --> $Identity" }
            if (($OutputErrors) -and ($ErrorOut)) { $ErrorOut | Out-File -FilePath $OutputErrorsFile -Append }
        }

        
        ## 
        ## Test connection to Exchange.
        ## Checking the connection is needed in case of dealing with large lists of mailboxes.
        ## The reason to check the connection at this point is because if there is a check for Mailbox permissions with a long list of users (e.g.: a*), the connection can be lost at this point
        ##
        testConnection


        ##
        ### Initiating a loop for each of the identities.
        ### This is needed because the below Exchange/EXO do not support wildcards.
        ##

        if (-not(($SkipFolderPermission) -and ($SkipForwardingRule))) { # We only iterate through mailboxes if there is a real need: only when folder permissions and/or forwarding rules are required.
            $i = 0 # for counting the iterations.
            ($global:users).where({$_.PrimarySmtpAddress.Address -notmatch 'DiscoverySearchMailbox' -and $_.PrimarySmtpAddress.Address -notmatch 'AggregateGroupMailbox'}).foreach({
                $user_primarysmtpaddress = $_.PrimarySmtpAddress.Address 
                $user_displayname = $_.DisplayName
                $user_upn = $_.UserPrincipalName
                $user_name = $_.Name
                $user_emaildomain = [string](([regex]::Match($user_primarysmtpaddress, "(?<domain>@(\w+\.)+(\w+)?$)")).groups["domain"].value)
                $i++

                # Verbose
                Write-Verbose "[Invoke-MailboxAudit] Iterating through mailbox $i of $global:users_count..."
    
                #
                # Grab the Folder permissions.
                # Get-MailboxFolderPermission cmdlet do not accept wildcards and also comes with Culture names complicationes.
                # This a very expensive command that has to be combinated with Get-MailboxFolderStatistics to discover folder names. 
                #
    
                if (-not($SkipFolderPermission)) {
                    ## Top of Information Store:
                    getFolderPermission -Identity $user_primarysmtpaddress -Top -Exclude $user_displayname,$user_upn,$user_name
    
                    ## Others:
                    getFolderPermission -Identity $user_primarysmtpaddress -Folder $Folder -Exclude $user_displayname,$user_upn,$user_name
                }
    
    
                #
                # Grab the Forwarding rules.
                # Get-InboxRule cmdlet do not accept wildcards. It produces a big output, but that is mitigated by using PSRemoting and selecting the needed fields.
                # This a medium expensive command, as not all mailboxes will have forwarding rules. 
                #
                if (-not($SkipForwardingRule)) {
                    # Verbose message:
                    Write-Verbose "[Invoke-MailboxAudit] Listing Forwarding Rules of $user_primarysmtpaddress"
    
                    # Request and output the rules:
                    Invoke-Command -Session $global:session -ErrorVariable ErrorOut -ErrorAction silentlycontinue -ScriptBlock { Get-InboxRule -Mailbox $using:user_primarysmtpaddress | Select-Object Enabled,IsValid,ForwardAsAttachmentTo,ForwardTo,Description } | Where-Object {$_.enabled -eq $true -and $_.isvalid -eq $true -and ($_.ForwardAsAttachmentTo -ne $null -or $_.ForwardTo -ne $null) -and ($_.ForwardAsAttachmentTo -notmatch $user_emaildomain -or $_.ForwardTo -notmatch $user_emaildomain)} | Select-Object @{n='User';e={$user_primarysmtpaddress}},@{n='GrantedUser';e={$_.ForwardTo -join ',' -join $_.ForwardAsAttachmentTo}},@{n='AccessType';e={'ForwardRule'}},@{n='Permission';e={'Enabled'}},@{n='Details';e={$_.Description}} 
                    if ($ErrorOut) { Write-Error "[Invoke-MailboxAudit] Error requesting Forwarding rules of $user_primarysmtpaddress --> $ErrorOut"; $ErrorOut += " --> $user_primarysmtpaddress" }
                    if (($OutputErrors) -and ($ErrorOut)) { $ErrorOut | Out-File -FilePath $OutputErrorsFile -Append }
                }

                ## 
                ## Test connection to Exchange.
                ## Checking the connection is needed in case of dealing with large lists of mailboxes.
                ## This check is at the folder permission level, means will catch when there is a long list of user-created folders and the connection is lost while checking on some folder. 
                ##
                testConnection
            })
        }
    }
    
    end { 
        $global:watcher.Stop()
        $watcherResults = $global:watcher.Elapsed | Select-Object days,hours,minutes,seconds | Format-Table -AutoSize
        Write-Verbose "[Invoke-MailboxAudit] Completed for $Identity"
        Write-Verbose ($watcherResults | Out-String)
    }
}