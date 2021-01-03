# Invoke-MailboxAudit

Grab all Exchange permissions of a user, bulk of users or all users on the tenant. 

Tested against O365 Exchange Online.

## Supported permissions:

- MailBox level (assigned by tenant/mailbox administrator)
- SendAs
- SendOnBehalf
- Folder (Top of Information Store, Inbox, Sent Items, Calendar, etc). These are user-assigned permissions. 

You can use this module to list out all mailboxes where users gave Read (`Owner`, `FullAccess`, etc) permissions to `Everyone` or similar roles, thereby exposing their mailbox to other members in the organization.
           
## Usage

1. Use `Windows PowerShell 5.1`.
2. Install the module by invoking it or dot-sourcing it:
```powershell
iex((iwr https://raw.githubusercontent.com/daniel0x00/Invoke-MailboxAudit/master/Invoke-MailboxAudit.ps1 -UseBasicParsing).content)
```
3. Run the cmdlet as shown below.

**MFA support**:
Does your admin account use multi-factor authentication? 
Then load this script under a ["Microsoft Exchange Online Remote PowerShell Module"](https://docs.microsoft.com/en-us/powershell/exchange/v1-module-mfa-connect-to-exo-powershell?view=exchange-ps) special Windows PowerShell console and use the `-MFA` switch when using the cmdlet.


### List all the permissions for 'user2', 'user5', 'mike'

```powershell
PS C:\> $c = Get-Credential
PS C:\> 'user2','user5','mike' | Invoke-MailboxAudit -Credential $c -Verbose | Format-Table -AutoSize
```
```console
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
```

### List permissions for given users at bulk using a MFA admin account and under a proxied network

Note: The `-Proxy` switch forces the cmdlet to use the default proxy settings configured on the computer where the cmdlet runs.

```powershell
PS C:\> $c = Get-Credential
PS C:\> Import-Csv .\Users.csv | Invoke-MailboxAudit -Credential $c -MFA -Proxy -SkipMailboxPermission -SkipForwardingRules -Folder Inbox,Calendar,SentItems -Verbose
```

### List all the permissions for users with a mailbox that starts by: a*

```powershell
PS C:\> $c = Get-Credential
PS C:\> Invoke-MailboxAudit -Credential $c -Identity a* -SkipMailboxPermission -Verbose 
```

### List all the permissions for all users in the tenant, for the Inbox and Sent Items folders, skipping all child user-created folders as well as forwarding rules

```powershell
PS C:\> $c = Get-Credential
PS C:\> 97..(97+25) | select @{n='Identity';e={[char]$_+'*'}} | Invoke-MailboxAudit -Credential $c -Proxy -SkipMailboxPermission -SkipUserCreatedFolder -SkipForwardingRule -SkipSendAsPermission -Folder Inbox,SentItems -Verbose 
```

