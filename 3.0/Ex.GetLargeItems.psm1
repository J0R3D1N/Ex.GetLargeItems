<#
. [COPYRIGHT]
. © 2011-2019 Microsoft Corporation. All rights reserved. 
.
. [DISCLAIMER]
. This sample script is not supported under any Microsoft standard support
. program or service. The sample scripts are provided AS IS without warranty of
. any kind. Microsoft disclaims all implied warranties including, without
. limitation, any implied warranties of merchantability or of fitness for a
. particular purpose. The entire risk arising out of the use or performance of
. the sample scripts and documentation remains with you. In no event shall
. Microsoft, its authors, or anyone else involved in the creation, production,
. or delivery of the scripts be liable for any damages whatsoever (including,
. without limitation, damages for loss of business profits, business
. interruption, loss of business information, or other pecuniary loss) arising
. out of the use of or inability to use the sample scripts or documentation,
. even if Microsoft has been advised of the possibility of such damages.
.
. [AUTHOR]
. Jason Parker, Sr. Consultant
. 
. [CONTRIBUTORS]
. Michael Hall, Service Engineer
. Jayme Bowers, Senior Service Engineer
. Stuart Murray, Consultant
. Dmitry Kazantsev, Senior Consultant
. 
. [Module]
. Ex.GetMailboxItems
.
. [VERSION]
. 3.0.0
.
. [VERSION HISTORY / UPDATES]
. Removed all prevsion version notes.  Version history can be found in the archive.
.
. 2.3 - MAJOR UPDATE
. Jason Parker - Updated all function with proper verb / noun association
. Jason Parker - Fixed logic in Get-FolderItems where the script would get stuck in an endless loop
. Jason Parker - Updated Get-FolderItems so that when a set of itmes is returned, the whole array is evaluated for exceeding the ItemSizeLimit, which increases performance
. Jason Parker - Updated Get-FolderItems so that the search only returns a small subset of properties on the items instead of every possible property.
. Jason Parker - Removed Large Item Notice - will bring back in 2.3.1
. Jason Parker - Added functions for better menu and dialog feedback
. Jason Parker - Updated progress functionality
.
. 3.0 - NEW RELEASE
. Jason Parker - Converted LargeItemScript from a PS1 to a Module
. Jason Parker - Removed support for Exchange 2010
. Jason Parker - Added better support for EWS API detection
. Jason Parker - Fixed varibles to support module based format
#>

Function Get-MailboxItems {
    <#
    .SYNOPSIS
    This function will run a series of cmdlets / functions using Exchange Web Services to search mailboxes for items over a specified size. Useful for Office 365 on-boarding where you need to remediate large items before migration.

    .DESCRIPTION
    This function arranges building-block cmdlets / functions to connect to an Exchange environment and loops through all or a subset of mailboxes with an impersonator account using Exchange Web Services API.  The impersonator account will enumerate every item in every folder and identify items that are exceeding a specific size. The function is designed to be executed before on-boarding an Organization to Office 365.

    .PARAMETER ServiceAccountName
    Specifies the UserPrincipalName of the user which has elevated permissions (impersonation and mailbox export).

    .PARAMETER ServicePassword
    Specifies the password for the Service Account Users (stored in clear text).

    .PARAMETER ItemSizeLimit
    Sets the value from which you will measure items against (in MB).  This value should be set to the same value used in your Office 365 Tenant (Max Send / Receive Size).  The maximum size allowed in Office 365 is 150 MB.

    .PARAMETER MailboxLocation
    MailboxLocation is a [ValidateSet] parameter and will only accept 3 possible values:

    Primary:
    The function will target the users primary mailbox and folders (MsgFolderRoot)

    Archive:
    The function will target the users archive mailbox and folders (ArchiveMsgFolderRoot)

    RecoverableItems:
    The function will target the hidden mailbox dumpter and folders (RecoverableItemsRoot), which is useful for mailboxes under litigation hold.

    .PARAMETER Action
    Action is a [ValidateSet] parameter and will only accept 4 possible values:

    ReportOnly
    -----------
    No action is taken on any item, only data collection which is output into a CSV file

    MoveLargeItems
    ---------------
    Function will prompt for a folder name, get / create the folder (regardless of MailboxLocation parameter, all folders will be created in the users Primary mailbox), and MOVE any item found larger than the ItemSizeLimit into that folder. NOTE:  If moving large items from locations other than the users Primary mailbox, their mailbox must be able to support these new items!

    MoveAndExportItems
    -------------------
    Function will prompt for a folder name, get / create the folder (regardless of MailboxLocation parameter, all folders will be created in the users Primary mailbox), MOVE any item found larger than the ItemSizeLimit into that folde, and finally create a New-MailboxExportRequest which will only export items from the Large Item folder.

    ExportLargeItems
    -----------------
    Does NOT check any items and will ONLY prompt for the folder name used from a previous MoveLargeItems operation and and finally create a New-MailboxExportRequest which will only export items from the Large Item folder.

    -------------------------- PST EXPORT LOCATIONS --------------------------

    When selecting an action with Export, the function will prompt where the PST Export files should be stored.  The function will allow 2 choices:

    Home Directory
    ---------------
    Will read the homedirectory attribute from Active Directory and attempt to export to that location

    Centralized Network Share
    --------------------------
    Function will prompt for the Server and Share Name and allows the operator to select the folder to be used

    .PARAMETER Uri
    Sets the Uri for the Exchange Web Services endpoint.  Useful when you can't leverage Autodiscover or Autodiscover fails.

    .EXAMPLE
    Get-MailboxItems -ServiceAccountName <User@domain.com> -ServicePassword <Password> -ItemSizeLimit <Value in MB> -MailboxLocation Primary -ScriptAction ReportOnly

    -- CREATES CSV REPORT OF MAILBOXES WITH LARGE ITEMS --

    In this example, the -MailboxLocation is set to Primary and the -ScriptAction has been set to ReportOnly which will create a CSV file containing all the item violations from all the mailboxes that were scanned. This CSV Report can be used in a subsequent execution where an alternate -ScriptAction is used (e.g. MoveLargeItems or MoveAndExportItems).  This may save considerable execution time if the function is run multiple times.

    .EXAMPLE
    Get-MailboxItems -ServiceAccountName <User@domain.com> -ServicePassword <Password> -ItemSizeLimit <Value in MB> -MailboxLocation Archive -ScriptAction ReportOnly

    -- CREATES CSV REPORT OF ARCHIVE MAILBOXES WITH LARGE ITEMS --

    In this example, the -MailboxLocation is set to Archive and the -ScriptAction has been set to ReportOnly which will create a CSV file containing all the item violations from all the mailboxes that were scanned. This CSV Report can be used in a subsequent execution where an alternate -ScriptAction is used (e.g. MoveLargeItems or MoveAndExportItems).  This may save considerable execution time if the function is run multiple times.

    .EXAMPLE
    Get-MailboxItems -ServiceAccountName <User@domain.com> -ServicePassword <Password> -ItemSizeLimit <Value in MB> -MailboxLocation RecoverableItems -ScriptAction ReportOnly

    -- CREATE CSV REPORT OF MAILBOXES WITH LARGE ITEMS (LEGAL / LITIGATION HOLD) --

    In this example, the -MailboxLocation is set to RecoverableItems and the -ScriptAction has been set to ReportOnly which will create a CSV file containing all the item violations from all the mailboxes that were scanned. This CSV Report can be used in a subsequent execution where an alternate -ScriptAction is used (e.g. MoveLargeItems or MoveAndExportItems).  This may save considerable execution time if the function is run multiple times.

    .EXAMPLE
    Get-MailboxItems -ServiceAccountName <User@domain.com> -ServicePassword <Password> -ItemSizeLimit <Value in MB> -ScriptAction ExportLargeItems

    -- EXPORT LARGE ITEMS ONLY --

    In this example, the function is not using the -MailboxLocation parameter because the mailbox will not be searched using EWS.  This ScriptAction will only attempt to create a New-MailboxExportRequest and export the contents of the Large Item folder to a PST.

    .NOTES
    Large environments will take a significant amount of time to process (hours/days). You can reduce the run time by either using a CSV import file with a smaller subset of users or running multiple instances of the function concurrently, targeting mailboxes on different servers.  Running multiple instances assumes your Exchange Web Services endpoint is behind a network load balancer.

    Important: Do not run too many instances or against too many mailboxes at once. Doing so could cause performance issues, affecting users.  Microsoft is not responsible for any such performance issue or improper use and planning.

    [PERMISSIONS REQUIRED]
    This function requires elevated permissions beyond the typical RBAC roles.

    [EXCHANGE 2013/2016/2019 PERMISSIONS]
    There are two sets of permissions required to properly execute the function in an Exchange 2013 / 2016 / 2019 environment.  Impersonation and Export permissions. Both sets of permissions will require changing or creating of RBAC Management Role Assignments.

    [IMPERSONATION PERMISSIONS]
    From the Exchange Management Shell, run the New-ManagementRoleAssignment cmdlet to add the permission to impersonate to the specified user:
    New-ManagementRoleAssignment –Name:impersonationAssignmentName –Role:ApplicationImpersonation –User:ServiceAccount

    [NEW-MAILBOXEXPORTREQUEST PERMISSIONS]
    This cmdlet is available only in the Mailbox Import Export role, and by  default, that role isn't assigned to a role group. To use this cmdlet, you need to add the Mailbox Import Export role to a role group (for example, to the Organization Management role group). For more information, see the "Add a role to a role group" section in Manage role groups.
    New-ManagementRoleAssignment –Role “Mailbox Import Export” –User Domain\User

    When specifying an Export Action, please ensure that the network location has NTFS Read/Write permissions for the "Exchange Trusted Subsystem" Group

    .LINK
    Install the EWS Managed API 2.2:  http://www.microsoft.com/en-us/download/details.aspx?id=42951

    .LINK
    Configure Exchange Web Services Impersonation:  https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2007/bb204095(v=exchg.80)

    .LINK
    Exchange 2013 / 2016 / 2019 Manage Role Groups:  https://technet.microsoft.com/en-us/library/jj657480(v=exchg.160).aspx
    #>

    Param (       
        [Parameter(Position = 1, Mandatory = $True, HelpMessage = "Please provide the UserID for the Service Account")]
        [System.String]$ServiceAccountName,

        [Parameter(Position = 2, Mandatory = $True, HelpMessage = "Please provide the password for the Service Account")]
        [System.String]$ServicePassword,

        [Parameter(Position = 3, Mandatory = $True, ValueFromPipeline = $True, HelpMessage = "Enter the item size in Megabytes you want to search for in each mailbox")]
        [ValidateRange(1, 150)]
        [System.Int32]$ItemSizeLimit,

        [Parameter(Position = 4, Mandatory = $false)]
        [ValidateSet("Primary", "Archive", "RecoverableItems")]
        [System.String]$MailboxLocation,

        [Parameter(Position = 5, Mandatory = $true)]
        [ValidateSet("ReportOnly", "MoveLargeItems", "MoveAndExportItems", "ExportLargeItems")]
        [System.String]$Action,
        
        [Parameter(Mandatory = $False)]
        [System.URI]$Uri
    )

    #region MessageBanners
    $HomeDirWarning = (@"

Exporting to a User's Home Directory will require that the account used to
execute this function has sufficient permissions for the operation!

"@)

    $CentralizedExportWarning = (@"

Centralized Exports will require enough free space to accomodate all the
PST files for Users checked using this function.  For performance purposes,
do not select a Network location, but use a server with a good amount of
local storage.

"@)

    $ExportOnlyWarning = (@"

SELECTING THIS OPTION SHOULD ONLY BE DONE IF DURING A PREVIOUS OPERATION
LARGE ITEMS WERE MOVED TO A LARGE ITEM FOLDER.  YOU MUST USE THE *SAME*
LARGE ITEM FOLDER NAME FROM THE PREVIOUS OPERATION OR YOU WILL NOT EXPORT
THE CORRECT DATA!

THIS PROCESS EXPORTS FROM THE LARGE ITEM FOLDER ONLY!

"@)

    $GetMailboxWarning = (@"

By selecting [A]LL MAILBOXES, the function will use the Cmdlet Get-Mailbox
with the -ResultSize Unlimited parameter.  Depending on the size of your
organization, this may be a long operation.

"@)

    #endregion

    ##############################################################################
    # Main Script
    ##############################################################################
    $Error.Clear()
    Write-Debug "Script Start"
    $Stopwatch = New-Object System.Diagnostics.Stopwatch
    $Stopwatch.Start()
    $Script:LargeFolderName = $null
    $Script:CurrentPSSession = $null
    $Script:AllLargeItems = $null
    $Script:CentralizedExport = $null
    $Script:ExportPath = $null
    $Script:MailboxObjects = $null

    $myTitle = @"
\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
\\\\\\\\\\
\\\\\\\\\\  Title:    Office 365 Get-MailboxItems
\\\\\\\\\\  Purpose:  Find items in mailboxes over $($ItemSizeLimit) MB and performs an action(s)
\\\\\\\\\\  Actions:  ReportOnly, MoveLargeItems, MoveAndExportItems, ExportLargeItems
\\\\\\\\\\  Script:   Get-MailboxItems
\\\\\\\\\\
\\\\\\\\\\  Help:    Get-Help Get-MailboxItems -Full
\\\\\\\\\\
\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
"@

    Show-Menu -Title $myTitle -ClearScreen -DisplayOnly -Style None -Color White

    If (-NOT ((Test-ConsoleVersion -eq $true) -and (Test-ConsoleRights -eq $true))) {
        Show-Menu -Title "This Console does NOT meet the minimum requirements!" -DisplayOnly -Style Full -Color Yellow
        Write-Host " >> PLEASE USE POWERSHELL V3 AND RUN AS ADMINISTRATOR <<`n`r" -ForegroundColor Yellow
        Read-Host "Press any key to exit" | Out-Null
        Clear-Host
        Return
    }

    try {
        $ScriptPath = (Get-Location).Path
        $CurrentDateTime = Get-Date -Format yyyyMMdd_HHmmss
        $LogFile = "$ScriptPath\LargeItemChecks_ScriptLog_$CurrentDateTime.log"
    
        #region Validate / Connect to Exchange
        Write-Log -Type INFO -Text "Looking for Microsoft Exchange Server Management Tools"
        $RootRegPath = 'HKLM:\SOFTWARE\Microsoft'
        
        If (Test-Path -Path $RootRegPath'\ExchangeServer\v15\AdminTools') {
            [System.String]$ExchangeVersion = "E15"
            $env:ExchangeInstallPath = (Get-ItemProperty $RootRegPath'\ExchangeServer\v15\Setup').MsiInstallPath
            Write-Log -Type INFO -Text ("Setting Exchange Version to: {0}" -f $ExchangeVersion)
        } Else {
            Write-Log -Type ERROR -Text ("Microsoft Exchange Server Management Tools cannot be found or are not installed, function will now exit")
            Return
        }
        
        Write-Log -Type INFO -Text ("Checking PowerShell Console for Exchange Management Cmdlets")
        If (-NOT (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
            Write-Log -Type INFO -Text ("Attempting to load Exchange Management Shell based on version")
            If (Test-Path $env:ExchangeInstallPath'bin\RemoteExchange.ps1') { 
                . $env:ExchangeInstallPath'bin\RemoteExchange.ps1'
                Connect-ExchangeServer -auto
                $Script:CurrentPSSession = Get-PSSession -InstanceId (Get-OrganizationConfig).RunspaceId.Guid
            } Else {
                Write-Log -Type INFO -Text ("Microsoft Exchange Server Management Shell could not be loaded, function will now exit")
                Return
            }
        } Else {
            $Script:CurrentPSSession = Get-PSSession -InstanceId (Get-OrganizationConfig).RunspaceId.Guid
            Write-Log -Type INFO -Text ("Found Exchange Management Cmdlets | Connected to: {0}" -f $CurrentPSSession.ComputerName)
        }
        #endregion
        #region Validating ACTION based parameters
        Write-Log -Type INFO -Text ("Validating ACTION based parameters")
        Switch ($Action) {
            "ReportOnly" {
                Write-Log -Type INFO -Text ("ACTION: {0}" -f $Action)
                If ([System.String]::IsNullOrEmpty($MailboxLocation)) { $Script:MailboxLocation = "Primary" }
                [System.Collections.ArrayList]$Script:AllLargeItems = @()
            }
            "MoveLargeItems" {
                Write-Log -Type INFO -Text ("ACTION: {0}" -f $Action)
                Write-Log -Type INFO -Text ("Gathering required variables...")
                If ([System.String]::IsNullOrEmpty($MailboxLocation)) { $Script:MailboxLocation = "Primary" }
                Get-LargeFolderName
                [System.Collections.ArrayList]$Script:AllLargeItems = @()
            }
            "MoveAndExportItems" {
                Write-Log -Type INFO -Text ("ACTION: {0}" -f $Action)
                Write-Log -Type INFO -Text ("Gathering required variables...")
                If ([System.String]::IsNullOrEmpty($MailboxLocation)) { $Script:MailboxLocation = "Primary" }
                Get-LargeFolderName
                $optHomeDirs = New-Object System.Management.Automation.Host.ChoiceDescription "&User Home Directory", "Query Active Directory for the User's Home Directory attribute and Export the PST to that location"
                $optCentralExport = New-Object System.Management.Automation.Host.ChoiceDescription "&Centralized Location", "Select a folder location to store all PST Export files"
                $Options = [System.Management.Automation.Host.ChoiceDescription[]]($optHomeDirs, $optCentralExport)
                Switch ($Host.UI.PromptForChoice("`n >> PST EXPORT LOCATION <<", "`n Which type of location do you want to store the PST Export Files?`n`r", $Options, 1)) {
                    0 {
                        Show-Menu -Title "WARNING: HOME DIRECTORY EXPORT SELECTED!" -ClearScreen -DisplayOnly -Style Info -Color Yellow
                        Write-Host $HomeDirWarning -ForegroundColor Yellow
                        If ((Get-ChoicePrompt -OptionList "&OK", "&Cancel" -Title " OK to Continue, Cancel to exit`n`r" -Message $null -default 1) -eq 1) {
                            Write-Log -Type INFO -Text "// HOME DIRECTORY SELECTION | User Cancelled the operation!"
                            Return
                        } Else {
                            $Script:CentralizedExport = $False
                            $Script:ExportPath = "Home Directories"
                        }
                    }
                    1 {
                        Show-Menu -Title "WARNING: CENTRALIZED EXPORT SELECTED!" -ClearScreen -DisplayOnly -Style Info -Color Yellow
                        Write-Host $CentralizedExportWarning -ForegroundColor Yellow
                        If ((Get-ChoicePrompt -OptionList "&OK", "&Cancel" -Title " OK to Continue, Cancel to exit`n`r" -Message $null -default 1) -eq 1) {
                            Write-Log -Type INFO -Text "// CENTRALIZED EXPORT SELECTION | User Cancelled the operation!"
                            Return
                        } Else {
                            $Script:CentralizedExport = $True
                            $NetworkSharePath = Get-NetworkSharePath
                            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
                            $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                            $FolderDialog.SelectedPath = $NetworkSharePath
                            Do { $DialogResult = $FolderDialog.ShowDialog() }
                            Until ($DialogResult -eq "OK")
                            $Script:ExportPath = $FolderDialog.SelectedPath
                        }
                    }
                }
                [System.Collections.ArrayList]$Script:AllLargeItems = @()
            }
            "ExportLargeItems" {
                Write-Log -Type INFO -Text ("ACTION: {0}" -f $Action)
                Write-Log -Type INFO -Text ("Gathering required variables...")

                Show-Menu -Title "WARNING: EXPORT LARGE ITEMS SELECTED!" -ClearScreen -DisplayOnly -Style Info -Color Yellow
                Write-Host $ExportOnlyWarning -ForegroundColor Yellow
                If ((Get-ChoicePrompt -OptionList "&OK", "&Cancel" -Title " OK to Continue, Cancel to exit`n`r" -Message $null -default 1) -eq 1) {
                    Write-Log -Type INFO -Text "// EXPORT LARGE ITEMS SELECTION | User Cancelled the operation!"
                    Return
                }

                Get-LargeFolderName
                $optHomeDirs = New-Object System.Management.Automation.Host.ChoiceDescription "&User Home Directory", "Query Active Directory for the User's Home Directory attribute and Export the PST to that location"
                $optCentralExport = New-Object System.Management.Automation.Host.ChoiceDescription "&Centralized Location", "Select a folder location to store all PST Export files"
                $Options = [System.Management.Automation.Host.ChoiceDescription[]]($optHomeDirs, $optCentralExport)
                Switch ($Host.UI.PromptForChoice("`n >> PST EXPORT LOCATION <<", "`n Which type of location do you want to store the PST Export Files?`n`r", $Options, 1)) {
                    0 {
                        Show-Menu -Title "WARNING: HOME DIRECTORY EXPORT SELECTED!" -ClearScreen -DisplayOnly -Style Info -Color Yellow
                        Write-Host $HomeDirWarning -ForegroundColor Yellow
                        If ((Get-ChoicePrompt -OptionList "&OK", "&Cancel" -Title " OK to Continue, Cancel to exit`n`r" -Message $null -default 1) -eq 1) {
                            Write-Log -Type INFO -Text "// HOME DIRECTORY SELECTION | User Cancelled the operation!"
                            Return
                        } Else {
                            $Script:CentralizedExport = $False
                            $Script:ExportPath = "Home Directories"
                        }
                    }
                    1 {
                        Show-Menu -Title "WARNING: CENTRALIZED EXPORT SELECTED!" -ClearScreen -DisplayOnly -Style Info -Color Yellow
                        Write-Host $CentralizedExportWarning -ForegroundColor Yellow
                        If ((Get-ChoicePrompt -OptionList "&OK", "&Cancel" -Title " OK to Continue, Cancel to exit`n`r" -Message $null -default 1) -eq 1) {
                            Write-Log -Type INFO -Text "// CENTRALIZED EXPORT SELECTION | User Cancelled the operation!"
                            Return
                        } Else {
                            $Script:CentralizedExport = $True
                            $NetworkSharePath = Get-NetworkSharePath
                            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
                            $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                            $FolderDialog.SelectedPath = $NetworkSharePath
                            Do { $DialogResult = $FolderDialog.ShowDialog() }
                            Until ($DialogResult -eq "OK")
                            $Script:ExportPath = $FolderDialog.SelectedPath
                        }
                    }
                }
            }
        }
        #endregion
        #region Mailbox Selection Process (Single User, CSV, or ALL)

        Write-Log -Type INFO -Text ("Determining which Exchange Mailboxes to process")
        $optSingleUser = New-Object System.Management.Automation.Host.ChoiceDescription "&Single User", "Provide a single user PrimarySMTPAddress"
        $optCSVImport = New-Object System.Management.Automation.Host.ChoiceDescription "&CSV Import", "Select a CSV file to Import - MUST HAVE PrimarySMTPAddress as a header value"
        $optAllMailboxes = New-Object System.Management.Automation.Host.ChoiceDescription "&ALL MAILBOXES", "Query Exchange for ALL MAILBOXES in the Organization"
        $Options = [System.Management.Automation.Host.ChoiceDescription[]]($optSingleUser, $optCSVImport, $optAllMailboxes)
        Switch ($Host.UI.PromptForChoice("`n >> MAILBOX SELECTION <<", "`n Please select which Mailboxes to peform Large Item Checks?`n`r", $Options, 0)) {
            # Single User
            0 {
                $MailboxSelection = "SINGLE USER"
                Write-Log -Type INFO -Text ("Mailbox Selection Criteria:  {0}" -f $MailboxSelection)
                $Script:MailboxObjects = Get-SingleMailboxUser
            }
            # CSV Import
            1 {
                $MailboxSelection = "CSV IMPORT"
                Write-Log -Type INFO -Text ("Mailbox Selection Criteria:  {0}" -f $MailboxSelection)
                Do {
                    $ImportFile = Get-CSVFileDialog
                    If ([System.String]::IsNullOrEmpty($ImportFile)) {
                        Write-Log -Type ERROR -Text ("No File was selected")
                        Return
                    } Else {
                        If ((Import-Csv -Path $ImportFile | Get-Member).Name -Contains "PrimarySMTPAddress") { $ValidCSVFile = $true }
                        Else {
                            Write-Log -Type WARNING -Text ("The CSV File {0} does not contain a valid PrimarySMTPAddress header!" -f $ImportFile)
                            $ValidCSVFile = $false
                            Start-Sleep -Milliseconds 999
                        }
                    }
                } Until ($ValidCSVFile -eq $True)
                $Script:MailboxObjects = Import-Csv -Path $ImportFile
                Write-Log -Type INFO -Text ("Mailbox Objects to process:  {0:N0}" -f $Script:MailboxObjects.Count)
                Start-Sleep -Milliseconds 999
            }
            # ALL MAILBOXES
            2 {
                $MailboxSelection = "ALL MAILBOXES"
                Write-Log -Type INFO -Text ("Mailbox Selection Criteria:  {0}" -f $MailboxSelection)

                Show-Menu -Title "WARNING: ALL MAILBOXES SELECTED!" -ClearScreen -DisplayOnly -Style Info -Color Yellow
                Write-Host $GetMailboxWarning -ForegroundColor Yellow
                Switch (Get-ChoicePrompt -OptionList "&OK", "&Cancel" -Title " OK to Continue, Cancel to exit`n`r" -Message $null -default 1) {
                    0 {
                        If ((Get-Job).Count -gt 0) {
                            Write-Log -Type WARNING -Text ("Multiple Jobs exists, running Remove-Job...")
                            Start-Sleep -Milliseconds 2250
                            Get-Job | Remove-Job
                        }
                        If (Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }) {
                            Write-Log -Type INFO -Text ("Getting current Exchange PS Session")
                            If (-NOT $CurrentPSSession) { $CurrentPSSession = Get-PSSession -InstanceId (Get-OrganizationConfig).RunspaceId.Guid }
                            Write-Log -Type INFO -Text ("PS Session Connected to: {0}" -f $CurrentPSSession.ComputerName)
                            $ScriptBlock = { Get-Mailbox -ResultSize Unlimited -WA SilentlyContinue | Select-Object PrimarySMTPAddress }
                            Invoke-Command -Session $CurrentPSSession -ScriptBlock $ScriptBlock -AsJob -JobName "GetMailboxes" | Out-Null
                            Write-Host "`n`r Getting Mailboxes, please wait..." -NoNewline
                            While ((Get-Job -Name GetMailboxes).State -eq "Running") {
                                Write-Host "." -NoNewline
                                Start-Sleep -Milliseconds 7500
                            }

                            If ((Get-Job -Name GetMailboxes).State -eq "Failed") {
                                Write-Host "FAILED!" -ForegroundColor White -BackgroundColor Red
                                Write-Log -Type ERROR -Text ("Background Job Failed running Get-Mailboxes, using alternate method (Jobless)")
                                $Script:MailboxObjects = Get-Mailbox -ResultSize Unlimited -WA SilentlyContinue | Where-Object { $_.PrimarySMTPAddress -notlike "extest*" } | Select-Object PrimarySMTPAddress
                                Write-Log -Type INFO -Text ("Mailbox Objects to process:  {0:N0}" -f $MailboxObjects.Count)
                            } ElseIf ((Get-Job -Name GetMailboxes).State -eq "Completed") {
                                Write-Host "SUCCESS! " -ForegroundColor Green
                                $Script:MailboxObjects = Get-Job -Name GetMailboxes | Receive-Job
                                $MailboxObjects = $MailboxObjects | Where-Object { $_.PrimarySMTPAddress -notlike "extest*" } | Select-Object PrimarySMTPAddress
                                If (($Script:MailboxObjects | Measure-Object).Count -gt 0) { Get-Job -Name GetMailboxes | Remove-Job }
                                Write-Log -Type INFO -Text ("Mailbox Objects to process:  {0:N0}" -f $MailboxObjects.Count)
                                Start-Sleep -Milliseconds 2250
                            } Else {
                                Write-Host "UNKNOWN!" -ForegroundColor Black -BackgroundColor Gray
                                Write-Log -Type ERROR -Text ("Getting Mailboxes failed in an UNKNOWN State")
                                Return
                            }
                        } Else {
                            Write-Log -Type ERROR -Text ("Failed to find a valid PowerShell Session")
                            Return
                        }
                    }
                    1 {
                        Write-Log -Type INFO -Text "// ALL MAILBOXES SELECTION | User Cancelled the operation!"
                        Return
                    }
                }
            }
        }
        #endregion
        [System.Console]::Clear()
        $InfoBanner = (@"

################################################################################
#  LARGE ITEM CHECK PRE-EXECUTION REPORT                                       
#  -----------------------------------------------------------------------------
#  [*] SERVICE ACCOUNT:      $ServiceAccountName
#  [*] SIZE LIMIT:           $ItemSizeLimit
#  [*] EXCHANGE VERSION:     $ExchangeVersion
#  [*] PS SESSION:           $($CurrentPSSession.ComputerName)
#  [*] MAILBOX LOCATION:     $MailboxLocation
#  [*] ACTION:               $Action
#  [*] LARGE ITEM FOLDER:    $LargeFolderName
#  [*] CENTRALIZE EXPORT:    $CentralizedExport
#  [*] EXPORT PATH:          $ExportPath
#  [*] MAILBOX SELECTION:    $MailboxSelection
#  [*] MAILBOX OBJECTS:      $("{0:N0}" -f ($MailboxObjects | Measure-Object).Count)
#                          
################################################################################             

Press any key to continue or [N] to cancel
"@)

        [System.Console]::ForegroundColor = "Cyan"
        $Execute = Read-Host -Prompt $InfoBanner
        [System.Console]::ResetColor()
        If ($Execute.ToUpper() -eq "N") {
            Write-Log -Type INFO -Text (" // PRE-EXECUTION REPORT | User cancelled the operation")
            Return
        }
        
        Write-Debug "Start Foreach Loop"
        $i = 0
        $Count = ($MailboxObjects | Measure-Object).Count
        Foreach ($Mailbox in $MailboxObjects) {
            If ($Count -gt 1) { Write-Progress -Id 7 -Activity ("Current User: {0}" -f $Mailbox.PrimarySMTPAddress) -Status ("[ACTION: {0}] Processing Mailboxes..." -f $Action) -CurrentOperation ("Mailbox {0} of {1}" -f ($i + 1), $Count) -PercentComplete (($i / $Count) * 100) }

            Write-Log -Type INFO -Text ("MAILBOX: {0} | Getting Active Directory Properties" -f $Mailbox.PrimarySMTPAddress)
            $Script:ADInfo = Get-ADProperties -Property mail -Value $Mailbox.PrimarySMTPAddress

            If ($Action -eq "ExportLargeItems") {
                Write-Debug "BEGIN ExportLargeItems"
                If ($CentralizedExport) { Export-LargeItems -Identity $Mailbox.PrimarySMTPAddress -Path $ExportPath -FolderName $LargeFolderName }
                Else { Export-LargeItems -Identity $Mailbox.PrimarySMTPAddress -Path $ADInfo.homedirectory -FolderName $LargeFolderName }
            } Else {
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Creating EWS Impersonation Service Object" -f $Mailbox.PrimarySMTPAddress)
                Write-Debug "Creating EWS Service"
                If ([System.String]::IsNullOrEmpty($Uri.AbsoluteUri)) { $EWSService = New-ImpersonationService -Identity $Mailbox.PrimarySMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ExchangeVersion Exchange2013_SP1 }
                Else { $EWSService = New-ImpersonationService -Identity $Mailbox.PrimarySMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ExchangeVersion Exchange2013_SP1 -Uri $uri.AbsoluteUri }
                
                <# Commented out due to Exchange 2010 reaching end of support
                Switch ($ExchangeVersion) {
                    "E14" {
                        If ([System.String]::IsNullOrEmpty($Uri.AbsoluteUri)) { $EWSService = New-ImpersonationService -Identity $Mailbox.PrimarySMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ExchangeVersion Exchange2010_SP2 }
                        Else { $EWSService = New-ImpersonationService -Identity $Mailbox.PrimarySMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ExchangeVersion Exchange2010_SP2 -Uri $uri.AbsoluteUri }
                    }
                    "E15" {
                        If ([System.String]::IsNullOrEmpty($Uri.AbsoluteUri)) { $EWSService = New-ImpersonationService -Identity $Mailbox.PrimarySMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ExchangeVersion Exchange2013_SP1 }
                        Else { $EWSService = New-ImpersonationService -Identity $Mailbox.PrimarySMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ExchangeVersion Exchange2013_SP1 -Uri $uri.AbsoluteUri }
                    }
                }
                #>

                Write-Log -Type INFO -Text ("MAILBOX: {0} | Attempting to get Mailbox Folders ({1})" -f $Mailbox.PrimarySMTPAddress, $MailboxLocation)
                Write-Debug "Get Mailbox Folders"
                $MailboxFolders = Get-MailboxFolders -Service $EWSService -SearchLocation $MailboxLocation

                Write-Log -Type INFO -Text ("MAILBOX: {0} | Attempting to get Mailbox Large Items" -f $Mailbox.PrimarySMTPAddress)
                Write-Debug "Get Mailbox Large Items"
                If ($MailboxFolders.DisplayName -contains $LargeFolderName) { $MailboxFolders = $MailboxFolders | Where-Object { $_.DisplayName -ne $LargeFolderName } }
                $MailboxLargeItems = Get-FolderItems -ItemSizeLimit $ItemSizeLimit -Folders $MailboxFolders -Service $EWSService -Action $Action
                If (($MailboxLargeItems | Measure-Object).Count -gt 0) {
                    Write-Log -Type INFO -Text ("MAILBOX: {0} | Merging Mailbox Large Items to Large Item Output" -f $Mailbox.PrimarySMTPAddress)
                    Write-Debug "Merge Large Items"
                    $MailboxLargeItems | ForEach-Object { [Void]$Script:AllLargeItems.Add($_) }
                    If ($Action -eq "MoveAndExportItems") {
                        If ($CentralizedExport) { Export-LargeItems -Identity $Mailbox.PrimarySMTPAddress -Path $ExportPath -FolderName $LargeFolderName }
                        Else { Export-LargeItems -Identity $Mailbox.PrimarySMTPAddress -Path $ADInfo.homedirectory -FolderName $LargeFolderName }
                    }
                }
            }
            $i++
        }
        Write-Progress -Id 7 -Activity ("Current User: {0}" -f $Mailbox.PrimarySMTPAddress) -Completed
        Write-Debug "end of loop"
        If (($Script:AllLargeItems | Measure-Object).Count -gt 0) { $Script:AllLargeItems | New-LargeItemReport }
    } Catch {
        Write-Debug "Catch block"
        Write-Log -Type ERROR -Text ("Function terminated unexpectedly, details saved to $ScriptPath\ErrorOutput.txt") -Verbose
        $ErrorOutput = $_ | Get-ErrorDetails -ScriptSyntax $PSCmdlet.MyInvocation.Line
        [System.Console]::ForegroundColor = "Red"
        $ErrorOutput
        [System.Console]::ResetColor()
        $ErrorOutput | Out-File "$ScriptPath\ErrorOutput.txt" -Force
    } Finally {
        $Stopwatch.Stop()
        $ElapsedTime = ("{0} Days, {1} Hours, {2} Minutes, {3} Seconds" -f $StopWatch.Elapsed.Days, $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds)
        Write-Log -Type INFO -Text ("Large Items found: {0:N0}" -f $Script:AllLargeItems.Count) -Verbose
        Write-Log -Type INFO -Text ("Log file: {0}" -f $LogFile) -Verbose
        Write-Log -Type INFO -Text ("Process Completed | $ElapsedTime") -Verbose
    }
}

Function New-ImpersonationService {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [System.String]$Identity,
        [Parameter(Mandatory = $true)]
        [System.String]$ImpersonatorAccountName,
        [Parameter(Mandatory = $true)]
        [System.String]$ImpersonatorAccountPassword,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Exchange2010_SP2", "Exchange2013_SP1")]
        $ExchangeVersion,
        [System.URI]$Uri
    )
    BEGIN {
        TRY {
            If (-NOT [System.String]::IsNullOrEmpty($MBX.DisplayName)) { $Identity = $MBX.DisplayName }
            Write-Log -Type INFO -Text ("MAILBOX: {0} | Validating Installation of EWS Managed API" -f $Identity)
            $EWSRegistryPath = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Exchange\Web Services' -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name
            If ($EWSRegistryPath) {
                $EWSInstallDirectory = (Get-ItemProperty -Path Registry::$EWSRegistryPath).'Install Directory'
                $EWSVersion = $EWSInstallDirectory.SubString(($EWSInstallDirectory.Length - 4), 3)
                $EWSDLL = $EWSInstallDirectory + "Microsoft.Exchange.WebServices.dll"

                If (Test-Path $EWSDLL) {
                    If ($EWSVersion -lt 2.0) {
                        $PSCmdlet.ThrowTerminatingError(
                            [System.Management.Automation.ErrorRecord]::New(
                                [System.ArgumentOutOfRangeException]::New("EWS Version is too old, Please install EWS Managed API 2.0 or later"),
                                "EWSVersionOutOfDate",
                                [System.Management.Automation.ErrorCategory]::InvalidResult,
                                $EWSVersion
                            )
                        )
                    } Else {
                        Write-Log -Type INFO -Text ("MAILBOX: {0} | EWS Managed API 2.0 or later is INSTALLED" -f $Identity)
                        Import-Module $EWSDLL
                    }
                } Else {
                    $PSCmdlet.ThrowTerminatingError(
                        [System.Management.Automation.ErrorRecord]::New(
                            [System.IO.FileNotFoundException]::New("Unable to find EWS Managed API DLL"),
                            "FileNotFound",
                            [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                            $EWSDLL
                        )
                    )
                }
            } Else {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::New(
                        [System.IO.FileNotFoundException]::New("EWS Managed API Registry Path Not Found"),
                        "FileNotFound",
                        [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                        "HKLM:\SOFTWARE\Microsoft\Exchange\Web Services"
                    )
                )
            }
        } CATCH { $PSCmdlet.ThrowTerminatingError($PSItem) }
    }
    PROCESS {
        TRY {
            <# SSL Check / Bypass functionality
                . [AUTHOR]
                . Carter Shanklin
                . 
                . [URL]
                . http://poshcode.org/624
                #>
            $Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
            [Void]$Provider.CreateCompiler()
            $Params = New-Object System.CodeDom.Compiler.CompilerParameters
            $Params.GenerateExecutable = $False
            $Params.GenerateInMemory = $True
            $Params.IncludeDebugInformation = $False
            [Void]$Params.ReferencedAssemblies.Add("System.DLL")

            $TASource = @"
namespace Local.ToolkitExtensions.Net.CertificatePolicy {
       public class TrustAll : System.Net.ICertificatePolicy {
              public TrustAll() {}
        public bool CheckValidationResult(
                     System.Net.ServicePoint sp,
                     System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                     System.Net.WebRequest req, int problem
              ) {return true;}
    }
}
"@

            $TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
            $TAAssembly = $TAResults.CompiledAssembly
            $TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
            [System.Net.ServicePointManager]::CertificatePolicy = $TrustAll

            Write-Debug ("[{0}] Check Exchange Version" -f $PSCmdlet.MyInvocation.MyCommand)
            Switch ($ExchangeVersion) {
                "Exchange2010_SP2" {
                    Write-Log -Type INFO -Text ("MAILBOX: {0} | Creating EWS Service Object (Exchange2010_SP2)" -f $Identity)
                    [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)
                }
                "Exchange2013_SP1" {
                    Write-Log -Type INFO -Text ("MAILBOX: {0} | Creating EWS Service Object (Exchange2013_SP1)" -f $Identity)
                    [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
                }
            }
                     
            $Service.Credentials = New-Object Net.NetworkCredential($ImpersonatorAccountName, $ImpersonatorAccountPassword)

            Write-Debug ("[{0}] Check for Uri value to determine Autodiscover address" -f $PSCmdlet.MyInvocation.MyCommand)
            If ([String]::IsNullOrEmpty($Uri)) {
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Autodiscover in process" -f $Identity)
                $Service.AutodiscoverUrl($Identity, { $True })
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Using EWS URL: {1}" -f $Identity, $Service.Url)
            } Else { $Service.Url = $Uri.AbsoluteUri }

            Write-Log -Type INFO -Text ("MAILBOX: {0} | Attemping to Impersonate {1}" -f $Identity, $Identity)
            $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Identity)
       
            #Increase the timeout for larger mailboxes
            $Service.Timeout = 150000
            Return $Service
        } CATCH {
            Write-Log -Type ERROR -Text ("MAILBOX: {0} | Failed to Impersonate the User" -f $Identity)
            Write-Log -Type ERROR -Text ("{0}" -f $_.Exception.Message)
            Continue
        }
    }      
}

Function Get-MailboxFolders {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $Service,
        [Parameter(Mandatory = $true)]           
        [ValidateSet("Primary", "Archive", "RecoverableItems")]
        [String]$SearchLocation
    )
    try {
        [Microsoft.Exchange.WebServices.Data.FolderView]$View = New-Object Microsoft.Exchange.WebServices.Data.FolderView([System.Int32]::MaxValue)
        $View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $View.PropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
        $View.PropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount)
        $View.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep

        Switch ($SearchLocation) {
            "Primary" {
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Finding WellKnownFolders in MsgFolderRoot" -f $Service.ImpersonatedUserId.Id)
                [Microsoft.Exchange.WebServices.Data.FindFoldersResults]$Folders = $Service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $View)
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Found {1} folders (PRIMARY MAILBOX)" -f $Service.ImpersonatedUserId.Id, $Folders.TotalCount)
            }
            "Archive" {
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Finding WellKnownFolders in ArchiveMsgFolderRoot" -f $Service.ImpersonatedUserId.Id)
                [Microsoft.Exchange.WebServices.Data.FindFoldersResults]$Folders = $Service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $View)
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Found {1} folders (ARCHIVE MAILBOX)" -f $Service.ImpersonatedUserId.Id, $Folders.TotalCount)
            }
            "RecoverableItems" {
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Finding WellKnownFolders in RecoverableItemsRoot" -f $Service.ImpersonatedUserId.Id)
                [Microsoft.Exchange.WebServices.Data.FindFoldersResults]$Folders = $Service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsRoot, $View)
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Found {1} folders (RECOVERABLE ITEMS STORE)" -f $Service.ImpersonatedUserId.Id, $Folders.TotalCount)
            }
        }
        Return $Folders
    } CATCH { 
        Write-Log -Type ERROR -Text ("MAILBOX: {0} | Failed to get Mailbox Folders ({1})" -f $Service.ImpersonatedUserId.Id, $SearchLocation)
        Write-Log -Type ERROR -Text ("{0}" -f $_.Exception.Message)
        Continue
    }
}

Function Get-FolderItems {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [int]$ItemSizeLimit,
        [Parameter(Mandatory = $true)]
        $Folders,
        [Parameter(Mandatory = $true)]
        $Service,
        [Parameter(Mandatory = $true)]
        $Action
    )
       
    $LargeItemCount = 0
    [System.Collections.ArrayList]$colLargeItems = @()
    $fldrIndex = 0
    foreach ($Folder in $Folders) {
        $CurrentFolder = $Folder.DisplayName
        Write-Progress -ParentId 7 -Id 42 -Activity ("Current Folder: {0}" -f $CurrentFolder) -Status "Checking folders for items larger than: $ItemSizeLimit MB" -CurrentOperation ("Processing: {0:N0} of {1:N0} | Large Items: {2}" -f ($fldrIndex + 1), $Folders.Count, $LargeItemCount) -PercentComplete (($fldrIndex / $Folders.count) * 100)
        $Items = $Null
        $PageSize = 1000
        $Offset = 0   
        $MoreItemsAvailable = $True     
        Write-Log -Type INFO -Text ("MAILBOX: {0} | Started Processing folder: {1}" -f $Service.ImpersonatedUserId.Id, $CurrentFolder)
        $TotalItems = 0

        Do {
            try {
                $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
                $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                $ItemView.PropertySet = $PropertySet
                $ItemView.PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
                $ItemView.PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
                $ItemView.PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass)
                $ItemView.PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
                Write-Debug "FindItems Method on Folder: $CurrentFolder"
                $Items = $Folder.FindItems($ItemView)
            } catch {
                Write-Debug ("CATCH BLOCK -> Folder: {0}" -f $CurrentFolder)
                Write-Log -Type ERROR -Text ("MAILBOX: {0}, FOLDER: {1} | Unable to read items, SKIPPING Folder" -f $Service.ImpersonatedUserId.Id, $CurrentFolder)
                Write-Log -Type ERROR -Text ("{0}" -f $_.Exception.Message)
                $MoreItemsAvailable = $False
                Break
            }

            $TotalItems = ($TotalItems + ($Items | Measure-Object).Count)
            Write-Progress -ParentId 42 -Activity ("Folder contains MORE THAN {0:N0} Items:  {1}" -f $PageSize, $Items.MoreAvailable) -Status ("Processing items...") -CurrentOperation ("Items in folder: {0:N0}" -f $TotalItems) -PercentComplete -1
            $LargeItems = $null
            $LargeItems = $Items | Where-Object { [Math]::Round(($_.Size / 1000000), 2) -gt $ItemSizeLimit }
            If (($LargeItems | Measure-Object).Count -gt 0) {
                Write-Debug ("Found {0} Items Larger than {1} MB" -f ($LargeItems | Measure-Object).Count, $ItemSizeLimit)
                Foreach ($Item in $LargeItems) {
                    If ([System.String]::IsNullOrEmpty($Item.Subject)) {
                        $Subject = "NULL"
                        $ItemViolation = New-LargeItemViolation -SMTPAddress $Service.ImpersonatedUserId.Id -Created $Item.DateTimeCreated -Subject $Subject -FolderDisplayName $Folder.DisplayName -Size $Item.Size -ItemClass $Item.ItemClass
                        [Void]$colLargeItems.Add($ItemViolation)
                    } Else {
                        Write-Debug ("[{0}] - Logging Item Violation" -f $PSCmdlet.MyInvocation.MyCommand)
                        $ItemViolation = New-LargeItemViolation -SMTPAddress $Service.ImpersonatedUserId.Id -Created $Item.DateTimeCreated -Subject $Item.Subject -FolderDisplayName $Folder.DisplayName -Size $Item.Size -ItemClass $Item.ItemClass
                        [Void]$colLargeItems.Add($ItemViolation)
                    }

                    If ($Action -eq "MoveLargeItems" -or $Action -eq "MoveAndExportItems") {
                        Write-Debug ("Moving Large Items to: {0}" -f $LargeFolderName)
                        $LargeItemFolder = New-MailboxFolder -FolderName $LargeFolderName -Service $Service
                        If ($LargeItemFolder) {
                            Write-Log -Type INFO -Text ("MAILBOX: {0} | Moving Item [{1}] from folder [{2}] to folder [{3}]" -f $Service.ImpersonatedUserId.Id, $Item.Subject, $CurrentFolder, $LargeFolderName)
                            [void]$Item.Move($LargeItemFolder.Id)
                        } Else { Write-Log -Type ERROR -Text ("MAILBOX: {0} | Unable to Move Item [{1}], FAILED to create folder [{2}]" -f $Service.ImpersonatedUserId.Id, $Item.Subject, $LargeFolderName) }
                    }
                    $LargeItemCount++
                }
            }

            If ($Items.MoreAvailable -eq $False) {
                $MoreItemsAvailable = $false
                Write-Log -Type INFO -Text ("MAILBOX: {0} | Finished Processing folder: {1} ({2:N0} Items)" -f $Service.ImpersonatedUserId.Id, $CurrentFolder, $TotalItems)
            } ElseIf ($Items.MoreAvailable -eq $true) { $Offset += $PageSize }
        }
        While ($MoreItemsAvailable)
        Write-Progress -ParentId 42 -Activity ("Processing items...") -Completed
        $fldrIndex++
    }
    Write-Progress -Id 1 -Activity "Checking folders for items larger than: $ItemSizeLimit MB" -Completed
    Write-Log -Type INFO -Text ("MAILBOX: {0} | Number of Large Items found: {1}" -f $Service.ImpersonatedUserId.Id, $LargeItemCount)
    $colLargeItems
}

Function New-LargeItemViolation {
    [CmdletBinding()]
    Param (
        [System.String]$SMTPAddress,
        [System.String]$Subject,
        [System.String]$ItemClass,
        [System.String]$FolderDisplayName,
        [System.DateTime]$Created,
        [System.Int32]$Size
    )
    Write-Debug "Item Violation"
    $Violations = [PSCustomObject][Ordered]@{
        PrimarySMTPAddress = $SMTPAddress
        Subject            = $Subject
        ItemClass          = $ItemClass
        FolderName         = $FolderDisplayName
        CreatedDate        = [DateTime]$Created
        Size               = ("{0:N2}" -f [Math]::Round($Size / 1000000, 2))
    }
    $Violations
}

Function New-MailboxFolder {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [String]$FolderName,
        [Parameter(Mandatory = $true)]
        $Service
    )

    try {
        Write-Debug "start"
        $FolderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
        $FolderRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $FolderID)
        $View = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        $Filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $FolderName)
        $LargeItemFolder = $Service.FindFolders($FolderRoot.Id, $Filter, $View)
        If ($LargeItemFolder.DisplayName -ne $FolderName) {
            Write-Log -Type INFO -Text ("MAILBOX: {0} | Folder {1} was not found, creating the folder" -f $Service.ImpersonatedUserId.Id, $FolderName)
            $LargeItemFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
            $LargeItemFolder.DisplayName = $FolderName
            $LargeItemFolder.FolderClass = "IPF.Note"
            $LargeItemFolder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
            $LargeItemFolder
        }
        Else { $LargeItemFolder }
    } catch {
        Write-Log -Type ERROR -Text ("MAILBOX: {0} | Unable to create or find the {1} Folder" -f $Service.ImpersonatedUserId.Id, $FolderName)
        Write-Log -Type ERROR -Text $_.Exception.Message
    }
}

Function New-LargeItemReport {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNull()]
        $InputObject
    )
    begin {
        $OutputFile = ("{0}\{1}_Large_Item_Violations.csv" -f $ScriptPath, (Get-Date -Format yyyyMMdd_HHmmss))
        Show-Menu -Title ("Creating the Large Item Report: {0}" -f $OutputFile) -DisplayOnly -Style Info -Color Green
        Write-Log -Type INFO -Text ("Exporting records to {0}" -f $OutputFile) -Verbose
    }
    process { $InputObject | Export-Csv $OutputFile -NoTypeInformation -Append }
}

Function Export-LargeItems {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $Identity,
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [System.String]$Path,
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [System.String]$FolderName
    )
    If (Test-Path $Path) {
        Write-Log -Type INFO -Text ("MAILBOX: {0} | Export Path Validated ({1})" -f $Identity, $Path)
        Write-Log -Type INFO -Text ("MAILBOX: {0} | Attempting to Export Large Item Folder to PST" -f $Identity)

        If (Get-MailboxExportRequest -Mailbox $Identity) {
            Write-Warning ("[{0}] | Found Existing Mailbox Export Request!" -f $Identity)
            Write-Log -Type INFO -Text ("MAILBOX: {0} | FOUND and REMOVING Mailbox Export Request" -f $Identity)
            Get-MailboxExportRequest -Mailbox $Identity | Remove-MailboxExportRequest -Confirm:$False
            Write-Log -Type INFO -Text ("MAILBOX: {0} | Creating New Mailbox Export Request" -f $Identity)
            $NewMERObject = New-MailboxExportRequest -Name ("{0}_LargeItems" -f $ADInfo.samaccountname) -Mailbox $Identity -FilePath ("{0}\{1}_LargeItems.pst" -f $Path, $ADInfo.samaccountname) -IncludeFolders $FolderName -Confirm:$False -ExcludeDumpster -ErrorAction SilentlyContinue
            If ($NewMERObject) { Write-Log -Type INFO -Text ("MAILBOX: {0} | New Mailbox Export Request created SUCCESSFULLY" -f $Identity) }
            Else {
                Write-Warning ("[{0}] | FAILED to Create New Mailbox Export Request" -f $Identity)
                Write-Log -Type WARNING -Text ("MAILBOX: {0} | New Mailbox Export Request FAILED" -f $Identity)
            }
        }
        Else {
            
            Write-Log -Type INFO -Text ("MAILBOX: {0} | Creating New Mailbox Export Request" -f $Identity)
            Write-Debug "Create New Export Request"
            $NewMERObject = New-MailboxExportRequest -Name ("{0}_LargeItems" -f $ADInfo.samaccountname) -Mailbox $Identity -FilePath ("{0}\{1}_LargeItems.pst" -f $Path, $ADInfo.samaccountname) -IncludeFolders $FolderName -Confirm:$False -ExcludeDumpster -ErrorAction SilentlyContinue
            
            If ($NewMERObject) { Write-Log -Type INFO -Text ("MAILBOX: {0} | New Mailbox Export Request created SUCCESSFULLY" -f $Identity) }
            Else {
                Write-Warning ("[{0}] | FAILED to Create New Mailbox Export Request" -f $Identity)
                Write-Log -Type WARNING -Text ("MAILBOX: {0} | New Mailbox Export Request FAILED" -f $Identity)
            }
        }
    }
    Else {
        Write-Warning ("[{0}] | FAILED to Validate the Export Path" -f $Identity)
        Write-Log -Type WARNING -Text ("MAILBOX: {0} | Export to PST failed becasue the path: {1} was not found" -f $Identity, $Path)
    }
}

Function Get-ErrorDetails {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $ErrorRecord,
        $ScriptSyntax
    )
    process {
        If ($ErrorRecord -is [Management.Automation.ErrorRecord]) {
            [PSCustomObject]@{
                Reason     = $ErrorRecord.CategoryInfo.Reason
                Exception  = $ErrorRecord.Exception.Message
                Target     = $ErrorRecord.CategoryInfo.TargetName
                Script     = $ErrorRecord.InvocationInfo.ScriptName
                Syntax     = $ScriptSyntax
                Command    = $ErrorRecord.InvocationInfo.MyCommand
                LineNumber = $ErrorRecord.InvocationInfo.ScriptLineNumber
                Column     = $ErrorRecord.InvocationInfo.OffsetInLine
                Date       = Get-Date
                User       = $env:USERNAME
            }
        }
    }
}

Function Get-ChoicePrompt {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [String[]]$OptionList, 
        [Parameter(Mandatory = $false)]
        [String]$Title, 
        [Parameter(Mandatory = $False)]
        [String]$Message = $null, 
        [int]$Default = 0 
    )
    $Options = New-Object System.Collections.ObjectModel.Collection[System.Management.Automation.Host.ChoiceDescription] 
    $OptionList | ForEach-Object { $Options.Add((New-Object "System.Management.Automation.Host.ChoiceDescription" -ArgumentList $_)) } 
    $Host.ui.PromptForChoice($Title, $Message, $Options, $Default) 
}

Function Show-Menu {
    Param(
        [Parameter(Mandatory = $true)]
        [System.String]$Title,
        [System.String]$Menu,
        [Switch]$ClearScreen,
        [Switch]$DisplayOnly,
        [ValidateSet("Full", "Mini", "Info", "None")]
        $Style,
        [ValidateSet("White", "Cyan", "Magenta", "Yellow", "Green", "Red", "Gray", "DarkGray")]
        $Color = "Gray"
    )
    If ($ClearScreen) { [System.Console]::Clear() }

    Switch ($Style) {
        "Full" {
            $menuPrompt = "/" * (95)
            $menuPrompt += "`n`r////`n`r//// $Title`n`r////`n`r"
            $menuPrompt += "/" * (95)
            $menuPrompt += "`n`n"
        }
        "Mini" {
            $menuPrompt = "`n`r"
            $menuPrompt += "\" * (80)
            $menuPrompt += "`n\\\\  $Title`n"
            $menuPrompt += "\" * (80)
            $menuPrompt += "`n"
        }
        "Info" {
            $menuPrompt = "`n`r"
            $menuPrompt += "-" * (80)
            $menuPrompt += "`n-- $Title`n"
            $menuPrompt += "-" * (80)
            $menuPrompt += "`n"
        }
        "None" {
            $menuPrompt = $Title
        }
        Default {
            $menuPrompt = "`n`r"
            $menuPrompt += "\" * (80)
            $menuPrompt += "`n\\\\  $Title`n"
            $menuPrompt += "\" * (80)
            $menuPrompt += "`n"
        }
    }

    [System.Console]::ForegroundColor = $Color
    If ($DisplayOnly) { Write-Host $menuPrompt }
    Else {
        $menuPrompt += $menu
        Read-Host -Prompt $menuprompt
    }
    [System.Console]::ResetColor()
}

Function Write-Log {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [String]$Type = "INFO",
        [String]$Text
    )
    If (-Not ([System.String]::IsNullOrEmpty($Logfile))) {
        If (-Not (Test-Path -Path $LogFile)) {
            If ($VerbosePreference -eq "Continue") { Write-Verbose "[$Type] - $Text" }
            Else {
                If ($Type -eq "WARNING") { Write-Host "[$Type] - $Text" -ForegroundColor Yellow }
                If ($Type -eq "ERROR") { Write-Host "[$Type] - $Text" -ForegroundColor Red }
            }
            New-Item $LogFile -ItemType File -Force | Out-Null
            $fsMode = [System.IO.FileMode]::Append
            $fsAccess = [System.IO.FileAccess]::Write
            $fsSharing = [System.IO.FileShare]::Read
            $fsLog = New-Object System.IO.FileStream($Logfile, $fsMode, $fsAccess, $fsSharing)
            $swLog = New-Object System.IO.StreamWriter($fsLog)
            $swLog.WriteLine("$(Get-Date), [$Type], ====> $Text")
            $swLog.Close()
        } Else {
            If ($VerbosePreference -eq "Continue") { Write-Verbose "[$Type] - $Text" }
            Else {
                If ($Type -eq "WARNING") { Write-Host "[$Type] - $Text" -ForegroundColor Yellow }
                If ($Type -eq "ERROR") { Write-Host "[$Type] - $Text" -ForegroundColor Red }
            }
            $fsMode = [System.IO.FileMode]::Append
            $fsAccess = [System.IO.FileAccess]::Write
            $fsSharing = [System.IO.FileShare]::Read
            $fsLog = New-Object System.IO.FileStream($Logfile, $fsMode, $fsAccess, $fsSharing)
            $swLog = New-Object System.IO.StreamWriter($fsLog)
            $swLog.WriteLine("$(Get-Date), [$Type], ====> $Text")
            $swLog.Close()
        }
    } Else { Write-Host "//MISSING LOGFILE// [$Type] - $Text" -ForegroundColor Yellow }
}

Function Get-ADProperties {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [String]$Property,
        [Parameter(Mandatory = $true)]
        [String]$Value
    )

    $ADFilter = "(&($Property=$Value))"
    $ADSearch = New-Object System.DirectoryServices.DirectorySearcher
    $ADSearch.ClientTimeout = "00:00:15"
    $ADSearch.ServerTimeLimit = "00:00:30"
    $ADSearch.ServerPageTimeLimit = "00:00:15"
    $ADSearch.Filter = $ADFilter
    $colPropList = @("samaccountname", "mailnickname", "homedirectory")
      
    $ADSearch.PropertiesToLoad.AddRange($colPropList)

    $ADResult = $ADSearch.FindOne()
    $ADInfo = $ADResult | Select-Object `
    @{N = "samaccountname"; E = { $_.Properties["samaccountname"] } },
    @{N = "mailnickname"; E = { $_.Properties["mailnickname"] } },
    @{N = "homedirectory"; E = { $_.Properties["homedirectory"] } }
       
    Return $ADInfo
}

Function Test-ConsoleRights {
    Try {
        If (-Not ([security.principal.windowsprincipal][security.principal.windowsidentity]::GetCurrent()).IsInRole([security.principal.windowsbuiltinrole] "administrator")) {
            Return $false
        }
        return $true
    } Catch [System.Exception] {
        Throw "Unable to determine if console is running with elevated permissions"; Return
    }
}

Function Test-ConsoleVersion {
    Try {
        If ($PSVersionTable.PSVersion.Major -gt "2") {
            Return $true        
        }
        Return $false
    } Catch [System.Exception] {
        Throw "Unable to determine console version"; Return
    }
}

Function Get-LargeFolderName {
    [CmdletBinding()]
    Param()
    Show-Menu -Title "Function: Get-LargeFolderName" -DisplayOnly -Style Info -Color White
    $Complete = $false
    $Script:LargeFolderName = Read-Host "Enter the folder name where Large Items will be moved"

    Write-Host ("`n`rLarge Folder Name: {0}" -f $Script:LargeFolderName) -ForegroundColor Cyan
    Do {
        Switch (Get-ChoicePrompt -OptionList "&Yes", "&No" -Message "`nIs the Large Item Folder name correct?" -Default 1) {
            "0" {
                Write-Host "`nLarge Item Folder stored as `$Script:LargeFolderName`n" -ForegroundColor Green
                $Complete = $true
            }
            "1" { Get-LargeFolderName }
        }
    } While ($Complete -eq $false)
}

Function Get-SingleMailboxUser {
    [CmdletBinding()]
    Param()
    Write-Host "`n"
    Show-Menu -Title "Function: Get-SingleMailboxUser" -DisplayOnly -Style Info -Color White
    $SingleMailboxUser = [PSCustomObject] @{
        PrimarySMTPAddress = (Read-Host "Enter the PrimarySMTPAddress of the Mailbox User")
    }

    Write-Host ("`n`rPrimarySMTPAddress: {0}" -f $SingleMailboxUser.PrimarySMTPAddress) -ForegroundColor Cyan

    Switch (Get-ChoicePrompt -OptionList "&Yes", "&No" -Message "`nIs the Mailbox User PrimarySMTPAddress correct?" -Default 1) {
        "0" { Return $SingleMailboxUser }
        "1" { Get-SingleMailboxUser }
    }
}

Function Get-NetworkSharePath {
    [CmdletBinding()]
    Param()
    Show-Menu -Title "Function: Get-NetworkSharePath" -DisplayOnly -Style Info -Color White
    $Complete = $false
    $ServerName = Read-Host " Enter the Server FQDN for the Network Share"
    $ShareName = Read-Host " Enter the Share Folder Name (Hidden shared allowed)"
    $NetworkSharePath = ("\\{0}\{1}" -f $ServerName, $ShareName)

    Write-Host ("`n`r Network Share Path: {0}" -f $NetworkSharePath) -ForegroundColor Cyan
    Do {
        Switch (Get-ChoicePrompt -OptionList "&Yes", "&No" -Title "`n Is the Network Share Path correct?`n" -Message $null -Default 1) {
            "0" { $Complete = $true }
            "1" { Get-NetworkSharePath }
        }
    } While ($Complete -eq $false)
    Return $NetworkSharePath
}

Function Get-CSVFileDialog {
    Param ()
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "[LargeItemChecks] CSV Import File"
    $OpenFileDialog.initialDirectory = [System.Environment]::GetFolderPath("Desktop")
    $OpenFileDialog.filter = "CSV files (*.csv)| *.csv| All files (*.*)| *.*"
    $Result = $OpenFileDialog.ShowDialog()
    If ($Result -eq "OK") {
        $File = $OpenFileDialog.filename
        Return $File
    } Else { Return }
}
#endregion