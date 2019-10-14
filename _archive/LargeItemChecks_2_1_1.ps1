<#
. [COPYRIGHT]
. © 2011-2014 Microsoft Corporation. All rights reserved. 
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
. Dmitry Kazantsev, Senior Consultant
. 
. [CONTRIBUTORS]
. Michael Hall, Service Engineer
. Jayme Bowers, Senior Service Engineer
. Stuart Murray, Consultant
. Jason Parker, Consultant
. 
. [SCRIPT]
. LargeItemChecks_2_1_1.ps1
.
. [VERSION]
. 2.1.1
.
. [VERSION HISTORY / UPDATES]
. 1.1.0 
. Added on-prem Exchange PowerShell snap-in call to ensure that execution dosent have to be initiated explicitly form "Exchange PowerShell" console
.
. 1.1.1
. Fixed invalid parameter names when calling Get-Impersonator "method" (ImpersonatorAccountName and ImpersonatorAccountPassword)
. 
. 1.2.0
. Michael Hall - Converted PowerShell scripts to functions and consolidated code into single script for easier usage.
. Michael Hall - Added version checking for 2007/2010/2013.
. Michael Hall - Added Impersonation Service configuration checks.
. Michael Hall - Added registry checks for Exchange admin tools and Web Services installation.
. Michael Hall - Added write progress and transcription logging.
. Michael Hall - Added export to csv functionality.
. Michael Hall - Added item size limit parameter and validation checks for parameters. 
. Michael Hall - Added Exchange 2007/2010/2013 Impersonation Permissions commands in script header.
.
. 1.2.1
. Michael Hall - Added Autodiscover lookup for each mailbox. 
. Michael Hall - Added the ability to scan specific mailboxes in a CSV
. Michael Hall - Added ignore SSL when using self signed certificate.
.
. 1.2.2
. Jayme Bowers - Added error handling for the following checks:
. Add try/catch/finally blocks for error trapping
. Handle importCSV formatting, i.e. should contain one column of Primary Smtp Addresses with no header; non-SMTP
. address throws error
. Check importCSV contains at least one entry
. Add suggestion on 401/Unauthorized error to check service/impersonation account credentials
. Add suggestion on Autodiscover error to check if mailbox is hosted in current Exchange organization
. Other errors are trapped in the general try/catch blocks, e.g. file system permissions, existence of import CSV file, etc.
. 
. 1.2.3
. Michael Hall - Memory optimization for large mailbox count
. 
. 1.2.4
. Michael Hall - Corrected Exchange 2007 Admin Tools registry check
. Michael Hall - Added Microsoft.Exchange.WebServices Timeout to cater for timeouts on very large mailboxes.
.
. 1.2.6
. Michael Hall - Changed EWS credential mechanism to utilize Net.NetworkCredential com object.
.
. 1.2.7
. Michael Hall - Added archive support.
. 
. 1.2.8
. Michael Hall - Changed item view limit to 1000 per call to avoid resource issues.
.
. 1.3
. Michael Hall - Added ItemClass
. Michael Hall - Fixed weird character exports, reverted back to Export-csv.
.
. 2.0
. Jason Parker - Created additional functions (Create-Folder, Get-ADAttributes)
. Jason Parker - Added Function for Single Mailbox without the need for a CSV import
. Jason Parker - Added -MoveLargeItems parameter
. Jason Parker - Added -ExportLargeItems parameter
. Jason Parker - Added option to export large item folder to PST (custom path or AD Home Directory)
. Jason Parker - Added logic to support Exchange 2010 and 2013 PowerShell Remoting (needed for New-MailboxExportRequest)
. Jason Parker - Removed Transcript and replace with a Log function
. Jason Parker - Fixed issues with items reporting invalid properties and looping issues.  Only searching folders with FolderClass = IPF.Note and items with ItemClass = IPM.Note
. Jason Parker - Fixed the Get-Impersonator function so that it will continue when it fails an Autodiscover check, rather than exit
. Jason Parker - Added validation for ACTION based parameters
. Jason Parker - Added -ExportOnly parameter (does NOT search folders or items - FAST)
. Jason Parker - Added -LargeItemNotification parameter which sends the user an e-mail with information about their export of items to PST
. Jason Parker - Added better support for both Exchange 2007 and Exchange 2010 or newer
.
. 2.0.1
. Jason Parker - Included a sample CSV import file and corrected the comment based help for -ImportCSV
.
. 2.0.2
. Jason Parker - Added a parameter to the Exchange 2010 New-MailboxExportRequest (-ExcludeDumpster).  Causing over inflated PST files.
.
. 2.1.0
. Jason Parker - Resolved issues where the script would run and create an empty CSV with no errors (see below).
. Jason Parker - Added Try/Catch blocks in the Get-Folders function.
. Jason Parker - Added If/Else statement to the $CreateCSV switch so that it wouldn't create an empty CSV file.
.
. 2.1.1
. Jason Parker - Fixed the Try/Catch blocks in Get-Folders function to properly exit when an exception is found.
. Jason Parker - Added custom text outlining how to properly setup application impersonation.
#>

<#
.SYNOPSIS
This script will run a series of cmdlets / functions using Exchange Web Servers to search mailboxes for items over a specified size. Useful for Office 365 engagements where you need to remediate large items before migration.

.DESCRIPTION
This script arranges building-block cmdlets / functions to connect to an Exchange environment and loops through all or a subset of mailboxes with an impersonator account using Exchange Web Services API.  The impersonator account will enumerate every item in every folder and identify items that are exceeding a specific size. The script is designed to be executed before a cross-forest migration or any time an Organization needs to report, export, or move items that may not be compliant with an organization or target organization item size quota / limitation. This script is most commonly used in conjunction with Office 365 migrations due to the 25MB item size limitation.

.PARAMETER ServiceAccountDomain
Specifies the NETBIOS Domain Name from which the Service Account User resides.

.PARAMETER ServiceAccountName
Specifies the SamAccountName of the user which has elevated permissions (impersonation and mailbox export).

.PARAMETER ServicePassword
Specifies the password for the Service Account Users (stored in clear text).

.PARAMETER ItemSizeLimit
Sets the value from which you will measure items against (in MB).  This value should be set to 25, which is the maximum value allowed when moving to Office 365.

.PARAMETER ImportCSV
Specifies the CSV file to be used as the source of mailboxes to search through. The CSV file should contain a single column of SMTP addresses of the mailboxes you want to scan.  There must be a header row with the name "PrimarySMTPAddress".

.PARAMETER CreateCSV
Tells the script to create a master CSV file of all mailboxes and all items found which are in violation of the -ItemSizeLimit parameter.

.PARAMETER MoveLargeItems
Tells the script to move items which are in violation of the -ItemSizeLimit parameter into a specific folder.  Works optionally with -FolderName parameter, but will prompt if not provided.

.PARAMETER ExportLargeItems
Tells the script to peform an EXPORT of the mailbox.  The export function will ONLY export the items which have previously been moved by the -MoveLargeItems parameter and relies on the same -FolderName parameter.  This will not export the entire mailbox.

.PARAMETER ExportOnly
Tells the script NOT to enumerate through all folders and items, rather it will only export items from a specific folder location within the mailbox.  Using this option assumes that you have previously moved the items you want to export into a specific folder.

.PARAMETER LargeItemNotice
Used during an Export Action which will send the user an e-mail (requires Template.htm file) detailing the path of their PST export and other instructions.

.PARAMETER FolderName
Sets the name of the folder to be created when either moving large items or choosing from which folder to export items from.  If not specified, the script will prompt when required by the functions.

.PARAMETER PSTPath
Sets the UNC folder path when exporting the large items to a PST.  If not specified, the script will attempt to use the users home directory.  If neither values exist, the script will abort the export function.

.PARAMETER ArchiveCheck
Valid only for Exchange 2010 / 2013 and will perform all actions above, but for the users archive mailbox if it exists.

.PARAMETER Uri
Sets the Uri for the Exchange Web Services endpoint.  Useful when you can't leverage Autodiscover or Autodiscover fails.

.EXAMPLE
LargeItemChecks_2_1_1.ps1

When running the script with no parameters, it will prompt for any values which are mandatory.  When the script completes, it will display the location of the log file which provides a detailed account of what the script did during the last execution. This method works great for simple testing, but try not to run the script without any parameters, *especially if you are processing ALL or a lot of large mailboxes* because it won't give you any valuable output for the amount of time it took for the script to complete.

.EXAMPLE
LargeItemChecks_2_1_1.ps1 -ServiceAccountDomain <Domain> -ServiceAccountName <User> -ServicePassword <Password> -ItemSizeLimit <Value in MB> -CreateCSV

In this example, the mandatory parameters have been provided and the ACTION -CreateCSV has been enabled which will create a CSV file containing all the item violations from all the mailboxes that were scanned. From this CSV you can create your -ImportCSV file as the source input for when you want to perform any of the other ACTION based switches (-MoveLargeItems or -ExportLargeItems). This will provide a more efficient processing the next time you run the script.

.EXAMPLE
LargeItemChecks_2_1_1.ps1 -ServiceAccountDomain <Domain> -ServiceAccountName <User> -ServicePassword <Password> -ItemSizeLimit <Value in MB> -ImportCSV .\users.csv -MoveLargeItems -ExportLargeItems -LargeItemNotice

In this example, the mandatory parameters have been provided and the ACTION(s) -MoveLargeItems and -ExportLargeItems have been enabled.  The script will prompt for the name of the folder that is to be created and is where the item violations will be moved to.  After all the messages have been moved, the script will attempt to export all the items from the *newly created folder (Large Item Folder)*.  An e-mail will be sent to the user providing them with the location of their PST and the folder that was created.  The e-mail is based on an HTML template file which has been provided with this script.

.EXAMPLE
LargeItemChecks_2_1_1.ps1 -ServiceAccountDomain <Domain> -ServiceAccountName <User> -ServicePassword <Password> -ItemSizeLimit <Value in MB> -ExportLargeItems -ExportOnly

In this example, the mandatory parameters have been provided and the ACTION(s) -ExportLargeItems and -ExportOnly have been enabled.  This is a fast execution which will enumerate through the mailboxes and only attempt to perform and export of the items in a specific folder.  The -ExportOnly switch will *NOT* enumerate or evaluate any items for size compliance.

.NOTES
Large environments will take a significant amount of time to scan (days/weeks). You can reduce the run time by either using a CSV import file with a smaller subset of users or running multiple instances of the script concurrently, targeting mailboxes on different servers.  Running multiple instances assumes your Exchange Web Services endpoint is behind a network load balancer.

Important: Do not run too many instances or against too many mailboxes at once. Doing so could cause performance issues, affecting users.  Microsoft is not responsible for any such performance issue or improper use and planning.

[PERMISSIONS REQUIRED]
This script requires elevated permissions beyond the typical RBAC roles.

[EXCHANGE 2007 IMPERSONATION PERMISSIONS]
Get-ExchangeServer | where {$_.IsClientAccessServer -eq $TRUE} | ForEach-Object {Add-ADPermission -Identity $_.distinguishedname -User (Get-User -Identity ServiceAccount | select-object).identity -ExtendedRight ms-Exch-EPI-Impersonation}
Get-MailboxDatabase | ForEach-Object {Add-ADPermission -Identity $_.DistinguishedName -User ServiceAccount -ExtendedRights ms-Exch-EPI-May-Impersonate}

[EXCHANGE 2010/2013 PERMISSIONS]
There are two sets of permissions required to properly execute the script in an Exchange 2010 / 2013 environment.  Impersonation and Export permissions. Both sets of permissions will require changing or creating	of RBAC Management Role Assignments.

[IMPERSONATION PERMISSIONS]
From the Exchange Management Shell, run the New-ManagementRoleAssignment cmdlet to add the permission to impersonate to the specified user:
New-ManagementRoleAssignment –Name:impersonationAssignmentName –Role:ApplicationImpersonation –User:ServiceAccount

[NEW-MAILBOXEXPORTREQUEST PERMISSIONS]
This cmdlet is available only in the Mailbox Import Export role, and by	default, that role isn't assigned to a role group. To use this cmdlet, you need to add the Mailbox Import Export role to a role group (for example, to the Organization Management role group). For more information, see the "Add a role to a role group" section in Manage role groups.
New-ManagementRoleAssignment –Role “Mailbox Import Export” –User Domain\User

When specifying the -PSTPath or relying on the users' AD Home Directory	value, the network share will need to have NTFS Read/Write permissions for the "Exchange Trusted Subsystem" Group.

[HTML TEMPLATE]
The HTML template file for the Large Item Notification e-mail is a fairly basic HTML file.  You can customize the HTML to suit your business or customer needs.  There are 3 variables in the HTML file that get replaced during the script.  The script will expect to find the HTML template file in the same directory as the script.  There are options for adding attachments if desired, but are not included.

#MAILBOXNAME#*
This will get replaced with the actual users display name.

#LARGEITEMFOLDER#*
This is the folder where all the item violations were moved to.

#LARGEITEMPATH#*
This is the location of their PST archive.

.LINK
Install the EWS Managed API 2.0:  http://www.microsoft.com/en-GB/download/details.aspx?id=35371

.LINK
Exchange 2007 Configure Impersonation:  http://msdn.microsoft.com/en-us/library/bb204095(v=exchg.80).aspx

.LINK
Exchange 2010/2013 Configure Impersonation:  http://msdn.microsoft.com/en-us/library/bb204095(v=exchg.140).aspx

.LINK
Exchange 2010 / 2013 Manage Role Groups:  http://technet.microsoft.com/en-us/library/jj657480(v=exchg.150).aspx
#>

Param 
(
	[Parameter(Position=0, Mandatory = $True, HelpMessage="Please provide the NETBIOS Domain Name for the Service Account")]
	[System.String]$ServiceAccountDomain,
	
	[Parameter(Position=1, Mandatory = $True, HelpMessage="Please provide the UserID for the Service Account")]
	[System.String]$ServiceAccountName,

	[Parameter(Position=2, Mandatory = $True, HelpMessage="Please provide the password for the Service Account")]
	[System.String]$ServicePassword,

	[Parameter(Position=3, Mandatory = $True, ValueFromPipeline = $True, HelpMessage="Enter the item size in Megabytes you want to search for in each mailbox")]
	[ValidateRange(1,999)]
	[System.Int32]$ItemSizeLimit,

    [Parameter(Mandatory = $False)]
    [System.String]$ImportCSV,

	[Parameter(Mandatory = $False)]
	[Switch]$CreateCSV,
	
	[Parameter(Mandatory = $False)]
	[Switch]$MoveLargeItems,
    
	[Parameter(Mandatory = $False)]
	[Switch]$ExportLargeItems,

    [Parameter(Mandatory = $False)]
    [Switch]$ExportOnly,

    [Parameter(Mandatory = $False)]
    [Switch]$LargeItemNotice,
    
    [Parameter(Mandatory = $False)]
    [System.String]$FolderName,

    [Parameter(Mandatory = $False)]
    [System.String]$PSTPath,

    [Parameter(Mandatory = $False)]
	[switch]$ArchiveCheck,
	
	[Parameter(Mandatory = $False)]
	[System.URI]$Uri
)

Function Write-Log($Text)
{
<# Write-Log Function
. [AUTHOR]
. Jon Lundfeldt
. 
. [VERSION]
. 1.0
#>
    Write-Verbose $Text
    $Log = New-Object System.IO.StreamWriter("$Logfile", [System.Text.Encoding]::UTF8)
    $Log.WriteLine("$(Get-Date),$Text")
    $Log.Close()
}

Function Log-ItemViolations
{
    Param
    (
        [System.String]$SMTPAddress,
        [System.String]$Subject,
        [System.String]$ItemClass,
        [System.String]$FolderDisplayName,
        [System.DateTime]$Created,
        [System.Int32]$Size
    )

    $Violations = New-Object System.Object;
    $Violations | Add-Member -Type NoteProperty -Name "SMTP" -Value $SMTPAddress;
    $Violations | Add-Member -Type NoteProperty -Name "Subject" -Value $Subject;
    $Violations | Add-Member -Type NoteProperty -Name "ItemClass" -Value $ItemClass;
    $Violations | Add-Member -Type NoteProperty -Name "FolderDisplayName" -Value $FolderDisplayName;
    $Violations | Add-Member -Type NoteProperty -Name "CreationTime" -Value $Created;
    $Violations | Add-Member -Type NoteProperty -Name "Size" -Value $Size;

    Return $Violations;
}

Function Get-Items ($ItemSizeLimit,$Folders,$Service)
{
    $LargeItemCount = 0;
    $ItemViolationLog = @();
	
	foreach ($Folder in $Folders)
	{
        $CurrentFolder = $Folder.DisplayName
        $Items = $Null;
        $PageSize = 1000
	    $Offset = 0
	    $MoreItemsAvailable = $True
	    $ItemCounter = 0
        
        Write-Log ("====>  MAILBOX: {0} | Started Processing folder: {1}" -f $MBX,$CurrentFolder)

	    Do
		{
	        $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize,$Offset,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
			$ItemFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.Note");
			
			Try
			{
                $Items = $Service.FindItems($Folder.Id, $ItemFilter, $ItemView);
                #$Items = $Service.FindItems($Folder.Id, $ItemView)
			}
			Catch
			{

                [System.Console]::ForegroundColor = "Magenta"
                Write-Host ("`n`r====>  MAILBOX: {0}, FOLDER: {1} | Unable to process an item" -f $MBX,$Folder.DisplayName);
                [System.Console]::ResetColor()

                Write-Log ("====>  ERROR  <==== | Processing an item in {0}" -f $CurrentFolder);
                Write-Log ("====>  {0}" -f $_.Exception.Message)

                $ErrorItemLog = Log-ItemViolations -SMTPAddress $Service.ImpersonatedUserId.Id -Created "01:00:00 AM" -Subject "ERROR: Unable to Process Item" -FolderDisplayName $Folder.DisplayName -Size 0 -ItemClass $null;
                $ItemViolationLog += $ErrorItemLog;
                Continue;
			}

            $TotalItems = $Items.TotalCount;

	        foreach ($Item in $Items.Items)
	        {
                Write-Progress -ParentId 1 -Activity "Processing Item: $ItemCounter" -status "Processing Folder: $CurrentFolder [$TotalItems]" -PercentComplete (($ItemCounter / $TotalItems)  * 100) -Verbose;

				[System.Int32] $Size = $Item.Size / 1000000; #Less than 1048576 to cater for buffer
				[System.String] $Subject = $Item.Subject;
				[System.String] $ItemClass = $Item.ItemClass;
				[System.DateTime] $Created = $Item.DateTimeCreated;
				
				if ($Size -gt $ItemSizeLimit)
				{ 
                    if ([System.String]::IsNullOrEmpty($Subject))
					{
						$Subject = "NULL"
                        $ItemViolation = Log-ItemViolations -SMTPAddress $Service.ImpersonatedUserId.Id -Created $Created -Subject $Subject -FolderDisplayName $Folder.DisplayName -Size $Size -ItemClass $ItemClass;
                        $ItemViolationLog += $ItemViolation;
					}
					else
					{
                        $ItemViolation = Log-ItemViolations -SMTPAddress $Service.ImpersonatedUserId.Id -Created $Created -Subject $Subject -FolderDisplayName $Folder.DisplayName -Size $Size -ItemClass $ItemClass;
                        $ItemViolationLog += $ItemViolation;
					}

                    if ($MoveLargeItems)
                    {
                        $FolderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MBXSMTPAddress);
                        $FolderRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$FolderID);
                        $FolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1);
                        $Filter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName);
                        $LargeItemFolder = $Service.FindFolders($FolderRoot.Id,$Filter,$FolderView);

                        Write-Log ("====>  MAILBOX: {0} | Moving Item {1} from folder {2} to folder {3}" -f $MBX,$Subject,$Folder.DisplayName,$FolderName); 
                        [void]$Item.Move($LargeItemFolder.Folders[0].Id);
                    }
                    $LargeItemCount += 1;
				}
				$ItemCounter += 1;
	        }
			
	    	if ($Items.MoreAvailable -eq $False)
	        {
                $MoreItemsAvailable = $False
                Write-Log ("====>  MAILBOX: {0} | Finished Processing folder: {1} ({2} Items)" -f $MBX,$CurrentFolder,$TotalItems);
            }

	        if ($MoreItemsAvailable)
	        {$Offset += $PageSize}
			
	    }while ($MoreItemsAvailable)

	}
    
    Write-Log ("====>  MAILBOX: {0} | Number of Large Items found: {1}" -f $MBX,$LargeItemCount);

    if ($ItemViolationLog)
    {
        Return $ItemViolationLog
    }
}

Function Create-Folder($FolderName,$Service)
{
    Try
    {    
        $FolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot);
        $FolderRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$FolderID);
        $View = new-object Microsoft.Exchange.WebServices.Data.FolderView(1);
        $Filter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName);
        $FindFolder = $Service.FindFolders($FolderRoot.Id,$Filter,$View);
        if($FindFolder.TotalCount -eq 0)
        {
            Write-Log ("====>  MAILBOX: {0} | Folder {1} was not found, creating the folder" -f $MBX,$FolderName); 
            
            $LargeItemFolder = new-object Microsoft.Exchange.WebServices.Data.Folder($Service);
            $LargeItemFolder.DisplayName = $FolderName;
            $LargeItemFolder.FolderClass = "IPF.Note"
            $LargeItemFolder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot);
        }
        else
        {
            Write-Log ("====>  WARNING  <==== | Large Item Folder already exists ({0})" -f $FolderName);

            [System.Console]::ForegroundColor = "Yellow"
            Write-Host ("`n`r====>  WARNING  <==== | Folder {0} already exists!" -f $FolderName);
            [System.Console]::ResetColor()
        }
    }
    Catch
    {
        [System.Console]::ForegroundColor = "Red"
        Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
        Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  MAILBOX: {0} | Unable to create or find the {1} Folder" -f $MBX,$FolderName);
        Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
        [System.Console]::ResetColor()

        Write-Log ("====>  ERROR  <==== | Unable to find or create Large Item Folder ({0})" -f $FolderName);
        Write-Log "====>  "$_.Exception.Message;
    }
}

Function Get-Folders($Service)
{
	[Microsoft.Exchange.WebServices.Data.FolderView]$View = New-Object Microsoft.Exchange.WebServices.Data.FolderView([System.Int32]::MaxValue);
	$View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly);
	$View.PropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName);
	$View.PropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount);
	$View.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;

    # Filter which only will look at folders that contain the IPM.Note Item Class.
    $ClassFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Note")
	
	if($archiveCheck)
	{
		Try
        {
            [Microsoft.Exchange.WebServices.Data.FindFoldersResults]$Folders = $Service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $ClassFilter,$View);
	        Write-Log ("====>  MAILBOX: {0} | Found {1} folders (ARCHIVE MAILBOX)" -f $MBX,$Folders.TotalCount);
        }
	    Catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException]
	    {
            $myError = @"

\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
\\\\\\\\\\
\\\\\\\\\\  Application Impersonation FAILED!
\\\\\\\\\\
\\\\\\\\\\  Exchange Version:  $($ExchangeVersion)
\\\\\\\\\\  Service Account:   $($ServiceAccountDomain)\$($ServiceAccountName)
\\\\\\\\\\
\\\\\\\\\\  Depending on where the mailboxes are hosted in your environment, you will need to properly
\\\\\\\\\\  assign the service account with application impersonation rights.  The script blocks below
\\\\\\\\\\  depicts the cmdlets required to assign these permissions.
\\\\\\\\\\
\\\\\\\\\\  Get-Help .\LargeItemChecks_2_1_1.ps1 -Full
\\\\\\\\\\
\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
"@

            $myScriptBlock = @"

------------------------------------------ [Script Blocks] ----------------------------------------------
[EXCHANGE 2010 (E14) / 2013 (E15) IMPERSONATION PERMISSIONS]
New-ManagementRoleAssignment –Name:<AssignmentName> –Role:ApplicationImpersonation –User:<ServiceAccount>
"@
    
            Write-Host $myError -ForegroundColor Red
            Write-Host $myScriptBlock -ForegroundColor Yellow
            
            Write-Log "====>  ERROR  <==== | Check the Service Account for correct EWS Impersonation rights"
            Write-Log "====>  "$_.Exception.Message
            Exit 1
        }
    }
	else
	{
        Try
        {
            [Microsoft.Exchange.WebServices.Data.FindFoldersResults]$Folders = $Service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $ClassFilter,$View);
        
            Write-Log ("====>  MAILBOX: {0} | Found {1} folders" -f $MBX,$Folders.TotalCount);
        }
	    Catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException]
	    {
            $myError = @"

\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
\\\\\\\\\\
\\\\\\\\\\  Application Impersonation FAILED!
\\\\\\\\\\
\\\\\\\\\\  Exchange Version:  $($ExchangeVersion)
\\\\\\\\\\  Service Account:   $($ServiceAccountDomain)\$($ServiceAccountName)
\\\\\\\\\\
\\\\\\\\\\  Depending on where the mailboxes are hosted in your environment, you will need to properly
\\\\\\\\\\  assign the service account with application impersonation rights.  The script blocks below
\\\\\\\\\\  depicts the cmdlets required to assign these permissions.
\\\\\\\\\\
\\\\\\\\\\  Get-Help .\LargeItemChecks_2_1_1.ps1 -Full
\\\\\\\\\\
\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
"@

            $myScriptBlock = @"

------------------------------------------ [Script Blocks] ----------------------------------------------
[EXCHANGE 2007 (E12) IMPERSONATION PERMISSIONS]
Get-ExchangeServer | where {`$_.IsClientAccessServer -eq `$TRUE} | ForEach-Object {Add-ADPermission
 -Identity `$_.distinguishedname -User (Get-User -Identity <ServiceAccount> | select-object).identity
 -ExtendedRight ms-Exch-EPI-Impersonation}

Get-MailboxDatabase | ForEach-Object {Add-ADPermission -Identity `$_.DistinguishedName
 -User <ServiceAccount> -ExtendedRights ms-Exch-EPI-May-Impersonate}

[EXCHANGE 2010 (E14) / 2013 (E15) IMPERSONATION PERMISSIONS]
New-ManagementRoleAssignment –Name:<AssignmentName> –Role:ApplicationImpersonation –User:<ServiceAccount>
"@
    
            Write-Host $myError -ForegroundColor Red
            Write-Host $myScriptBlock -ForegroundColor Yellow
            
            Write-Log "====>  ERROR  <==== | Check the Service Account for correct EWS Impersonation rights"
            Write-Log "====>  "$_.Exception.Message
            Exit 1
        }
	}
Return $Folders
}

Function Get-Impersonator 
{
	Param 
	(
        [System.String]$Identity,
        [System.String]$ImpersonatorAccountName,
		[System.String]$ImpersonatorAccountPassword,
		[System.String]$ImpersonaterAccountDomain,
        [System.URI]$Uri
	)
<# SSL Check / Bypass functionality
. [AUTHOR]
. Carter Shanklin
. 
. [URL]
. http://poshcode.org/624
#>
	$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
	$Compiler=$Provider.CreateCompiler()
	$Params=New-Object System.CodeDom.Compiler.CompilerParameters
	$Params.GenerateExecutable=$False
	$Params.GenerateInMemory=$True
	$Params.IncludeDebugInformation=$False
	$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

	$TASource=@'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
	        public class TrustAll : System.Net.ICertificatePolicy {
	            public TrustAll() { 
	            }
	            public bool CheckValidationResult(System.Net.ServicePoint sp,
	                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
	                System.Net.WebRequest req, int problem) {
	                return true;
	            }
	        }
	    }
'@ 
	$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
	$TAAssembly=$TAResults.CompiledAssembly
    
    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
	$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
	[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
	
	if($ExchangeVersion -eq "E14" -or $ExchangeVersion -eq "E15")
	{
		[Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1);
	}
	else
	{
		[Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1);
	}
	
	if ([String]::IsNullOrEmpty($Uri))
	{
        Try
        {
		    Write-Log ("====>  MAILBOX: {0} | Autodiscover in process" -f $MBX)
		    $Service.AutodiscoverUrl($Identity,{$True}) 
            Write-Log ("====>  MAILBOX: {0} | Using EWS URL: {1}" -f $MBX,$Service.Url)
        }
        Catch [System.Exception]
        {            
            [System.Console]::ForegroundColor = "Red"
            Write-Host "`n`r====>  MAILBOX: {0} | Unable to get the URI from the Autodiscover Service"
            [System.Console]::ResetColor()
            
            Write-Log ("====>  ERROR  <==== | Unable to get the URI from Autodiscover Service");
            Write-Log "====>  "$_.Exception.Message;
            Continue
        }
	}
	else
	{
		$Service.Url = $Uri.AbsoluteUri
	}

	$Service.Credentials = New-Object Net.NetworkCredential($ImpersonatorAccountName, $ImpersonatorAccountPassword, $ImpersonaterAccountDomain)
	$Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Identity)
	
    #Increase the timeout to cater for large mailboxes
	$Service.Timeout = 150000
	Return $Service
}

Function Get-ADAttributes($DistinguishedName)
{
    $ADFilter = "(&(objectCategory=person)(objectCategory=User))"; 

    $ADSearch = New-Object System.DirectoryServices.DirectorySearcher; 
    $ADSearch.PageSize = 15000;
    $ADSearch.Filter = $ADFilter;  
    $ADSearch.SearchRoot = "LDAP://$DistinguishedName"; 

    $ADResult = $ADSearch.FindOne();
    $ADUser = $ADResult.Properties;

    Return $ADUser;
}

Function Get-DateTime
{
	$DateTime = (Get-Date -Format MMddyy_HHmmss).ToString()
	Return $DateTime
}

Function Create-LargeItemReport
{
    Param
    (
        $CSVObject,
        $OutputFile = (".\LargeItemChecks_Results_{0}.csv") -f (Get-Date -Format MMddyy_HHmmss).ToString()
    )
 
    $File = $OutputFile.SubString(2)

    [System.Console]::ForegroundColor = "Green"
    Write-Host "`n`r////////////////////////////////////////////////////////////////////////////////////////////////////////"
    Write-Host ("//////////`n`r//////////  Creating the Large Item Report...`n`r//////////  {0}\{1}" -f $ScriptPath,$File)
    Write-Host "//////////`n`r////////////////////////////////////////////////////////////////////////////////////////////////////////"
    [System.Console]::ResetColor()
    
    Write-Log ("====>  Attempting to create CSV file: {0}\{1}" -f $ScriptPath,$File);
    $CSVObject | Export-Csv $OutputFile -NoTypeInformation
}

Function Export-LargeItems($PSTPath,$FolderName,$MBX,$HomeDir)
{
    if ($CentralizedExport)
    {
        if($Exchange2007)
        {
            Write-Log ("====>  MAILBOX: {0} | Using Exchange 2007 Export-Mailbox cmdlet" -f $MBX)
            Write-Log ("====>  MAILBOX: {0} | Starting Mailbox Export" -f $MBX);

            Try
            {
                $ExportItemViolation = Export-Mailbox -Identity $MBX -PSTFolderPath "$PSTPath\$SamAccountName-LargeItems.pst" -IncludeFolders "\$FolderName" -Confirm:$False;
            
                if($ExportItemViolation.statuscode -eq 0)
                {
		            Write-Log ("====>  MAILBOX: {0} | Export was completed!" -f $MBX);
		            
                    #================================================================================================================
                    # This Section will delete the LargeItemFolder.  Commented out since 2010 doesn't provide a status code like 2007
                    # Uncomment if you want to delete the LargeItemFolder
                    #================================================================================================================

                    #$LargeItemFolder = Get-LargeItemFolder -FolderName $FolderName -Service $Service
		            #$DeleteFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$LargeItemFolder.id)
		            #$DeleteFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)	
	    	        #Write-host "Folder deleted" -ForegroundColor Cyan
                }
                else
                {
                    Write-Log ("<====  WARNING  ====> | Mailbox Export process failed!" -f $MBX);
                }
            }
            Catch
            {
                [System.Console]::ForegroundColor = "Red"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  Mailbox Export Failed!`n`r\\\\\\\\\\")
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\`n`r"
                Write-Host ("ERROR: {0}" -f $_.Exception.Message)
                [System.Console]::ResetColor()

                Write-Log "====>  ERROR  <==== | Mailbox Export Failed!";
                Write-Log ("====>  {0}" -f $_.Exception.Message);
	            $Error.Clear()
            }
        }
        else
        {
            Write-Log ("====>  MAILBOX: {0} | Using Exchange 2010 New-MailboxExportRequest cmdlet" -f $MBX)
            Write-Log ("====>  MAILBOX: {0} | Starting Mailbox Export" -f $MBX);

            if (Get-MailboxExportRequest -Name $MBX)
            {
                [System.Console]::ForegroundColor = "Yellow"
                Write-Host "`n`r||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
                Write-Host ("||||||||||`n`r||||||||||  Found an existing export request for: {0}`n`r||||||||||  Removing the old export request...`n`r||||||||||" -f $MBX);
                Write-Host "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
                [System.Console]::ResetColor()

                Write-Log ("====>  MAILBOX: {0} | Found existing export request, removing the old request." -f $MBX);
                Get-MailboxExportRequest -Name $MBX | Remove-MailboxExportRequest -Confirm:$False;
            }

            New-MailboxExportRequest -Name $MBX -Mailbox $MBX -FilePath "$PSTPath\$SamAccountName-LargeItems.pst" -IncludeFolders $FolderName -Confirm:$False -ExcludeDumpster;
        }
    }
    else
    {
        if($Exchange2007)
        {
            Write-Log ("====>  MAILBOX: {0} | Using Exchange 2007 Export-Mailbox cmdlet" -f $MBX)
            Write-Log ("====>  MAILBOX: {0} | Starting Mailbox Export" -f $MBX);

            Try
            {
                $ExportItemViolation = Export-Mailbox -Identity $MBX -PSTFolderPath "$HomeDir\$SamAccountName-LargeItems.pst" -IncludeFolders $FolderName -Confirm:$False;
            
                if($ExportItemViolation.statuscode -eq 0)
                {
		            Write-Log ("====>  MAILBOX: {0} | Export was completed!" -f $MBX);
		            
                    #================================================================================================================
                    # This Section will delete the LargeItemFolder.  Commented out since 2010 doesn't provide a status code like 2007
                    # Uncomment if you want to delete the LargeItemFolder
                    #================================================================================================================

                    #$LargeItemFolder = Get-LargeItemFolder -FolderName $FolderName -Service $Service
		            #$DeleteFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$LargeItemFolder.id)
		            #$DeleteFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)	
	    	        #Write-host "Folder deleted" -ForegroundColor Cyan
                }
                else
                {
                    Write-Log ("<====  WARNING  ====> | Mailbox Export process failed!" -f $MBX);
                }
            }
            Catch
            {
                [System.Console]::ForegroundColor = "Red"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  Mailbox Export Failed!`n`r\\\\\\\\\\  ERROR: {0}" -f $_.Exception.Message)
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                [System.Console]::ResetColor()

                Write-Log "====>  ERROR  <==== | Mailbox Export Failed!";
                Write-Log ("====>  {0}" -f $_.Exception.Message);
	            $Error.Clear()
            }
        }
        else
        {
            Write-Log ("====>  MAILBOX: {0} | Using Exchange 2010 New-MailboxExportRequest cmdlet" -f $MBX)
            Write-Log ("====>  MAILBOX: {0} | Starting Mailbox Export" -f $MBX);

            if (Get-MailboxExportRequest -Name $MBX)
            {
                [System.Console]::ForegroundColor = "Yellow"
                Write-Host "`n`r||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
                Write-Host ("||||||||||`n`r||||||||||  Found an existing export request for: {0}`n`r||||||||||  Removing the old export request...`n`r||||||||||" -f $MBX);
                Write-Host "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
                [System.Console]::ResetColor()

                Write-Log ("====>  MAILBOX: {0} | Found existing export request, removing the old request." -f $MBX);
                Get-MailboxExportRequest -Name $MBX | Remove-MailboxExportRequest -Confirm:$False;
            }

            New-MailboxExportRequest -Name $MBX -Mailbox $MBX -FilePath "$HomeDir\$SamAccountName-LargeItems.pst" -IncludeFolders $FolderName -Confirm:$False -ExcludeDumpster;
        }
    }
}

Function Send-LargeItemNotice($Mailbox,$Recipient,$PSTPath,$HomeDir,$LargeItemFolder)
{
    #Set this to the e-mail address / mailbox you want to "Send As"
    $O365Admin = "Office 365 Admin <office365@humongousinsurance.com>"
    
    Write-Log ("====>  MAILBOX: {0} | Building Large Item Notification template" -f $Mailbox)
    $MessageBody = $null
    $MessageBody = Get-Content .\LargeItemNotice.htm

    [System.String]$TempName = $MessageBody | ? {$_ -like "*MAILBOXNAME*"}
    [System.String]$TempItemFolder = $MessageBody | ? {$_ -like "*LARGEITEMFOLDER*"}
    [System.String]$TempLocation = $MessageBody | ? {$_ -like "*LARGEITEMPATH*"}

    $Name = $TempName -replace "#MAILBOXNAME#",$Mailbox
    $ItemFolder = $TempItemFolder -replace "#LARGEITEMFOLDER#",$LargeItemFolder

    if ($CentralizedExport)
    {
        $Location = $TempLocation -replace "#LARGEITEMPATH#","$PSTPath\$SamAccountName-LargeItems.pst"
    }
    else
    {
        $Location = $TempLocation -replace "#LARGEITEMPATH#","$HomeDir\$SamAccountName-LargeItems.pst"
    }

    $MessageBody = $MessageBody.Replace($TempName,$Name)
    $MessageBody = $MessageBody.Replace($TempItemFolder,$ItemFolder)
    $MessageBody = $MessageBody.Replace($TempLocation,$Location)

    $MessageBody = $MessageBody | Out-String

    #Add, Remove, Change the attachments based on your needs and file locations
    #$Attachment1 = ".\image001.jpg"
    #$Attachment2 = ".\image002.gif"

    Write-Log ("====>  MAILBOX: {0} | Sending Large Item Notification e-mail" -f $Mailbox)

    Send-MailMessage -Attachments -Body $MessageBody -BodyAsHtml -From $O365Admin -SmtpServer $SMTPServer -Subject "[NOTIFICATION] Large E-mails Detected (> 25MB)" -To $Recipient -Bcc ""

}

Function Process-ALLMailboxes()
{
	Write-Log "====>  Enumerating ALL mailboxes....please be patient, this may take a while."
	$CSVObject = @();
	$MBXCounter = 1;

	$Mailboxes = Get-Mailbox -ResultSize "Unlimited"
    $TotalMailboxes = ($Mailboxes | Measure-Object).Count

    Write-Log ("====>  Total number of mailboxes to process: {0}" -f $TotalMailboxes)

    foreach ($MBX in $Mailboxes)
    {
        Write-Progress -Id 1 -Activity "Searching Mailboxes: $MBXCounter of $TotalMailboxes" -status "Processing Mailbox: $MBX" -PercentComplete (($MBXCounter / $TotalMailboxes)  * 100);
		
        $Error.Clear()
		Try
		{
			Write-Log ("====>  MAILBOX: {0} of {1}" -f $MBXCounter, $TotalMailboxes);

            [System.String]$DN = $MBX.DistinguishedName;
            [System.String]$DisplayName = $MBX.DisplayName;
            [System.String]$MBXSMTPAddress = $MBX.PrimarySMTPAddress
            [System.String]$SamAccountName = $MBX.SamAccountName

            Write-Log ("====>  MAILBOX: {0} | Collecting Active Directory Attributes" -f $MBX);
            $ADInfo = Get-ADAttributes -DistinguishedName $DN;
            $HomeDir = $ADInfo.homedirectory;

            if ([System.String]::IsNullOrEmpty($HomeDir))
            {
                Write-Log ("====>  MAILBOX: {0} | Home Directory not found" -f $MBX)
            }
            else
            {
                Write-Log ("====>  MAILBOX: {0} | Home Directory set to '{1}'" -f $MBX,$HomeDir)
            }

            if ([System.String]::IsNullOrEmpty($Uri))
		    {
                Write-Log ("====>  MAILBOX: {0} | Connecting to EWS as {1}" -f $MBX,$ServiceAccountName);
			    $Service = Get-Impersonator -Identity $MBXSMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ImpersonaterAccountDomain $ServiceAccountDomain;
		    }
		    else
		    {
			    Write-Log ("====>  MAILBOX: {0} | Connecting to EWS as {1}" -f $MBX,$ServiceAccountName);
                $Service = Get-Impersonator -Identity $MBXSMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ImpersonaterAccountDomain $ServiceAccountDomain -Uri $Uri;
		    }
            
            Write-Log ("====>  MAILBOX: {0} | Enumerating Folders" -f $MBX);
            $Folders = Get-Folders -Service $Service;

            if (!$EXPORTONLY)
            {
                if ($MoveLargeItems)
                {
                    Write-Log ("====>  MAILBOX: {0} | Validating / Creating the Large Item Folder" -f $MBX);
                    Create-Folder -FolderName $FolderName -Service $Service
                }

                Write-Log ("====>  MAILBOX: {0} | Enumerating Items" -f $MBX);
                $ItemViolationLog = Get-Items -ItemSizeLimit $ItemSizeLimit -Folders $Folders -Service $Service;

                if ($ItemViolationLog)
                {
                    Write-Log ("====>  MAILBOX: {0} | Writing Item Violations to CSVObject" -f $MBX)
                    $CSVObject += $ItemViolationLog
                }
            }

            if ($ExportLargeItems)
            {
                Write-Log ("====>  MAILBOX: {0} | Exporting Large Item Folder into PST" -f $MBX)
                Export-LargeItems -PSTPath $PSTPath -FolderName $FolderName -MBX $MBX -HomeDir $HomeDir

                if ($LargeItemNotice)
                {
                    Write-Log ("====>  MAILBOX: {0} | Generating Large Item Notification" -f $MBX)
                    Send-LargeItemNotice -Mailbox $MBX -Recipient $MBXSMTPAddress -PSTPath $PSTPath -HomeDir $HomeDir -LargeItemFolder $FolderName
                }
            }

            $MBXCounter += 1;
        }
	    Catch [System.Exception]
	    {
            $ErrorMsg = $_.ToString()

            if ($ErrorMsg.contains("Autodiscover"))
		    {
			    [System.Console]::ForegroundColor = "Red"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  Check if mailbox is hosted in the current Exchange organization`n`r\\\\\\\\\\  ERROR: {0}`n`r\\\\\\\\\\" -f $_.Exception.Message)
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                [System.Console]::ResetColor()

                Write-Log ("====>  ERROR  <==== | Check if mailbox ({0}) is hosted in the current Exchange organization" -f $MBX)
                Write-Log ("====>  {0}" -f $_.Exception.Message)
            }
		    elseif ($ErrorMsg.contains("(401)")) #Unauthorized
		    {
			    [System.Console]::ForegroundColor = "Red"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  Check service account credentials`n`r\\\\\\\\\\  ERROR: {0}`n`r\\\\\\\\\\" -f $_.Exception.Message)
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                [System.Console]::ResetColor()

                Write-Log ("====>  ERROR  <==== | Check service account credentials on mailbox: {0}" -f $MBX)
                Write-Log ("====>  {0}" -f $_.Exception.Message)
		    }
            $Error.Clear()
        }	
    }
    Return $CSVObject;
}

Function Process-CSVMailboxes()
{
	$Mailboxes = Import-Csv $ImportCSV 
	$TotalMailboxes = ($Mailboxes | Measure-Object).Count
	
    Write-Log ("====>  Total number of mailboxes from CSV: {0}" -f $TotalMailboxes);
	$CSVObject = @();
	$MBXCounter = 1;

    foreach ($MBX in $Mailboxes)
    {
        $MBX = Get-Mailbox -Identity $MBX.PrimarySMTPAddress

        Write-Progress -Id 1 -Activity "Searching Mailboxes: $MBXCounter of $TotalMailboxes" -status "Processing Mailbox: $MBX" -PercentComplete (($MBXCounter / $TotalMailboxes)  * 100);
		
        $Error.Clear()
		Try
		{
			Write-Log ("====>  MAILBOX: {0} of {1}" -f $MBXCounter, $TotalMailboxes);

            [System.String]$DN = $MBX.DistinguishedName;
            [System.String]$DisplayName = $MBX.DisplayName;
            [System.String]$MBXSMTPAddress = $MBX.PrimarySMTPAddress
            [System.String]$SamAccountName = $MBX.SamAccountName

            Write-Log ("====>  MAILBOX: {0} | Collecting Active Directory Attributes" -f $MBX);
            $ADInfo = Get-ADAttributes -DistinguishedName $DN;
            [System.String]$HomeDir = $ADInfo.homedirectory

            if ([System.String]::IsNullOrEmpty($HomeDir))
            {
                Write-Log ("====>  MAILBOX: {0} | Home Directory not found" -f $MBX)
            }
            else
            {
                Write-Log ("====>  MAILBOX: {0} | Home Directory set to '{1}'" -f $MBX,$HomeDir)
            }

            if ([System.String]::IsNullOrEmpty($Uri))
		    {
                Write-Log ("====>  MAILBOX: {0} | Connecting to EWS as {1}" -f $MBX,$ServiceAccountName);
			    $Service = Get-Impersonator -Identity $MBXSMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ImpersonaterAccountDomain $ServiceAccountDomain;
		    }
		    else
		    {
			    Write-Log ("====>  MAILBOX: {0} | Connecting to EWS as {1}" -f $MBX,$ServiceAccountName);
                $Service = Get-Impersonator -Identity $MBXSMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ImpersonaterAccountDomain $ServiceAccountDomain -Uri $Uri;
		    }
            
            Write-Log ("====>  MAILBOX: {0} | Enumerating Folders" -f $MBX);
            $Folders = Get-Folders -Service $Service;

            if (!$EXPORTONLY)
            {
                if ($MoveLargeItems)
                {
                    Write-Log ("====>  MAILBOX: {0} | Validating / Creating the Large Item Folder" -f $MBX);
                    Create-Folder -FolderName $FolderName -Service $Service
                }

                Write-Log ("====>  MAILBOX: {0} | Enumerating Items" -f $MBX);
                $ItemViolationLog = Get-Items -ItemSizeLimit $ItemSizeLimit -Folders $Folders -Service $Service;

                if ($ItemViolationLog)
                {
                    Write-Log ("====>  MAILBOX: {0} | Writing Item Violations to CSVObject" -f $MBX)
                    $CSVObject += $ItemViolationLog
                }
            }

            if ($ExportLargeItems)
            {
                Write-Log ("====>  MAILBOX: {0} | Exporting Large Item Folder into PST" -f $MBX)
                Export-LargeItems -PSTPath $PSTPath -FolderName $FolderName -MBX $MBX -HomeDir $HomeDir

                if ($LargeItemNotice)
                {
                    Write-Log ("====>  MAILBOX: {0} | Generating Large Item Notification" -f $MBX)
                    Send-LargeItemNotice -Mailbox $MBX -Recipient $MBXSMTPAddress -PSTPath $PSTPath -HomeDir $HomeDir -LargeItemFolder $FolderName
                }
            }

            $MBXCounter += 1;
        }
	    Catch [System.Exception]
	    {
            $ErrorMsg = $_.ToString()

            if ($ErrorMsg.contains("Autodiscover"))
		    {
			    [System.Console]::ForegroundColor = "Red"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  Check if mailbox is hosted in the current Exchange organization`n`r\\\\\\\\\\  ERROR: {0}`n`r\\\\\\\\\\" -f $_.Exception.Message)
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                [System.Console]::ResetColor()

                Write-Log ("====>  ERROR  <==== | Check if mailbox ({0}) is hosted in the current Exchange organization" -f $MBX)
                Write-Log ("====>  {0}" -f $_.Exception.Message)
            }
		    elseif ($ErrorMsg.contains("(401)")) #Unauthorized
		    {
			    [System.Console]::ForegroundColor = "Red"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  Check service account credentials`n`r\\\\\\\\\\  ERROR: {0}`n`r\\\\\\\\\\" -f $_.Exception.Message)
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                [System.Console]::ResetColor()

                Write-Log ("====>  ERROR  <==== | Check service account credentials on mailbox: {0}" -f $MBX)
                Write-Log ("====>  {0}" -f $_.Exception.Message)
		    }
            $Error.Clear()
        }	
    }
    Return $CSVObject;
}

Function Process-Mailbox($Mailbox)
{
	$TotalMailboxes = ($Mailbox | Measure-Object).Count
	
    Write-Log ("====>  Total number of mailboxes to process: {0}" -f $TotalMailboxes);
	$CSVObject = @();
    $MBXCounter = 1;

    $MBX = Get-Mailbox -Identity $Mailbox

    $Error.Clear()
	Try
	{
		Write-Log ("====>  MAILBOX: {0} of {1}" -f $MBXCounter, $TotalMailboxes);

        [System.String]$DN = $MBX.DistinguishedName
        [System.String]$MBXSMTPAddress = $MBX.PrimarySMTPAddress
        [System.String]$SamAccountName = $MBX.SamAccountName

        Write-Log ("====>  MAILBOX: {0} | Collecting Active Directory Attributes" -f $MBX);
        $ADInfo = Get-ADAttributes -DistinguishedName $DN;
        $HomeDir = $ADInfo.homedirectory;

        if ([System.String]::IsNullOrEmpty($HomeDir))
        {
            Write-Log ("====>  MAILBOX: {0} | Home Directory not found" -f $MBX)
        }
        else
        {
            Write-Log ("====>  MAILBOX: {0} | Home Directory set to '{1}'" -f $MBX,$HomeDir)
        }

        if ([System.String]::IsNullOrEmpty($Uri))
		{
            Write-Log ("====>  MAILBOX: {0} | Connecting to EWS as {1}" -f $MBX,$ServiceAccountName);
			$Service = Get-Impersonator -Identity $MBXSMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ImpersonaterAccountDomain $ServiceAccountDomain;
		}
		else
		{
			Write-Log ("====>  MAILBOX: {0} | Connecting to EWS as {1}" -f $MBX,$ServiceAccountName);
            $Service = Get-Impersonator -Identity $MBXSMTPAddress -ImpersonatorAccountName $ServiceAccountName -ImpersonatorAccountPassword $ServicePassword -ImpersonaterAccountDomain $ServiceAccountDomain -Uri $Uri;
		}
            
        Write-Log ("====>  MAILBOX: {0} | Enumerating Folders" -f $MBX);
        $Folders = Get-Folders -Service $Service;

        if (!$EXPORTONLY)
        {
            if ($MoveLargeItems)
            {
                Write-Log ("====>  MAILBOX: {0} | Validating / Creating the Large Item Folder" -f $MBX);
                Create-Folder -FolderName $FolderName -Service $Service
            }

            Write-Log ("====>  MAILBOX: {0} | Enumerating Items" -f $MBX);
            $ItemViolationLog = Get-Items -ItemSizeLimit $ItemSizeLimit -Folders $Folders -Service $Service;

            if ($ItemViolationLog)
            {
                Write-Log ("====>  MAILBOX: {0} | Writing Item Violations to CSVObject" -f $MBX)
                $CSVObject += $ItemViolationLog
            }
        }

        if ($ExportLargeItems)
        {
            Write-Log ("====>  MAILBOX: {0} | Exporting Large Item Folder into PST" -f $MBX)
            Export-LargeItems -PSTPath $PSTPath -FolderName $FolderName -MBX $MBX -HomeDir $HomeDir
                
            if ($LargeItemNotice)
            {
                Write-Log ("====>  MAILBOX: {0} | Generating Large Item Notification" -f $MBX)
                Send-LargeItemNotice -Mailbox $MBX -Recipient $MBXSMTPAddress -PSTPath $PSTPath -HomeDir $HomeDir -LargeItemFolder $FolderName
            }
        }

        $MBXCounter += 1;
    }
	Catch [System.Exception]
	{
        $ErrorMsg = $_.ToString()

        if ($ErrorMsg.contains("Autodiscover"))
		{
			[System.Console]::ForegroundColor = "Red"
            Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  Check if mailbox is hosted in the current Exchange organization`n`r\\\\\\\\\\  ERROR: {0}`n`r\\\\\\\\\\" -f $_.Exception.Message)
            Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            [System.Console]::ResetColor()

            Write-Log ("====>  ERROR  <==== | Check if mailbox ({0}) is hosted in the current Exchange organization" -f $MBX)
            Write-Log ("====>  {0}" -f $_.Exception.Message)
        }
		elseif ($ErrorMsg.contains("(401)")) #Unauthorized
		{
			[System.Console]::ForegroundColor = "Red"
            Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  Check service account credentials`n`r\\\\\\\\\\  ERROR: {0}`n`r\\\\\\\\\\" -f $_.Exception.Message)
            Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            [System.Console]::ResetColor()

            Write-Log ("====>  ERROR  <==== | Check service account credentials on mailbox: {0}" -f $MBX)
            Write-Log ("====>  {0}" -f $_.Exception.Message)
		}
        $Error.Clear()
    }
    Return $CSVObject;
}

Clear-Host
$Error.Clear()

$myTitle = @"
\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
\\\\\\\\\\
\\\\\\\\\\  Title:    Office 365 Large Item Script
\\\\\\\\\\  Purpose:  Find items in mailboxes over $($ItemSizeLimit) MB and perform an action
\\\\\\\\\\  Actions:  -CreateCSV, -MoveLargeItems, -ExportLargeItems
\\\\\\\\\\  Script:   LargeItemChecks_2_1_1.ps1
\\\\\\\\\\
\\\\\\\\\\  Help:    Get-Help .\LargeItemChecks_2_1_1.ps1 -Full
\\\\\\\\\\
\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
"@

[System.Console]::ForegroundColor = "White"
Write-Host $myTitle
[System.Console]::ResetColor()

Try
{
	$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
	$CurrentDateTime = Get-DateTime
    $LogFile = "$ScriptPath\LargeItemChecks_ScriptLog_$CurrentDateTime.log"

	Write-Log "====>  Determining if Exchange Web Services 2.0 is installed"
    $CheckWebServicesInstalled = Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Exchange\Web Services\2.0'
	
	if (!$CheckWebServicesInstalled)
	{
        [System.Console]::ForegroundColor = "Yellow"
        Write-Host "`n`r||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
        Write-Host "||||||||||`n`r||||||||||  Exchange Web Services API cannot be found or is not installed.`n`r||||||||||  Download EWS API 2.0 here: http://www.microsoft.com/en-us/download/details.aspx?id=35371`n`r||||||||||"
        Write-Host "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||";
        [System.Console]::ResetColor()

		Write-Log "====>  WARNING  <==== | Exchange Web Services API 2.0 was not found, Script will now exit"
        Exit 1
	}
	
    Write-Log "====>  Looking for Microsoft Exchange Server Management Tools"
    $RootRegPath = 'HKLM:\SOFTWARE\Microsoft'
	
	if (Test-Path -Path $RootRegPath'\Exchange\v8.0\AdminTools') 
	{
        $SnapinToLoad = 'Microsoft.Exchange.Management.PowerShell.Admin'
        [System.String]$ExchangeVersion = "E12"
        Write-Log ("====>  PowerShell Snap-in to Load: {0}" -f $SnapinToLoad)
        Write-Log ("====>  Setting Exchange Version to: {0}" -f $ExchangeVersion)
    }
	elseif (Test-Path -Path $RootRegPath'\ExchangeServer\v14\AdminTools')
	{
        $SnapinToLoad = 'Microsoft.Exchange.Management.PowerShell.E2010'
        [System.String]$ExchangeVersion = "E14"
        Write-Log ("====>  PowerShell Snap-in to Load: {0}" -f $SnapinToLoad)
        Write-Log ("====>  Setting Exchange Version to: {0}" -f $ExchangeVersion)
    }
	elseif (Test-Path -Path $RootRegPath'\ExchangeServer\v15\AdminTools') 
	{
        $SnapinToLoad = 'Microsoft.Exchange.Management.PowerShell.SnapIn'
        [System.String]$ExchangeVersion = "E15"
        Write-Log ("====>  PowerShell Snap-in to Load: {0}" -f $SnapinToLoad)
        Write-Log ("====>  Setting Exchange Version to: {0}" -f $ExchangeVersion)
    }
	else
	{
        [System.Console]::ForegroundColor = "Yellow"
        Write-Host "`n`r||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
        Write-Host "||||||||||`n`r||||||||||  Microsoft Exchange Server Management Tools cannot be found or are not installed.`n`r||||||||||"
        Write-Host "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||";
        [System.Console]::ResetColor()

        Write-Log "====>  WARNING  <==== | Microsoft Exchange Server Management Tools cannot be found or are not installed, Script will now exit"
        Exit 1
    }

    $Exchange2007 = $False

    if ($ExchangeVersion -eq "E12")
    {$Exchange2007 = $true}
			
	if (! (Get-PSSnapin $SnapinToLoad -ErrorAction:SilentlyContinue) )
	{
        Write-Log "====>  Loading Microsoft Exchange Server Management Snapin..."

        $ExchangeDir = Test-Path -Path $env:ExchangeInstallPath

		Add-PSSnapin $SnapinToLoad

        if ($ExchangeDir -and (!$Exchange2007))
        {
            Write-Log "====>  Loading Microsoft Exchange Management Shell..."
            . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
            Connect-ExchangeServer -Auto
        }
	}
	
	Write-Log "====>  Loading Microsoft.Exchange.WebServices.dll"
	[System.String]$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
	[System.Void][Reflection.Assembly]::LoadFile($dllpath)
	
	$Mailboxes = $null
    
    Write-Log ("====>  Validating ACTION based parameters");

    if ($CreateCSV)
    { Write-Log ("====>  ACTION: -CreateCSV: {0}  <====" -f $CreateCSV)}
    else
    {Write-Log ("====>  ACTION: -CreateCSV: {0}  <====" -f $CreateCSV)}

    if ($MoveLargeItems)
    {
        Write-Log ("====>  ACTION: -MoveLargeItems: {0}  <====" -f $MoveLargeItems)
        Write-Log ("====>  Checking for required variables")

        if ([System.String]::IsNullOrEmpty($FolderName))
        {
            Write-Log ("====>  WARNING  <==== | Missing -FolderName parameter");

            [System.Console]::ForegroundColor = "Cyan"
		    Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            $CheckFolderName = Read-Host "The -MoveLargeItems and -ExportLargeItems parameter requires the -FolderName parameter.`n`rDo you want to provide a value for -FolderName (y/n)?"
            [System.Console]::ResetColor()

            if ($CheckFolderName.ToUpper() -eq "Y")
            {
                [System.Console]::ForegroundColor = "Cyan"
			    Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                $FolderName = Read-Host "Type the Name of the Large Item Folder (NO SPACES)"
                [System.Console]::ResetColor()

                Write-Log ("====>  Large Item Folder set to: '{0}'" -f $FolderName)
            }
            else
            {
                [System.Console]::ForegroundColor = "Red"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host "\\\\\\\\\\`n`r\\\\\\\\\\  The parameter -FolderName was not provided, the script is exiting...`n`r\\\\\\\\\\"
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                [System.Console]::ResetColor()

                Write-Log "====>  ERROR  <==== | The parameter -FolderName was not provided, the Script will now exit!"
                Exit 1;
            }
        }
        else
        {Write-Log ("====>  Large Item Folder set to: '{0}'" -f $FolderName)}
    }
    else
    {Write-Log ("====>  ACTION: -MoveLargeItems: {0}  <====" -f $MoveLargeItems)}

    if ($ExportLargeItems)
    {
        Write-Log ("====>  ACTION: -ExportLargeItems: {0}  <====" -f $ExportLargeItems)
        Write-Log ("====>  Checking for required variables...")

        if ([System.String]::IsNullOrEmpty($FolderName))
        {
            Write-Log ("====>  WARNING  <==== | Missing -FolderName parameter");

            [System.Console]::ForegroundColor = "Cyan"
		    Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            $CheckFolderName = Read-Host "The -MoveLargeItems and -ExportLargeItems parameter requires the -FolderName parameter.`n`rDo you want to provide a value for -FolderName (y/n)?"
            [System.Console]::ResetColor()

            if ($CheckFolderName.ToUpper() -eq "Y")
            {
                [System.Console]::ForegroundColor = "Cyan"
			    Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                $FolderName = Read-Host "Type the Name of the Large Item Folder (NO SPACES)"
                [System.Console]::ResetColor()

                Write-Log ("====>  Large Item Folder set to: '{0}'" -f $FolderName)
            }
            else
            {
                [System.Console]::ForegroundColor = "Red"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host "\\\\\\\\\\`n`r\\\\\\\\\\  The parameter -FolderName was not provided, the script is exiting...`n`r\\\\\\\\\\"
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                [System.Console]::ResetColor()

                Write-Log "====>  ERROR  <==== | The parameter -FolderName was not provided, the Script will now exit!"
                Exit 1;
            }
        }
        else
        {Write-Log ("====>  Large Item Folder set to: '{0}'" -f $FolderName)}

        if ([System.String]::IsNullOrEmpty($PSTPath))
        {
            $Title = "Mailbox Export Destination"
            $Message =@"
Where do you want the Export function to store the PST files?

- Centralized Network Share (Recommended)
This option will place all PST files created by the Export function
in a central location.  If you choose this option, you will be
prompted to provide a value for the -PST Parameter.

- Users' Home Directory
This option will query AD for the users' home directory attribute.
If the attribute has a value it will attempt to create the PST
in that location.  If the attribute doesn't have a value, then
the Export function will not work.  This option should only be
used if all users have a valid home directory.

To skip this prompt, you can use the -PSTPath parameter to force
centralized export!
"@

            $Centralized = New-Object System.Management.Automation.Host.ChoiceDescription "&Centralized Network Share?", `
                "Stores all the PST files in a centralize network share."

            $UserBased = New-Object System.Management.Automation.Host.ChoiceDescription "&Users' Home Directory?", `
                "Stores the PST in the Home Directory for the User."

            $Options = [System.Management.Automation.Host.ChoiceDescription[]]($Centralized, $UserBased)

            $Result = $Host.UI.PromptForChoice($Title, $Message, $Options, 0) 

            Switch ($Result)
                {
                    0   {
                        $CentralizedExport = $True

                        Write-Log ("====>  -PSTPath parameter is empty...")
		
                        [System.Console]::ForegroundColor = "Cyan"
                        Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                        $PSTPath = Read-Host "Please provide the full UNC path for the Centralized Export location (\\server\share)"
                        [System.Console]::ResetColor()

                        Write-Log ("====>  -PSTPath parameter was set to {0}" -f $PSTPath);
                        }

                    1 {$CentralizedExport = $false}
                }
        }
        else
        {
            $CentralizedExport = $True
            Write-Log ("====>  Centralize Export:  {0}" -f $CentralizedExport)
        }

        Write-Log ("====>  Centralized Export:  {0}" -f $CentralizedExport)

        if ($LargeItemNotice)
        {
		    
            [System.Console]::ForegroundColor = "Cyan"
		    Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            $CheckSMTPServer = Read-Host "The -LargeItemNotice parameter requires the FQDN of an SMTP Server. Do you want to continue (y/n)?"
            [System.Console]::ResetColor()

            if ($CheckSMTPServer.ToUpper() -eq "Y")
            {
                [System.Console]::ForegroundColor = "Cyan"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                $SMTPServer = Read-Host "Please provide the FQDN of your SMTP Server"
		        [System.Console]::ResetColor()

                Write-Log ("====>  SMTP Server set to: '{0}'" -f $SMTPServer)
            }
            else
            {
                [System.Console]::ForegroundColor = "Red"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host "\\\\\\\\\\`n`r\\\\\\\\\\  No value provided for the SMTP Server, the script will now exit...`n`r\\\\\\\\\\"
                Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                [System.Console]::ResetColor()

                Write-Log "====>  ERROR  <==== | The SMTP Server was not provided, the Script will now exit!"
                Exit 1;
            }

            Write-Log ("====>  Large Item Notification: {0}" -f $LargeItemNotice)
        }
        else
        {Write-Log ("====>  Large Item Notification: {0}" -f $LargeItemNotice)}
    }
    else
    {Write-Log ("====>  ACTION: -ExportLargeItems: {0}  <====" -f $ExportLargeItems)}

    if (!$ExportOnly)
    {
        if ((!$MoveLargeItems -and !$ExportLargeItems) -and ![System.String]::IsNullOrEmpty($FolderName))
        {
            [System.Console]::ForegroundColor = "Red"
            Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            Write-Host "\\\\\\\\\\`n`r\\\\\\\\\\  Only specify the -FolderName parameter in conjunction with`n`r\\\\\\\\\\ the -MoveLargeItems and -ExportLargeItems parameters, the script is exiting...`n`r\\\\\\\\\\"
            Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            [System.Console]::ResetColor()

            Write-Log "====>  The parameter -FolderName was used incorrectly, the Script will now exit!"
            Exit 1;
        }

    }
    elseif (($CreateCSV -or $MoveLargeItems) -and $EXPORTONLY)
    {
        [System.Console]::ForegroundColor = "Red"
        Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
        Write-Host "\\\\\\\\\\`n`r\\\\\\\\\\  The parameter -EXPORTONLY cannot be used in conjunction with the`n`r\\\\\\\\\\  -CreateCSV or -MoveLargeItems parameters, the script is exiting...`n`r\\\\\\\\\\"
        Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
        [System.Console]::ResetColor()

        Write-Log "====>  The parameter -ExportOnly was used incorrectly, the Script will now exit!"
        Exit 1;
    }
    else
    {
        [System.Console]::ForegroundColor = "Yellow"
        Write-Host "`n`r||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
        Write-Host ("||||||||||`n`r||||||||||  :: WARNING :: YOU HAVE SELECTED -EXPORTONLY!`n`r||||||||||  IF YOU HAVE NOT PREVIOUSLY RUN THIS SCRIPT WITH THE -MOVELARGEITEMS`n`r||||||||||  PARAMETER, THEN YOUR EXPORT MAY NOT CONTAIN ANY ITEMS.`n`r||||||||||" -f $FolderName)
        Write-Host "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||";
        [System.Console]::ResetColor() 

        Write-Log ("====>  ACTION: -ExportOnly: {0}  <====" -f $ExportOnly)
    }
    
    Write-Log "====>  Determining which Exchange Mailboxes to process"
	
    if (!$ImportCSV)
	{
		[System.Console]::ForegroundColor = "Cyan"
        Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
		$CheckCSV = Read-host "The -ImportCSV parameter has not been specified.  The script wil process ALL Mailboxes!`n`rAre you sure? (y/n)"
		[System.Console]::ResetColor()
		if ($CheckCSV.toUpper() -eq 'Y')
		{	
            Write-Log ("====>  Mailbox Selection: ALL Mailboxes");
			$CSVObject = Process-ALLMailboxes
		}
		else
		{
            [System.Console]::ForegroundColor = "Cyan"
            write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
            $CheckSingleMBX = Read-Host "Aborted scanning ALL Mailboxes!  Do you want to process a single Mailbox? (y/n)"
            [System.Console]::ResetColor()
            if ($CheckSingleMBX.ToUpper() -eq 'Y')
            {
                [System.Console]::ForegroundColor = "Cyan"
                Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                $Identity = Read-Host "Enter the name of the Mailbox you want to process (PrimarySMTPAddress)"
                [System.Console]::ResetColor()

                Write-Log ("====>  Mailbox Selection: Single Mailbox");
                $CSVObject = Process-Mailbox -Mailbox $Identity
                
            }
            else
            {
                [System.Console]::ForegroundColor = "Yellow"
                Write-Host "`n`r||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
                Write-Host "||||||||||`n`r||||||||||  No Mailboxes selected to process, the script is exiting...`n`r||||||||||"
                Write-Host "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||";
                [System.Console]::ResetColor()

                Write-Log "====>  WARNING  <==== | No Mailboxes selected to process, the script is exiting..."
			    Exit 1
            }
		}
	}
	else
	{
        Write-Log ("====>  Mailbox Selection: Mailboxes from CSV Import file ($ImportCSV)");
		$CSVObject = Process-CSVMailboxes
	}

    if ($CreateCSV)
    {
        if ($CSVObject)
        {
            $OutputFile = (".\LargeItemChecks_Results_{0}.csv") -f (Get-Date -Format MMddyy_HHmmss).ToString()
 
            $File = $OutputFile.SubString(2)

            [System.Console]::ForegroundColor = "Green"
            Write-Host "`n`r////////////////////////////////////////////////////////////////////////////////////////////////////////"
            Write-Host ("//////////`n`r//////////  Creating the Large Item Report...`n`r//////////  {0}\{1}" -f $ScriptPath,$File)
            Write-Host "//////////`n`r////////////////////////////////////////////////////////////////////////////////////////////////////////"
            [System.Console]::ResetColor()
    
            Write-Log ("====>  Attempting to create CSV file: {0}\{1}" -f $ScriptPath,$File);
            $CSVObject | Export-Csv $OutputFile -NoTypeInformation
        }
        else
        {
            [System.Console]::ForegroundColor = "Yellow"
            Write-Host "`n`r||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
            Write-Host "||||||||||`n`r||||||||||  Unable to create CSV file, no data collected...`n`r||||||||||"
            Write-Host "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||";
            [System.Console]::ResetColor()

            Write-Log "====>  WARNING  <==== | No CSV file created, the `$CSVObject was empty or `$null..."   
        }
        #Create-LargeItemReport -CSVObject $CSVObject
    }

    if (!$CreateCSV -and !$MoveLargeItems -and !$ExportLargeItems)
    {
        [System.Console]::ForegroundColor = "Green"
        Write-Host "`n`r////////////////////////////////////////////////////////////////////////////////////////////////////////"
        Write-Host ("//////////`n`r//////////  NO PARAMETERS DEFINED (-CreateCSV, -MoveLargeItems, -ExportLargeItems)`n`r//////////  The script completed successfully, but no parameters were provided.`n`r//////////  For a detailed account of the script, please review the logfile:`n`r//////////  {0}" -f $LogFile)
        Write-Host "//////////`n`r////////////////////////////////////////////////////////////////////////////////////////////////////////n`r"
        [System.Console]::ResetColor()
    }
}
Catch
{
    [System.Console]::ForegroundColor = "Red"
    Write-Host "`n`r\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
    Write-Host ("\\\\\\\\\\`n`r\\\\\\\\\\  Unable to execute the Script!`n`r\\\\\\\\\\")
    Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\`n`r"
    Write-Host ("ERROR: {0}" -f $_.Exception.Message)
    [System.Console]::ResetColor()

    Write-Log "====>  ERROR  <==== | Unable to execute the Script!" -ForegroundColor Magenta;
    Write-Log ("====>  {0}" -f $_.Exception.Message);
	$Error.Clear()
}
Finally
{
    [System.Console]::ForegroundColor = "Green"
    Write-Host "`n`r////////////////////////////////////////////////////////////////////////////////////////////////////////"
    Write-Host ("//////////`n`r//////////  Script completed!...`n`r//////////  {0}" -f $LogFile)
    Write-Host "//////////`n`r////////////////////////////////////////////////////////////////////////////////////////////////////////"
    [System.Console]::ResetColor()

	Write-Log "====>  Script completed!"
}