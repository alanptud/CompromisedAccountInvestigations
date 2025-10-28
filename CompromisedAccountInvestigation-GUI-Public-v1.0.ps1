<#
.SYNOPSIS
    This script queries Azure Log Analytics (Sentinel) and Microsoft Graph API to investigate potential user account compromises between specific dates and times.
    Please note, dates and times entered are in UTC time

.DESCRIPTION
    This script is designed to be run by a member of the security operations team. 
    If the variable below "$usekeyvault" is set to $true, the user running the script will be prompted to authenticate.  This authentication process will access a keyvault as that user to retrive a client secret. 
    
    The script then connects to an Azure App with the client secret to then performs the following actions:
    1.  Queries Sentinel for all IP addresses used by the specified user within a given timeframe.
    2.  Checks each IP against the AzureSpeed API to determine if it's a known Azure IP.
    3.  Checks the Geo Location of each IP address to determine the country.
    4.  Queries the log analytics within Sentinel for 'MailItemsAccessed' operations to find message IDs of accessed emails.
    5.  Uses the Microsoft Graph API to retrieve details (Subject, From, To, Sent Time) for each accessed email.
    6.  Queries the log analytics within Sentinel for file sharing, file access, and inbox rule creation events.
    7.  Exports all collected data into separate sheets within a single Excel file for analysis.

.NOTES
    Prerequisites:
    - The PowerShell modules 'MSAL.PS' and 'ImportExcel' are required. The script will attempt to install them if they are not found (will require OneDrive Sync tool to be running).
    - The Azure AD App Registration must be configured with the necessary Application permissions for:

        - Microsoft Graph: 
            
            * Mail.Read
            * Mail.ReadBasic
            * Mail.ReadBasic.All
            * MailboxItem.Read.All
            * Directory.Read.All
            * User.Read.All

        - Azure Log Analytics: 
        
            * Data.Read

    - The service principal of the Azure App will need to be given permission to read the various log analytics workspaces in Azure.

    Script created by Alan Pike - October 2025
#>

#=======================================================================================================================
# Global Variables
#=======================================================================================================================
# Define the necessary Azure AD and Workspace parameters below

# Specify your Azure tenant ID here
$TenantId = "xxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# Define the ClientID of the Azure App you have created (ensure if has API permissions specified above).  Ensure the App has read permissions on the Log analytics workspace where officeactivity AuditLogs are stored.
$ClientId = "xxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" 

# Specify the workspace ID of your log analytics worksapce (this data may reside in the log analytics workspace associated with Sentinel)
# Ensure the log analytics workspace contains OfficeActivity & AuditLog.  You will need to create an Azure app in your tenant and give this service principle reader access
# on this log analytics workspace  
$workspaceId = "xxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# Register a free account on ip2location.io and include the API token below.  This will be used to query the Geolocation information associated with each IP address
$ipinfoteoken = "XXXXXXXXXXXXXXXXXXXX"

# Define your Key Vault information here (if you dont have one, create one and note the subscription ID of where the keyvault was created)
# Subscription ID of where the Keyvault is stored:
$subscriptionId = "xxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

#Keyvault name (name of it in Azure, not the URI)
$keyVaultName = "<Enter keyvault name here>"

# Name of the secret within the Keyvault that will store the Azure App client secret
# Create a secret and give it the value of the client secret of the azure app that you have created
$secretName = "CompromisedAccountsSecret"

# Boolean Value to identify if keyvault should be used (set to true if you wish to use a keyvault, otherwise populate the client secret below)
$usekeyvault = $false
# Specify the client secret below (should only be used for testing purposes)
$ClientSecrettemp = "Enter Client secret of Azure app here"


#=======================================================================================================================
# ADMIN PRIVILEGE CHECK
#=======================================================================================================================

# Check if the script is running with Administrator privileges. If not, re-launch it as Admin.
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "This script requires Administrator privileges. Attempting to re-launch as Administrator..."
    Start-Process pwsh -Verb RunAs -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $MyInvocation.MyCommand.Path)
    exit
}

#=======================================================================================================================
# GUI SETUP
#=======================================================================================================================


# Load required assemblies for Windows Forms
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
catch {
    Write-Error "Failed to load Windows Forms assemblies. Ensure you are running a supported version of PowerShell."
    exit
}

# --------------------------------------------------------------------------
# Font Definition
# --------------------------------------------------------------------------
# Define a font that we will reuse for all controls.
$GlobalFont = New-Object System.Drawing.Font('Calibri', 9)
$GlobalFont2 = New-Object System.Drawing.Font('Calibri', 10)

# --------------------------------------------------------------------------
# Main Form Creation
# --------------------------------------------------------------------------
$Form = New-Object System.Windows.Forms.Form
$Form.Text = 'Compromised Account Investigation Tool'
$Form.Size = New-Object System.Drawing.Size(600, 750)
$Form.StartPosition = 'CenterScreen'

# --------------------------------------------------------------------------
# GroupBox 1: Investigation Parameters
# --------------------------------------------------------------------------
$GroupBoxInput = New-Object System.Windows.Forms.GroupBox
$GroupBoxInput.Text = 'Investigation Parameters'
$GroupBoxInput.Size = New-Object System.Drawing.Size(660, 310)
$GroupBoxInput.Location = New-Object System.Drawing.Point(15, 15)
$GroupBoxInput.Anchor = 'Top, Left, Right'
$GroupBoxInput.Font = $GlobalFont

# --- Controls for GroupBox 1 (New Vertical Layout with Spacing) ---


# Get the current local time
$localTime = Get-Date

# Convert local time to UTC
$utcTime = $localTime.ToUniversalTime()

# Display both times
Write-Host "Local Time: $localTime"
Write-Host "UTC Time:   $utcTime"

# Add Text to advise of UTC Time
$LabelUTCTime = New-Object System.Windows.Forms.Label
$LabelUTCTime.Text = "(Note: UTC Time is currently: '$($utcTime)'.  Please use UTC times above)"
$LabelUTCTime.Location = New-Object System.Drawing.Point(15, 280)
$LabelUTCTime.AutoSize = $true
$LabelUTCTime.Font = $GlobalFont


# Start Date/Time
$LabelStartDate = New-Object System.Windows.Forms.Label
$LabelStartDate.Text = 'Start Date/Time:'
$LabelStartDate.Location = New-Object System.Drawing.Point(15, 30)
$LabelStartDate.AutoSize = $true
$LabelStartDate.Font = $GlobalFont

$DateTimePickerStart = New-Object System.Windows.Forms.DateTimePicker
$DateTimePickerStart.Format = 'Custom'
$DateTimePickerStart.CustomFormat = 'yyyy-MM-dd HH:mm:ss'
$DateTimePickerStart.ShowUpDown = $true
$DateTimePickerStart.Location = New-Object System.Drawing.Point(15, 55)
$DateTimePickerStart.Width = 250
$DateTimePickerStart.Value = (Get-Date).AddDays(-1) # Default to yesterday
$DateTimePickerStart.Font = $GlobalFont2

# End Date/Time
$LabelEndDate = New-Object System.Windows.Forms.Label
$LabelEndDate.Text = 'End Date/Time:'
$LabelEndDate.Location = New-Object System.Drawing.Point(15, 90)
$LabelEndDate.AutoSize = $true
$LabelEndDate.Font = $GlobalFont

$DateTimePickerEnd = New-Object System.Windows.Forms.DateTimePicker
$DateTimePickerEnd.Format = 'Custom'
$DateTimePickerEnd.CustomFormat = 'yyyy-MM-dd HH:mm:ss'
$DateTimePickerEnd.ShowUpDown = $true
$DateTimePickerEnd.Location = New-Object System.Drawing.Point(15, 115)
$DateTimePickerEnd.Width = 250
$DateTimePickerEnd.Value = Get-Date # Default to now
$DateTimePickerEnd.Font = $GlobalFont2

# User Name
$LabelUsername = New-Object System.Windows.Forms.Label
$LabelUsername.Text = 'Username:'
$LabelUsername.Location = New-Object System.Drawing.Point(15, 150)
$LabelUsername.AutoSize = $true
$LabelUsername.Font = $GlobalFont

$TextBoxUsername = New-Object System.Windows.Forms.TextBox
$TextBoxUsername.Size = New-Object System.Drawing.Size(250, 20)
$TextBoxUsername.Location = New-Object System.Drawing.Point(15, 175)
$TextBoxUsername.Text = "joe.bloggs@tudublin.ie"
$TextBoxUsername.Font = $GlobalFont2

# Output Folder
$LabelOutputPath = New-Object System.Windows.Forms.Label
$LabelOutputPath.Text = 'Output Folder:'
$LabelOutputPath.Location = New-Object System.Drawing.Point(15, 210)
$LabelOutputPath.AutoSize = $true
$LabelOutputPath.Font = $GlobalFont

$TextBoxOutputPath = New-Object System.Windows.Forms.TextBox
$TextBoxOutputPath.Size = New-Object System.Drawing.Size(250, 20)
$TextBoxOutputPath.Location = New-Object System.Drawing.Point(15, 230)
$TextBoxOutputPath.ReadOnly = $true
$TextBoxOutputPath.Text = "C:\CompromisedAccountInvestigations"
$TextBoxOutputPath.Font = $GlobalFont2

$ButtonBrowse = New-Object System.Windows.Forms.Button
$ButtonBrowse.Text = 'Browse...'
$ButtonBrowse.Size = New-Object System.Drawing.Size(75, 25)
$ButtonBrowse.Location = New-Object System.Drawing.Point(275, 230)
$ButtonBrowse.Font = $GlobalFont

# --------------------------------------------------------------------------
# START: Added Checkboxes
# --------------------------------------------------------------------------

$CheckBoxEmailAccess = New-Object System.Windows.Forms.CheckBox
$CheckBoxEmailAccess.Text = 'Check Emails Access'
$CheckBoxEmailAccess.Location = New-Object System.Drawing.Point(360, 90)
$CheckBoxEmailAccess.AutoSize = $true
$CheckBoxEmailAccess.Font = $GlobalFont
$CheckBoxEmailAccess.Checked = $true # Default to checked

$CheckBoxFileAccess = New-Object System.Windows.Forms.CheckBox
$CheckBoxFileAccess.Text = 'Check Files Accessed'
$CheckBoxFileAccess.Location = New-Object System.Drawing.Point(360, 120)
$CheckBoxFileAccess.AutoSize = $true
$CheckBoxFileAccess.Font = $GlobalFont
$CheckBoxFileAccess.Checked = $true # Default to checked

$CheckBoxFileSharing = New-Object System.Windows.Forms.CheckBox
$CheckBoxFileSharing.Text = 'Check Files Shared'
$CheckBoxFileSharing.Location = New-Object System.Drawing.Point(360, 150)
$CheckBoxFileSharing.AutoSize = $true
$CheckBoxFileSharing.Font = $GlobalFont
$CheckBoxFileSharing.Checked = $true # Default to checked

$CheckBoxMailboxRules = New-Object System.Windows.Forms.CheckBox
$CheckBoxMailboxRules.Text = 'Mailbox Rules Added'
$CheckBoxMailboxRules.Location = New-Object System.Drawing.Point(360, 180)
$CheckBoxMailboxRules.AutoSize = $true
$CheckBoxMailboxRules.Font = $GlobalFont
$CheckBoxMailboxRules.Checked = $true # Default to checked

$CheckBoxMailSent = New-Object System.Windows.Forms.CheckBox
$CheckBoxMailSent.Text = 'List Emails Sent'
$CheckBoxMailSent.Location = New-Object System.Drawing.Point(360, 210)
$CheckBoxMailSent.AutoSize = $true
$CheckBoxMailSent.Font = $GlobalFont
$CheckBoxMailSent.Checked = $true # Default to checked

$Checksecinfo = New-Object System.Windows.Forms.CheckBox
$Checksecinfo.Text = 'Security Info Registered'
$Checksecinfo.Location = New-Object System.Drawing.Point(360, 240)
$Checksecinfo.AutoSize = $true
$Checksecinfo.Font = $GlobalFont
$Checksecinfo.Checked = $true # Default to checked

# --------------------------------------------------------------------------
# END: Added Checkboxes
# --------------------------------------------------------------------------

# Run Button
$ButtonRun = New-Object System.Windows.Forms.Button
$ButtonRun.Text = 'Run Investigation'
$ButtonRun.Size = New-Object System.Drawing.Size(180, 50)
$ButtonRun.Location = New-Object System.Drawing.Point(360, 30)
$ButtonRun.BackColor = [System.Drawing.Color]::LimeGreen
##$ButtonRun.Anchor = 'Top, Right'
$ButtonRun.Font = $GlobalFont

# --- Add all controls to GroupBox 1 ---
$GroupBoxInput.Controls.AddRange(@(
    $LabelUTCTime, $LabelStartDate, $DateTimePickerStart,
    $LabelEndDate, $DateTimePickerEnd,
    $LabelUsername, $TextBoxUsername,
    $LabelOutputPath, $TextBoxOutputPath,
    $ButtonBrowse, $ButtonRun,
    # --- Add the new checkboxes here ---
    $CheckBoxEmailAccess, $CheckBoxFileAccess,
    $CheckBoxFileSharing, $CheckBoxMailboxRules, 
    $CheckBoxMailSent, $Checksecinfo
))

# --------------------------------------------------------------------------
# GroupBox 2: Script Output
# --------------------------------------------------------------------------
$GroupBoxOutput = New-Object System.Windows.Forms.GroupBox
$GroupBoxOutput.Text = 'Script Output'
$GroupBoxOutput.Location = New-Object System.Drawing.Point(15, 320)
$GroupBoxOutput.Size = New-Object System.Drawing.Size(560, 400)
$GroupBoxOutput.Anchor = 'Top, Bottom, Left, Right'
$GroupBoxOutput.Font = $GlobalFont

# --- Controls for GroupBox 2 ---
$RichTextBoxOutput = New-Object System.Windows.Forms.RichTextBox
$RichTextBoxOutput.Dock = 'Fill'
$RichTextBoxOutput.ReadOnly = $true
$RichTextBoxOutput.Font = New-Object System.Drawing.Font('Consolas', 9)
$RichTextBoxOutput.BackColor = [System.Drawing.Color]::Black
$RichTextBoxOutput.ForeColor = [System.Drawing.Color]::White

# --- Add control to GroupBox 2 ---
$GroupBoxOutput.Controls.Add($RichTextBoxOutput)

# --------------------------------------------------------------------------
# Final Form Setup
# --------------------------------------------------------------------------
$Form.Controls.Add($GroupBoxInput)
$Form.Controls.Add($GroupBoxOutput)

# --------------------------------------------------------------------------
# Functions and Event Handlers (Add your logic here)
# --------------------------------------------------------------------------
function Write-OutputToGUI {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::White
    )
    $RichTextBoxOutput.SelectionColor = $Color
    $RichTextBoxOutput.AppendText("$Message`n")
    $RichTextBoxOutput.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}



#=======================================================================================================================
# EVENT HANDLERS
#=======================================================================================================================


# Event handler for the "Browse..." button
$ButtonBrowse.Add_Click({
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.SelectedPath = $TextBoxOutputPath.Text
    $result = $FolderBrowser.ShowDialog()
    if ($result -eq 'OK') {
        $TextBoxOutputPath.Text = $FolderBrowser.SelectedPath
    }
})

# Event handler for the "Run Investigation" button
$ButtonRun.Add_Click({
    # Disable the button and show a wait cursor
    $ButtonRun.Enabled = $false
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $RichTextBoxOutput.Clear()

    Add-Type -AssemblyName System.Windows.Forms

    if($usekeyvault -eq $true)
    {
    # Add message that user will now be prompted to enter in credentials for account to connect to Azure key vault to retrieve client secret
    [System.Windows.Forms.MessageBox]::Show(
    "You will now be asked to log into an account that has access to the Key Vault",
    "Azure Key Vault Login",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
    )

    ##############

    # This command will open a browser window or prompt for interactive login
    Write-Host "Please log in to your Azure account to access the Key Vault..."
    ##Connect-AzAccount -Credential (Get-Credential)
    Connect-AzAccount -Tenant $tenantId -Subscription $subscriptionId

    # Confirm subscription context
    Set-AzContext -SubscriptionId $subscriptionId -TenantId $tenantId

    # Retrieve the secret
    $secret = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName

    # Convert the secure string to plain text
    $plainTextSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret.SecretValue)
    )
    }
    elseif ($usekeyvault -eq $false)
    {
        $plainTextSecret = $ClientSecrettemp
    }



    #############

    # Get input from GUI
    $DateStart = $DateTimePickerStart.Value.ToString("yyyy-MM-dd")
    $timestart = $DateTimePickerStart.Value.ToString("HH:mm:ss")
    $DateEnd = $DateTimePickerEnd.Value.ToString("yyyy-MM-dd")
    $TimeEnd = $DateTimePickerEnd.Value.ToString("HH:mm:ss")
    $username = $TextBoxUsername.Text
    $outputPath = $TextBoxOutputPath.Text

    # Clear values for checkboxes
    $checkemailaccesstickbox = $false
    $checkfileaccesstickbox = $false
    $checkfilesharingtickbox = $false
    $checkmailboxrulestickbox = $false
    $checkmailitemssenttickbox = $false
    $checksecinforegticbox = $false

    #Check status of checkboxes

     if ($CheckBoxEmailAccess.Checked) {
     $checkemailaccesstickbox = $true
     }

     if ($CheckBoxFileAccess.Checked) {
     $checkfileaccesstickbox = $true
     }

     if ($CheckBoxFileSharing.Checked) {
     $checkfilesharingtickbox = $true
     }

     if ($CheckBoxMailboxRules.Checked) {
     $checkmailboxrulestickbox = $true
     }

     if ($CheckBoxMailSent.Checked) {
     $checkmailitemssenttickbox = $true
     }

     if($Checksecinfo.Checked) {
     $checksecinforegticbox = $true
     }


    if (-not $username) {
        Write-OutputToGUI -Message "Error: Please enter a username." -Color ([System.Drawing.Color]::Red)
        $ButtonRun.Enabled = $true
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }

    Write-OutputToGUI -Message "Starting investigation for '$username'..." -Color ([System.Drawing.Color]::LimeGreen)
    Write-OutputToGUI -Message "Timeframe: $DateStart $timestart to $DateEnd $TimeEnd"
    Write-OutputToGUI -Message "Output Path: $outputPath"

    # Define API endpoints
    $logAnalyticsUrl = "https://api.loganalytics.io/v1/workspaces/$workspaceId/query"
    $graphApiUrl = "https://graph.microsoft.com/v1.0"

    # Combine date and time into the required ISO 8601 format
    $startTime = "${DateStart}T${timestart}Z"
    $endTime = "${DateEnd}T${TimeEnd}Z"
    $userId = $username
    
    # Define Excel File location
    if (-not (Test-Path -Path $outputPath)) {
        Write-OutputToGUI -Message "Creating output directory: $outputPath" -Color ([System.Drawing.Color]::Yellow)
        New-Item -Path $outputPath -ItemType Directory
    }
    $excelPath = Join-Path -Path $outputPath -ChildPath "$($userName)-ExportData.xlsx"

    # Clean up previous export file if it exists
    if (Test-Path $excelPath) {
        Remove-Item $excelPath -Force
        Write-OutputToGUI -Message "Previous export file deleted." -Color ([System.Drawing.Color]::Yellow)
    }

    #=======================================================================================================================
    # PREREQUISITE CHECKS
    #=======================================================================================================================

    # Check for and install the MSAL.PS module for authentication
    if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
        Write-OutputToGUI -Message "MSAL.PS module not found. Installing..." -Color ([System.Drawing.Color]::Yellow)
        try {
            Install-Module -Name MSAL.PS -Force -AcceptLicense -Scope CurrentUser -ErrorAction Stop
            Write-OutputToGUI -Message "MSAL.PS module installed successfully." -Color ([System.Drawing.Color]::Green)
        }
        catch {
            Write-OutputToGUI -Message "Failed to install MSAL.PS module. Please install it manually and re-run the script." -Color ([System.Drawing.Color]::Red)
            $ButtonRun.Enabled = $true
            $Form.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }
    }
    Import-Module MSAL.PS

    # Check for and install the ImportExcel module for reporting
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-OutputToGUI -Message "ImportExcel module not found. Installing..." -Color ([System.Drawing.Color]::Yellow)
        try {
            Install-Module -Name ImportExcel -Force -AcceptLicense -Scope CurrentUser -ErrorAction Stop
            Write-OutputToGUI -Message "ImportExcel module installed successfully." -Color ([System.Drawing.Color]::Green)
        }
        catch {
            Write-OutputToGUI -Message "Failed to install ImportExcel module. Please install it manually and re-run the script." -Color ([System.Drawing.Color]::Red)
            $ButtonRun.Enabled = $true
            $Form.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }
    }
    Import-Module ImportExcel

    #=======================================================================================================================
    # AUTHENTICATION
    #=======================================================================================================================

$ClientSecret = $plainTextSecret | ConvertTo-SecureString -AsPlainText -Force

$logAnalyticsScope = "https://api.loganalytics.io/.default"
$graphScope = "https://graph.microsoft.com/.default"


try {
    # Acquire a token for Log Analytics using client credentials
    Write-OutputToGUI -Message "Acquiring token for Log Analytics..." -Color ([System.Drawing.Color]::White)
    $logAnalyticsTokenResponse = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret -Scopes $logAnalyticsScope -ErrorAction Stop
    $logAnalyticsToken = $logAnalyticsTokenResponse.AccessToken
    Write-OutputToGUI -Message "Successfully acquired application token for Log Analytics." -Color ([System.Drawing.Color]::Green)

    # Acquire a token for Microsoft Graph using client credentials
    Write-OutputToGUI -Message "Acquiring token for Microsoft Graph..." -Color ([System.Drawing.Color]::White)
    $graphTokenResponse = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret -Scopes $graphScope -ErrorAction Stop
    $graphToken = $graphTokenResponse.AccessToken
    Write-OutputToGUI -Message "Successfully acquired application token for Microsoft Graph." -Color ([System.Drawing.Color]::Green)
}
catch {
    Write-Error "Authentication failed. Please check the following:
    1. The Client ID, Tenant ID, and Client Secret are correct.
    2. The App Registration has been granted the required *Application* permissions in Azure AD.
    3. Admin consent has been granted for those permissions.
    Error: $_"
    return # Exit the script if authentication fails
}

# Begin the queries

#=======================================================================================================================
# QUERY 1: IP Addresses Used (working)
# Get list of IP addresses used during the timeframe
#=======================================================================================================================


Write-OutputToGUI -Message "`n** Querying for IP addresses used by the user..." -Color ([System.Drawing.Color]::Cyan)

$kqlQueryIPAddressQuery = @"
OfficeActivity
| where UserId contains '$username'
| where TimeGenerated >= datetime($startTime) and TimeGenerated <= datetime($endTime)
| extend ResolvedClientIP = coalesce(ClientIP, Client_IPAddress)
| where isnotempty(ResolvedClientIP)
| distinct ResolvedClientIP
"@

$requestIPAddressUsed = @{ "query" = $kqlQueryIPAddressQuery } | ConvertTo-Json
$headers = @{
    "Authorization" = "Bearer $logAnalyticsToken"
    "Content-Type"  = "application/json"
}

$responseIP = Invoke-RestMethod -Uri $logAnalyticsUrl -Method Post -Headers $headers -Body $requestIPAddressUsed

#Array used to store info on whether IP address is an Azure or not
$ipInfoArray = @()
#Array used to store Geo Location of IP address
$ipGeoLocationArray = @()
$ipfulldetails = @()

if ($responseIP.tables[0].rows.Count -gt 0) {
    $ipAddresses = $responseIP.tables[0].rows.ForEach({ $_[0] })
    Write-OutputToGUI -Message "Found $($ipAddresses.Count) unique IP addresses. Checking their status..." -Color ([System.Drawing.Color]::White)

    foreach ($ipAddress in $ipAddresses) {

        $apiUrl = "https://www.azurespeed.com/api/ipAddress?ipOrDomain=$ipAddress"
        $GeoLocation = "https://api.ip2location.io/?key=$($ipinfoteoken)&ip=$($ipAddress)"

        # Check if IP address is an IPv6 address starting with 2603
        If ($ipAddress -like "2603:*" -or $ipAddress -like "2a01:111*")
            {
               $status = "Microsoft IP range"
               
            }    
        else
        {

            # Query if the IP is an Azure IP address
            try {
            # Query the API to check if it's an Azure IP
            $apiResponse = Invoke-RestMethod -Uri $apiUrl -Method Get -ErrorAction Stop
            $status = if ($apiResponse) { "Azure IP Address" } else { "Public IP" }
            }
            catch {
            $status = "Assume Public IP" # Assume public if API call fails or returns non-true
            }

        }
            # Query the Geo Location of the Azure IP address
            try {
            
            $apigeoResponse = Invoke-RestMethod -Uri $GeoLocation -Method Get -ErrorAction Stop
            $city = $apigeoResponse.city_name
            $as = $apigeoResponse.as
            $country = $apigeoResponse.country_name
            write-host ""

            }
            catch {
            $country = "Unknown"
            $city = "Unknown"
            }
        

        $ipInfoArray += [PSCustomObject]@{
            IPAddress = $ipAddress
            Status    = $status
        }
        Write-Host "  - $ipAddress : $status"


         $ipGeoLocationArray += [PSCustomObject]@{
            IPAddress = $ipAddress
            Country = $country
            City = $city
            AS = $as
        }

         $ipfulldetails += [PSCustomObject]@{
            IPAddress = $ipAddress
            Status    = $status
            Country = $country
            City    = $city
            AS      = $as
        }

        Write-OutputToGUI -Message "  - $ipAddress : $status : (City: $City | County: $country | AS: $as )" -Color ([System.Drawing.Color]::Gray)

    }
}
else {
    Write-OutputToGUI -Message "No IP addresses found for the user in the specified time range."  -Color ([System.Drawing.Color]::Red)
    $NoIPAddressesFound = "No IP addresses Found"
    $NoIPAddressesFound | Export-Excel -Path $excelPath -WorksheetName "IPAddresses" -AutoSize -TableName "IPAddresses"

}

     $ipfulldetails | Export-Excel -Path $excelPath -WorksheetName "IPAddresses" -AutoSize -TableName "IPAddresses"
    Write-OutputToGUI -Message "IP address data exported to Excel." -Color ([System.Drawing.Color]::Green)


#=======================================================================================================================
# QUERY 2: ACCESSED EMAILS (working)
# Get message details for all emails accessed by the user.
#=======================================================================================================================

If($checkemailaccesstickbox -eq $true)
{
Write-OutputToGUI -Message "`n** Querying for accessed email messages..." -Color ([System.Drawing.Color]::Cyan)

$kqlQueryMessages = @"
OfficeActivity
| where UserId contains '$username'
| where Operation == 'MailItemsAccessed'
| where TimeGenerated >= datetime($startTime) and TimeGenerated <= datetime($endTime)
| extend ParsedData = parse_json(Folders)
| mv-expand ParsedData
| extend FolderItems = ParsedData.FolderItems
| mv-expand FolderItems
| extend InternetMessageId = tostring(FolderItems.InternetMessageId)
| where isnotempty(InternetMessageId)
| distinct TimeGenerated, InternetMessageId, Client_IPAddress
"@

$requestBodyMessages = @{ "query" = $kqlQueryMessages } | ConvertTo-Json

$responseMessages = Invoke-RestMethod -Uri $logAnalyticsUrl -Method Post -Headers $headers -Body $requestBodyMessages

$emailResults = @()
if ($responseMessages.tables[0].rows.Count -gt 0) {
    Write-OutputToGUI -Message  "Found $($responseMessages.tables[0].rows.Count) mail access events. Retrieving message details from Graph API..." -Color ([System.Drawing.Color]::White)
    foreach ($row in $responseMessages.tables[0].rows) {
        $timegenerated = $row[0]
        $internetMessageId = $row[1]
        $clientIpAddress = $row[2]

        $encodedMessageId = [System.Net.WebUtility]::UrlEncode($internetMessageId)
        $graphMessageUri = "$graphApiUrl/users/$userId/messages?`$filter=internetMessageId eq '$encodedMessageId'"

        $graphHeaders = @{ "Authorization" = "Bearer $graphToken" }

        $graphResponse = Invoke-RestMethod -Uri $graphMessageUri -Method Get -Headers $graphHeaders

         # Initialize variables for CSV data
        $fromAddress = $null
        $toAddresses = $null
        $subject = $null
        $sentDateTime = $null

         #Query the IP Address Status whether it is Azure or not
        $ipStatus = ($ipInfoArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).Status
         if (-not $ipStatus) {
            $ipStatus = "Unknown IP"  # Default if not found
        }

         #Query the Country of the Geo Location of the IP Address
        $ipGeoStatusCountry = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).Country
         if (-not $ipStatus) {
            $ipGeoStatusCountry = "Unknown Country"  # Default if not found
        }

        #Query the City of the Geo Location of the IP Address
        $ipGeoStatusCity = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).City
         if (-not $ipStatus) {
            $ipGeoStatusCity  = "Unknown City"  # Default if not found
        }

        #Query the AS of the Geo Location of the IP Address
        $ipGeoStatusAS = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).AS
         if (-not $ipStatus) {
            $ipGeoStatusAS  = "Unknown City"  # Default if not found
        }

        if ($graphResponse.value.Count -gt 0) {
            $message = $graphResponse.value[0]
            #Write-Host "  - Subject: $($message.subject)"
            Write-OutputToGUI -Message "  - Subject: $($message.subject)" -Color ([System.Drawing.Color]::Gray)
            $toAddresses = $message.toRecipients | ForEach-Object { $_.emailAddress.address }
            $toAddressProper = "$($toAddresses -join ', ')"
            $emailResults += [PSCustomObject]@{
                TimeGenerated     = $timegenerated
                InternetMessageId = $internetMessageId
                ClientIpAddress   = $clientIpAddress
                IPStatus          = $ipStatus
                IPCountry         = $ipGeoStatusCountry
                IPCity            = $ipGeoStatusCity
                ISP               = $ipGeoStatusAS
                Subject           = $message.subject
                SentDateTime      = $message.sentDateTime
                FromAddress       = $message.from.emailAddress.address
                ToAddresses       = $toAddressProper
                HasAttachments    = $message.hasAttachments
            }
        }
        else {
            #Write-Host "  - No message found for InternetMessageId: $internetMessageId"
            Write-OutputToGUI -Message "  - No message found for InternetMessageId: $internetMessageId" -Color ([System.Drawing.Color]::Red)
            $emailResults += [PSCustomObject]@{
                TimeGenerated     = $timegenerated
                InternetMessageId = $internetMessageId
                ClientIpAddress   = $clientIpAddress
                IPStatus          = $ipStatus
                IPCountry         = $ipGeoStatusCountry
                IPCity            = $ipGeoStatusCity
                ISP               = $ipGeoStatusAS
                Subject           = "Message not found in mailbox"
                SentDateTime      = $null
                FromAddress       = $null
                ToAddresses       = $null
                HasAttachments    = $null
            }
        }
    }
    $emailResults | Export-Excel -Path $excelPath -WorksheetName "AccessedEmails" -AutoSize -TableName "AccessedEmails"
    Write-Host "Email access data exported to Excel." -ForegroundColor Green
}
else {
    Write-Host "No 'MailItemsAccessed' events found." -ForegroundColor Yellow
    $noemailsfound = "No Mail Items accessed during this time frame"
    $noemailsfound | Export-Excel -Path $excelPath -WorksheetName "AccessedEmails" -AutoSize -TableName "AccessedEmails"
}

}
#=======================================================================================================================
# QUERY 3: FILE SHARING (working - but needs Client IP location)
# Get details of any files shared from OneDrive or SharePoint.
#=======================================================================================================================
if($checkfilesharingtickbox -eq $true)
{
Write-OutputToGUI -Message "`n** Querying for file sharing events..." -Color ([System.Drawing.Color]::Cyan)

$kqlQuerySharing = @"
OfficeActivity 
| where UserId contains "$username"
| where TimeGenerated >= datetime($startTime) and TimeGenerated <= datetime($endTime)
| where RecordType contains "SharePointSharingOperation"
| distinct TimeGenerated, Operation, OfficeObjectId, SourceFileName, ClientIP, Event_Data, TargetUserOrGroupName, TargetUserOrGroupType
"@

$requestBodySharing = @{ "query" = $kqlQuerySharing } | ConvertTo-Json
$responseSharing = Invoke-RestMethod -Uri $logAnalyticsUrl -Method Post -Headers $headers -Body $requestBodySharing

# Check the response
$filesharingcount = 0
if ($responseSharing) {
    Write-OutputToGUI -Message "Query results for Files Shared..." -Color ([System.Drawing.Color]::White)
    $responseSharing.tables[0].rows |
    ForEach-Object {
        $timegenerated = $_[0]
        $operation = $_[1]
        $OfficeObjectId = $_[2]
        $SourceFileName = $_[3]
        $ClientIP = $_[4]
        $Event_Data = $_[5]
        $targetuserorgroupname = $_[6]
        $targetuserorgrouptype = $_[7]
        $filesharingcount++

 #Query the IP Address Status whether it is Azure or not
        $ipStatus = ($ipInfoArray | Where-Object { $_.IPAddress -eq $ClientIP }).Status
         if (-not $ipStatus) {
            $ipStatus = "Unknown IP"  # Default if not found
            }

         #Query the Country of the Geo Location of the IP Address
        $ipGeoStatusCountry = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $ClientIP }).Country
         if (-not $ipStatus) {
            $ipGeoStatusCountry = "Unknown Country"  # Default if not found
            }

         #Query the City of the Geo Location of the IP Address
        $ipGeoStatusCity = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $ClientIP }).City
         if (-not $ipStatus) {
            $ipGeoStatusCity  = "Unknown City"  # Default if not found
            }

        #Query the AS of the Geo Location of the IP Address
        $ipGeoStatusAS = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $ClientIP }).AS
         if (-not $ipStatus) {
            $ipGeoStatusAS  = "Unknown City"  # Default if not found
            }

        $sharingResults = [PSCustomObject]@{
                TimeGenerated = $_[0]
                Operation = $_[1]
                OfficeObjectId = $_[2]
                SourceFileName = $_[3]
                ClientIP = $_[4]
                IPStatus = $ipStatus
                IPCountry = $ipGeoStatusCountry
                IPCity = $ipGeoStatusCity
                ISP = $ipGeoStatusAS
                EventData = $_[5]
                TargetUserOrGroupName = $_[6]
                TargetUserOrGroupType = $_[7]
                }
        
        $sharingResults | Export-Excel -Path $excelPath -WorksheetName "FileSharing" -AutoSize -Append -TableName "FileSharing"
        
        }
        Write-OutputToGUI -Message "A total of $filesharingcount File sharing events exported to Excel." -Color ([System.Drawing.Color]::Green)
    } else {
        Write-OutputToGUI -Message "No file sharing events found." -Color ([System.Drawing.Color]::Yellow)
        $nofilesshared = "No Files Shared during this time frame"
        $nofilesshared | Export-Excel -Path $excelPath -WorksheetName "FileSharing" -AutoSize -Append -TableName "FileSharing"
    }
}
#=======================================================================================================================
# QUERY 4: FILE ACCESS (working)
# Get details of files accessed or downloaded.
#=======================================================================================================================
if($checkfileaccesstickbox -eq $true)
{

Write-OutputToGUI -Message "`n** Querying for file access events..." -Color ([System.Drawing.Color]::Cyan)
$kqlQueryFileAccess = @"
OfficeActivity 
| where UserId contains "$username"
| where TimeGenerated >= datetime($startTime) and TimeGenerated <= datetime($endTime)
| where Operation startswith "File"
| distinct TimeGenerated, Operation, OfficeObjectId, SourceFileName, ClientIP, Event_Data
"@

$requestBodyFileAccess = @{ "query" = $kqlQueryFileAccess } | ConvertTo-Json
$responseFileAccess = Invoke-RestMethod -Uri $logAnalyticsUrl -Method Post -Headers $headers -Body $requestBodyFileAccess

# Check the response
$fileaccesscount = 0
if ($responseFileAccess) {
    Write-OutputToGUI -Message "Query results for Files Accessed..." -Color ([System.Drawing.Color]::White)
    $responseFileAccess.tables[0].rows |
    ForEach-Object {
        $timegenerated = $_[0]
        $operation = $_[1]
        $OfficeObjectId = $_[2]
        $SourceFileName = $_[3]
        $ClientIP = $_[4]
        $Event_Data = $_[5]
        $fileaccesscount++

        # Find the status from the $ipInfoArray for the current Client_IPAddress
        $ipStatus = ($ipInfoArray | Where-Object { $_.IPAddress -eq $ClientIP }).Status
        if (-not $ipStatus) {
            $ipStatus = "Unknown IP"  # Default if not found
        }

          #Query the Country of the Geo Location of the IP Address
        $ipGeoStatusCountry = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $ClientIP }).Country
         if (-not $ipGeoStatusCountry) {
            $ipGeoStatusCountry = "Unknown Country"  # Default if not found
        }

         #Query the City of the Geo Location of the IP Address
        $ipGeoStatusCity = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $ClientIP }).City
         if (-not $ipGeoStatusCity) {
            $ipGeoStatusCity  = "Unknown City"  # Default if not found
        }
       
        
        #Query the AS of the Geo Location of the IP Address
        $ipGeoStatusAS = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $ClientIP }).AS
         if (-not $ipGeoStatusAS) {
            $ipGeoStatusAS  = "Unknown AS"  # Default if not found
        }

# Create a custom object for each row of data, including Length as a dummy property
$csvObject = [pscustomobject]@{
            TimeGenerated         = $timegenerated
            Operation             = $operation
            OfficeObjectId        = $OfficeObjectId
            SourceFileName        = $SourceFileName
            ClientIP              = $ClientIP
            IPStatus              = $ipStatus
            IPCountry             = $ipGeoStatusCountry
            IPCity                = $ipGeoStatusCity
            ISP                   = $ipGeoStatusAS
            Event_Data            = $Event_Data
        }


$csvObject | Export-Excel -Path $excelPath -WorksheetName "FileAccess" -AutoSize -Append -TableName "FileAccess"

 }
     Write-OutputToGUI -Message "A total of $fileaccesscount File Access events exported to Excel." -Color ([System.Drawing.Color]::Green)
} else {
    Write-OutputToGUI -Message "No results found."
    $nofilesaccessed = "No Files Accessed during this time frame"
    $nofilesaccessed | Export-Excel -Path $excelPath -WorksheetName "FileAccess" -AutoSize -Append -TableName "FileAccess"
}
}
#=======================================================================================================================
# QUERY 5: MAILBOX RULES (working)
# Get details of any mailbox rules that were created or modified.
#=======================================================================================================================

if($checkmailboxrulestickbox -eq $true)
{
    Write-OutputToGUI -Message "`n** Querying for mailbox rule events..." -Color ([System.Drawing.Color]::Cyan)

    $kqlQueryMailboxRules = 
    @"
    OfficeActivity
    | where UserId contains '$username'
    | where TimeGenerated >= datetime($startTime) and TimeGenerated <= datetime($endTime)
    | where Operation has "InboxRule"
    | extend ParsedData = parse_json(Parameters)
    | mv-expand ParsedData
    | project TimeGenerated, Operation, RecordType, OfficeObjectId, UserId, ParsedData, ClientIP 
"@

    $requestBodyMailboxRules = @{ "query" = $kqlQueryMailboxRules } | ConvertTo-Json
    $responseMailboxRules = Invoke-RestMethod -Uri $logAnalyticsUrl -Method Post -Headers $headers -Body $requestBodyMailboxRules

    if ($responseMailboxRules.tables[0].rows.Count -gt 0) 
    {
        Write-OutputToGUI -Message "Found $($responseMailboxRules.tables[0].rows.Count) mailbox rule events. Exporting..." -Color ([System.Drawing.Color]::White)

        $responseMailboxRules.tables[0].rows |
        ForEach-Object {
            $timegenerated = $_[0]
            $operation = $_[1]
            $RecordType = $_[2]
            $OfficeObjectId = $_[3]
            $UserId = $_[4]
            $ParsedData = $_[5]
            $ClientIP = $_[6]

        # Find the status from the $ipInfoArray for the current Client_IPAddress
        $ipStatus = ($ipInfoArray | Where-Object { $_.IPAddress -eq $ClientIP }).Status
        if (-not $ipStatus) {
            $ipStatus = "Unknown IP"  # Default if not found
        }

          #Query the Country of the Geo Location of the IP Address
        $ipGeoStatusCountry = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $ClientIP }).Country
         if (-not $ipStatus) {
            $ipGeoStatusCountry = "Unknown Country"  # Default if not found
        }

         #Query the City of the Geo Location of the IP Address
        $ipGeoStatusCity = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $ClientIP }).City
         if (-not $ipStatus) {
            $ipGeoStatusCity  = "Unknown City"  # Default if not found
        }
        
        #Query the AS of the Geo Location of the IP Address
        $ipGeoStatusAS = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $ClientIP }).AS
         if (-not $ipStatus) {
            $ipGeoStatusAS  = "Unknown City"  # Default if not found
        }

        # Create a custom object for each row of data, including Length as a dummy property
        $csvObject = [pscustomobject]@{
            TimeGenerated         = $timegenerated
            Operation             = $operation
            UserId                = $UserId
            RecordType            = $RecordType
            OfficeObjectId        = $OfficeObjectId
            ClientIP              = $ClientIP
            IPStatus              = $ipStatus
            IPCountry             = $ipGeoStatusCountry
            IPCity                = $ipGeoStatusCity 
            ISP                   = $ipGeoStatusAS
            ParsedData            = $ParsedData
            }

        $csvObject | Export-Excel -Path $excelPath -WorksheetName "MailBoxRules" -AutoSize -Append -TableName "MailboxRules"

     
        }
        Write-OutputToGUI -Message "Mailbox rule data exported to Excel." -Color ([System.Drawing.Color]::Green)
        }
else {
    Write-OutputToGUI -Message "No mailbox rule events found during this timeframe." -ForegroundColor Yellow
    $nomailboxruledata = "No Mailbox Rules Added"
    $nomailboxruledata | Export-Excel -Path $excelPath -WorksheetName "MailBoxRules" -AutoSize -Append -TableName "MailboxRules"
}
}

#=======================================================================================================================
# QUERY 6: Sent Emails (working)
# Query for any emails sent by user during this time.
#=======================================================================================================================
If($checkmailitemssenttickbox -eq $true)
{
Write-OutputToGUI -Message "`n** Querying for emails sent..." -Color ([System.Drawing.Color]::Cyan)

$kqlQueryMessages = @"
OfficeActivity
| where UserId contains '$username'
| where Operation == 'Send'
| where TimeGenerated >= datetime($startTime) and TimeGenerated <= datetime($endTime)
| extend itemData = parse_json(Item)
| project TimeGenerated, InternetMessageId = itemData.InternetMessageId, ClientIP
"@

$requestBodyMessages = @{ "query" = $kqlQueryMessages } | ConvertTo-Json

$responseMessages = Invoke-RestMethod -Uri $logAnalyticsUrl -Method Post -Headers $headers -Body $requestBodyMessages

$emailResults = @()
$emailsentcount = 0
if ($responseMessages.tables[0].rows.Count -gt 0) {
    Write-OutputToGUI -Message  "Found $($responseMessages.tables[0].rows.Count) emails sent. Retrieving message details from Graph API..." -Color ([System.Drawing.Color]::White)
    foreach ($row in $responseMessages.tables[0].rows) {
        $timegenerated = $row[0]
        $internetMessageId = $row[1]
        $clientIpAddress = $row[2]
        $emailsentcount++

        $encodedMessageId = [System.Net.WebUtility]::UrlEncode($internetMessageId)
        $graphMessageUri = "$graphApiUrl/users/$userId/messages?`$filter=internetMessageId eq '$encodedMessageId'"

        $graphHeaders = @{ "Authorization" = "Bearer $graphToken" }

        $graphResponse = Invoke-RestMethod -Uri $graphMessageUri -Method Get -Headers $graphHeaders

         # Initialize variables for CSV data
        $fromAddress = $null
        $toAddresses = $null
        $subject = $null
        $sentDateTime = $null

         #Query the IP Address Status whether it is Azure or not
        $ipStatus = ($ipInfoArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).Status
         if (-not $ipStatus) {
            $ipStatus = "Unknown IP"  # Default if not found
        }

         #Query the Country of the Geo Location of the IP Address
        $ipGeoStatusCountry = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).Country
         if (-not $ipStatus) {
            $ipGeoStatusCountry = "Unknown Country"  # Default if not found
        }

        #Query the City of the Geo Location of the IP Address
        $ipGeoStatusCity = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).City
         if (-not $ipStatus) {
            $ipGeoStatusCity  = "Unknown City"  # Default if not found
        }
        
        #Query the AS of the Geo Location of the IP Address
        $ipGeoStatusAS = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).AS
         if (-not $ipStatus) {
            $ipGeoStatusAS  = "Unknown City"  # Default if not found
        }

        if ($graphResponse.value.Count -gt 0) {
            $message = $graphResponse.value[0]
            #Write-Host "  - Subject: $($message.subject)"
            Write-OutputToGUI -Message "  - Subject: $($message.subject)" -Color ([System.Drawing.Color]::Gray)
            $toAddresses = $message.toRecipients | ForEach-Object { $_.emailAddress.address }
            $toAddressProper = "$($toAddresses -join ', ')"
            $emailResults += [PSCustomObject]@{
                TimeGenerated     = $timegenerated
                InternetMessageId = $internetMessageId
                ClientIpAddress   = $clientIpAddress
                IPStatus          = $ipStatus
                IPCountry         = $ipGeoStatusCountry
                IPCity            = $ipGeoStatusCity
                ISP               = $ipGeoStatusAS
                Subject           = $message.subject
                SentDateTime      = $message.sentDateTime
                FromAddress       = $message.from.emailAddress.address
                ToAddresses       = $toAddressProper
            }
        }
        else {
            #Write-Host "  - No message found for InternetMessageId: $internetMessageId"
            Write-OutputToGUI -Message "  - No message found for InternetMessageId: $internetMessageId" -Color ([System.Drawing.Color]::Red)
            $emailResults += [PSCustomObject]@{
                TimeGenerated     = $timegenerated
                InternetMessageId = $internetMessageId
                ClientIpAddress   = $clientIpAddress
                IPStatus          = $ipStatus
                IPCountry         = $ipGeoStatusCountry
                IPCity            = $ipGeoStatusCity
                ISP               = $ipGeoStatusAS
                Subject           = "Message not found in mailbox"
                SentDateTime      = $null
                FromAddress       = $null
                ToAddresses       = $null
            }
        }
    }
    $emailResults | Export-Excel -Path $excelPath -WorksheetName "EmailsSent" -AutoSize -TableName "EmailsSent"
     Write-OutputToGUI -Message "A total of $emailsentcount emails sent during this period have been exported to Excel." -Color ([System.Drawing.Color]::Green)
}
else {
    Write-OutputToGUI -Message "No 'Emails Sent' events found." -ForegroundColor Yellow
    $noemailssent = "No emails sent during this time frame"
    $noemailssent | Export-Excel -Path $excelPath -WorksheetName "EmailsSent" -AutoSize -TableName "EmailsSent"
}

}

#=======================================================================================================================
# QUERY 7: Security Info Registered
# Query for any additional security info registered in this timeframe
#=======================================================================================================================

if ($checksecinforegticbox -eq $true)
{

Write-OutputToGUI -Message "`n** Querying for Security info registered..." -Color ([System.Drawing.Color]::Cyan)

$kqlQuerySecInfo = @"
AuditLogs
| where OperationName == "User registered security info"
| where Result == "success"
| extend UserPrincipalName = tostring(InitiatedBy.user.userPrincipalName)
| where UserPrincipalName contains '$username'
| where TimeGenerated >= datetime($startTime) and TimeGenerated <= datetime($endTime)
| extend IPAddress = tostring(InitiatedBy.user.ipAddress)
| project TimeGenerated, UserPrincipalName, IPAddress, ResultReason
| sort by TimeGenerated desc
"@

$requestsecinforeg = @{ "query" = $kqlQuerySecInfo } | ConvertTo-Json
$responsesecinforeg = Invoke-RestMethod -Uri $logAnalyticsUrl -Method Post -Headers $headers -Body $requestsecinforeg

$secinforegisteredcount = 0
#####
# Check the response
if ($responsesecinforeg.tables[0].rows.Count -gt 0) {
    Write-OutputToGUI -Message "Query results for security info registered on account..." -Color ([System.Drawing.Color]::White)
    $responsesecinforeg.tables[0].rows |
    ForEach-Object {
        $timegenerated = $_[0]
        $User = $_[1]
        $ClientIP = $_[2]
        $AuthMethod = $_[3]
        $secinforegisteredcount++

        # Find the status from the $ipInfoArray for the current Client_IPAddress
        $ipStatus = ($ipInfoArray | Where-Object { $_.IPAddress -eq $ClientIP }).Status
        if (-not $ipStatus) {
            $ipStatus = "Unknown IP"  # Default if not found
        }

          #Query the Country of the Geo Location of the IP Address
        $ipGeoStatusCountry = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).Country
         if (-not $ipStatus) {
            $ipGeoStatusCountry = "Unknown Country"  # Default if not found
        }

        #Query the City of the Geo Location of the IP Address
        $ipGeoStatusCity = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).City
         if (-not $ipStatus) {
            $ipGeoStatusCity  = "Unknown City"  # Default if not found
        }
        
        #Query the AS of the Geo Location of the IP Address
        $ipGeoStatusAS = ($ipGeoLocationArray | Where-Object { $_.IPAddress -eq $clientIpAddress }).AS
         if (-not $ipStatus) {
            $ipGeoStatusAS  = "Unknown City"  # Default if not found
        }


# Create a custom object for each row of data, including Length as a dummy property
$csvObject = [pscustomobject]@{
            TimeGenerated         = $timegenerated
            User                  = $User
            ClientIP              = $ClientIP
            IPStatus              = $ipStatus
            IPCountry             = $ipGeoStatusCountry
            IPCity                = $ipGeoStatusCity
            ISP                   = $ipGeoStatusAS
            AuthMethodAdded       = $AuthMethod
        }

$csvObject | Export-Excel -Path $excelPath -WorksheetName "SecInfoAdded" -AutoSize -Append -TableName "SecInfoAdded"

 }
    Write-OutputToGUI -Message "A total of $secinforegisteredcount events related to the registration of security info have been exported to Excel ." -Color ([System.Drawing.Color]::Green)
} else {
   Write-OutputToGUI -Message "No security information registered during this timeframe."
   $nosecurityinforegistered = "No Security Info registered during this time frame"
   $nosecurityinforegistered | Export-Excel -Path $excelPath -WorksheetName "SecInfoAdded" -AutoSize -Append -TableName "SecInfoAdded"
}
}

#=======================================================================================================================
# FINALIZATION
#=======================================================================================================================

 if (Test-Path $excelPath) {
        Write-OutputToGUI -Message "`nScript finished. All data has been exported to:" -Color ([System.Drawing.Color]::LimeGreen)
        Write-OutputToGUI -Message $excelPath -Color ([System.Drawing.Color]::White)
    } else {
        Write-OutputToGUI -Message "`nScript finished, but no data was found to export." -Color ([System.Drawing.Color]::Yellow)
    }

    $ButtonRun.Enabled = $true
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# Display the form
$Form.ShowDialog() | Out-Null