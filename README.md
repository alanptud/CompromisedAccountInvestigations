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
