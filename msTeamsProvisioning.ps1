param
([object]$WebhookData)
$VerbosePreference = 'continue'

#region Verify if Runbook is started from Webhook.

# If runbook was called from Webhook, WebhookData will not be null.
if ($WebHookData){

    # Collect properties of WebhookData
    $WebhookName     =     $WebHookData.WebhookName
    $WebhookHeaders  =     $WebHookData.RequestHeader
    $WebhookBody     =     $WebHookData.RequestBody

    # Collect individual headers. Input converted from JSON.
    $From = $WebhookHeaders.From
    
    $Input = (ConvertFrom-Json -InputObject $WebhookBody)
    Write-Verbose "WebhookBody: $Input"
    Write-Output -InputObject ('Runbook started from webhook' -f $WebhookName, $From)
}
else
{
   Write-Error -Message 'Runbook was not started from Webhook' -ErrorAction stop
}
#endregion

#Automatic Teams creation starts here
#Global variables
$tenantId = "yourTenantID"

$guestAccess = "$($Input.allowGuestAccess)"
$SPSite = $Input.SiteURL
$SPList = $Input.ListName
$SPListItemID = $Input.ListItemID
$teamsName = $Input.TeamsName
$teamsAlias  = $teamsName -replace '[^a_-zA-Z0-9]', ''

Function Update-site{
Param( 

    #add channel folders to SharePoint using PnP PowerShell
    $spoconn = Connect-PnPOnline –Url https://jh365dev.sharepoint.com/sites/$teamsAlias –Credentials (Get-AutomationPSCredential -Name 'YourAutomationAccount') -ReturnConnection
    Add-PnPFolder -Name "General" -Folder "/Shared Documents"
    Add-PnPFolder -Name "01 Planning" -Folder "/Shared Documents"
    Add-PnPFolder -Name "02 Execution" -Folder "/Shared Documents"
    Add-PnPFolder -Name "03 Final" -Folder "/Shared Documents"

    #add to hubsite if needed
    Add-PnPHubSiteAssociation -Site https://jh365dev.sharepoint.com/sites/$teamsAlias -HubSite "yourhubsiteURL"

    #copy files to new Team channel if needed
    Copy-PnPFile -SourceUrl /sites/templates/Shared%20Documents/Templates.docx -TargetUrl /sites/$teamsAlias/Shared%20Documents/General -Force -Confirm
    
} #End Update-site 


#Connecting to O365
Connect-MicrosoftTeams -TenantId $tenantId -Credential (Get-AutomationPSCredential -Name 'YourAutomationAccount')

#Create new Team
$team = New-Team -MailNickName $teamsAlias -DisplayName $Input.TeamsDisplayName -Visibility Private
Add-TeamUser -GroupId $team.GroupId -User $Input.TeamsOwner -Role Owner

#Add channels
New-TeamChannel -GroupId $team.GroupId -DisplayName "01 Planning"
New-TeamChannel -GroupId $team.GroupId -DisplayName "02 Execution"
New-TeamChannel -GroupId $team.GroupId -DisplayName "03 Final"

#Teams created
Write-Output 'Teams created'

#call Update site function
Update-site

#Disabling Guest Access to Teams
Write-Output "GuestAccess allowed: $guestAccess"


if($guestAccess -eq "No")
{
    try{
            #importing AzureADPreview modules
            Import-Module AzureADPreview
            Connect-AzureAD -TenantId $tenantId -Credential (Get-AutomationPSCredential -Name 'YourAutomationAccount')

            #Turn OFF guest access
            $template = Get-AzureADDirectorySettingTemplate | ? {$_.displayname -eq "group.unified.guest"}
            $settingsCopy = $template.CreateDirectorySetting()
            $settingsCopy["AllowToAddGuests"]=$False

            New-AzureADObjectSetting -TargetType Groups -TargetObjectId $team.GroupId -DirectorySetting $settingsCopy

            #Verify settings
            Get-AzureADObjectSetting -TargetObjectId $team.GroupId -TargetType Groups | fl Values
            
            #reset $guestaccess flag
            $guestAccess = "NA"
    }
    catch{
            #Catch errors
            Write-Output "An error occurred:"
            Write-Output $_.Exception.Message

            $spoconn = Connect-PnPOnline –Url $SPSite –Credentials (Get-AutomationPSCredential -Name 'YourAutomationAccount') -ReturnConnection -Verbose
            $itemupdate = Set-PnPListItem -List $SPList -Identity $SPListItemID -Values @{"TeamsCreated" = "Error Occured setting GuestAccess"} -Connection $spoconn
        }

}

#Updating SharePoint list item status
$spoconn = Connect-PnPOnline –Url $SPSite –Credentials (Get-AutomationPSCredential -Name 'YourAutomationAccount') -ReturnConnection -Verbose
$itemupdate = Set-PnPListItem -List $SPList -Identity $SPListItemID -Values @{"TeamsCreated" = "Success"; "Link" = "https://jh365dev.sharepoint.com/sites/$teamsAlias, Link"} -Connection $spoconn

Write-Output "All done.."


