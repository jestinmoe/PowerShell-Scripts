# Install necessary modules
if (-not (Get-Module -Name Microsoft.Graph* -ListAvailable)) {
    Install-Module -Name Microsoft.Graph -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
}
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
}

# Import necessary powershell modules
Import-Module Microsoft.Graph.Teams
Import-Module ExchangeOnlineManagement

# Set variables 
$m365GroupName = read-host "Enter the new M365 group" 

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline
Write-Host "Connected to Exchange Online!"

# Get members of the AD group


# Check if the M365 group exists, if not, create it. 
if (-not (Get-UnifiedGroup -Identity $m365GroupName -ErrorAction SilentlyContinue)) {
    $m365GroupDescription = read-host "Enter Description of new M365 Group"
    New-UnifiedGroup -DisplayName $m365GroupName -notes $m365GroupDescription -AccessType private -RequireSenderAuthenticationEnabled $true -EmailAddresses "$m365GroupName@aebs.com" 
    Set-UnifiedGroup -Identity $m365GroupName -UnifiedGroupWelcomeMessageEnabled:$false -HiddenFromAddressListsEnabled:$false -HiddenFromExchangeClientsEnabled:$true 

    #Assign group membership from existing AD group
    $adGroupName = read-host "Enter the source AD group to add membership to new M365 group - If no AD Group, press enter." 
    if (( $null -eq $adGroupName) -or ($adGroupName -eq '')){
        Write-Host "No source AD group provided"
    }
    else {
    # $adGroupMembers = Get-ADGroupMember -Identity $adGroupName | Get-ADUser -Properties * |  Select-Object -Property UserPrincipalName
    foreach ($member in $adGroupMembers) {
        Add-UnifiedGroupLinks -Identity $m365GroupName -LinkType members -Links $member
    }
    }
    Write-Host "M365 Group '$m365GroupName' created successfully."

    $CreateMSTeam = Read-Host "Would you like to create a MS Team associate with the newly created group? yes/no"
        if ($CreateMSTeam -eq 'yes') {
            $m365GroupId = Get-UnifiedGroup -Identity $m365GroupName | Select-Object ExternalDirectoryObjectId
            $m365GroupId = $m365GroupId.ExternalDirectoryObjectId
            # new-team -GroupId $m365GroupId.ExternalDirectoryObjectId -Visibility private -Description $m365GroupDescription
            # Connect to Teams Graph API and create team with below paramaters

            Connect-MgGraph -NoWelcome
            $params = @{
	            "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('b90c28b2-67ae-4551-bcbe-6826651ed53b')"
	            "group@odata.bind" = "https://graph.microsoft.com/v1.0/groups('$m365GroupId')"

                memberSettings = @{
                    allowCreateUpdateChannels = $false
                    allowDeleteChannels = $false
                    allowAddRemoveApps = $false
                    allowCreateUpdateRemoveTabs = $false
                    allowCreateUpdateRemoveConnectors = $false
                }
                guestSettings = @{
                    allowCreateUpdateChannels = $false
                    allowDeleteChannels = $false
                }
                funSettings = @{
                    allowGiphy = $true
                    giphyContentRating = "Moderate"
                    allowStickersAndMemes = $true
                    allowCustomMemes = $true
                }
                messagingSettings = @{
                    allowUserEditMessages = $true
                    allowUserDeleteMessages = $false
                    allowOwnerDeleteMessages = $false
                    allowTeamMentions = $true
                    allowChannelMentions = $true
                }
            }
            New-MgTeam -BodyParameter $params
            Write-Host "MS Team created successfully."
        }
        else {
            Write-Host "Creation of MS Team skipped."
        }
} 
else {
    Write-Host "$m365GroupName already exists in AzureAD directory. Please review group name to create new group"
}



