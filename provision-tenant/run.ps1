using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

#if (-not $name) {
#    $name = $Request.Body.Name
#}

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

# Getting the urls and site name
    $SourceSiteURL = $Request.Body.SourceSiteUrl #Read-Host "Enter the source site url"
    $TargetTenantURL = $Request.Body.TargetTenantUrl #Read-Host "Enter the target tenant's site (e. g. https://sgtqr.sharepoint.com)"
    $TargetTenantAppId = $Request.Body.TargetTenantAppId   
    $TargetNewSiteOwnerEmail = $Request.Body.TargetNewSiteOwnerEmail
    $NewSiteTitle = $Request.Body.TargetNewSiteTitle   

$NewSiteUrl = $TargetTenantURL + "/sites/" + $NewSiteTitle
$source_tenant_app_id = "7a479457-a9ff-497f-be5f-40049df14d58"
$target_tenant_app_id = $TargetTenantAppId

Write-Host $SourceSiteURL $TargetTenantURL $NewSiteTitle $NewSiteUrl

try {
    # Connecting the source site and the target tenant
    Write-Host "Log in to the source site."
    $SourceSiteConnection = Connect-PnPOnline -Url $SourceSiteURL -ManagedIdentity -ReturnConnection -UserAssignedManagedIdentityClientId $source_tenant_app_id

    # Write-Host "Log in to the target site."
    # $TargetTenantConnection = Connect-PnPOnline -URL $TargetTenantURL -ManagedIdentity -ReturnConnection -UserAssignedManagedIdentityClientId $target_tenant_app_id
    
    # Creating requested site
    # Write-Host "Creating new site at " $NewSiteUrl
    # New-PnPSite -Type CommunicationSite -Classification "CommunicationSite" -SiteDesign Blank -Wait -TimeZone UTCPLUS0100_AMSTERDAM_BERLIN_BERN_ROME_STOCKHOLM_VIENNA -Description "New site" -Lcid 1033 -PreferredDataLocation DEU -Title $NewSiteTitle -Url $NewSiteUrl -Owner $TargetNewSiteOwnerEmail -Connection $TargetTenantConnection -Verbose
    # Write-Host "Created new site with url: "$NewSiteUrl

    # Letting the user turn the custom scripts on
    # Write-Host "Turn the scripts on the target site now on."
    # Read-Host

    # Connecting the new site
    Write-Host "Logging to the new site"
    $NewSiteConnection = Connect-PnPOnline -URL $NewSiteUrl -ManagedIdentity -ReturnConnection -UserAssignedManagedIdentityClientId $target_tenant_app_id
    
    # Copying the template and script
    #Set-PnPTenantSite -Url $SourceSiteURL -DenyAddAndCustomizePages:$false -Connection $SourceSiteConnection
    Write-Host "Getting the template"
    $Template = Get-PnPSiteTemplate -OutputInstance -ListsToExtract "DA1EF9AB-E1F6-40CA-A798-0E3B3566ACF9", "664833a2-2b0d-484a-9ad4-179f18a6882e", "069dac20-2fd8-4549-b967-b886f5a29201", "dab8ca49-f92c-430f-8579-e3f7785905a1", "6b265a97-d0f7-44a3-b5fb-c2fefe761d76", "d6babeb0-eaf5-41f9-8e8a-608175230702" -IncludeNativePublishingFiles -ExcludeContentTypesFromSyndication -ExcludeHandlers SiteSecurity -Connection $SourceSiteConnection -PersistBrandingFiles -PersistPublishingFiles #-IncludeAllPages
    # $Script = Get-PnPSiteScriptFromWeb -Url $SourceSiteURL -IncludeAll -Connection $SourceSiteConnection
    Write-Host "Got the template and script!"
    
    # Customizing the new site
    Write-Host "Implementing the template"
    Set-PnPTenantSite -Url $NewSiteUrl -DenyAddAndCustomizePages:$false -Connection $NewSiteConnection
    Invoke-PnPSiteTemplate -InputInstance $Template -Connection $NewSiteConnection
    # Write-Host "Implementing the script"
    # Invoke-PnPSiteScript -WebUrl $NewSiteUrl -Script $Script -Connection $TargetTenantConnection -ErrorAction Continue
    Write-Host "Utilized the template and script!"
    
    # Set-PnPHomeSite -HomeSiteUrl $NewSiteUrl -Connection $TargetTenantConnection
    # Write-Host "Set the new site as home site!"
    
    <#
    # Deploying all adaptive cards
    $ACEs = Get-ChildItem -Path ./aces_packages
    foreach ($ACE in $ACEs) {
        $AcePath = "./aces_packages/" + $ACE.Name
        try {
            Add-PnPApp -Path $AcePath -Publish -SkipFeatureDeployment -Connection $TargetTenantConnection
        }
        catch {
            Write-Error $_.Exception
            Write-Host "The ACE" $ACE.Name "might have been already deployed or another error ocurred. Skipping to a next one..."
        }
    }
    #>

    Write-Host "Finished!"
    Write-Host "What is left to do is changing the site theme, creating a Dashboard - both in site settings bar - and lastly add chosen ACEs on the Dashboard itself. Have fun!"
}
catch {
    Write-Error $_.Exception
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
