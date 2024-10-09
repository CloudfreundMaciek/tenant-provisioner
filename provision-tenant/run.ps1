using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request.
$SiteUrl = $Request.Body.SiteUrl
$AppId = $Request.Body.AppId

#if (-not $name) {
#    $name = $Request.Body.Name
#}

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

#Connect to SharePoint Online App Catalog site
Connect-PnPOnline -Url $SiteUrl -ManagedIdentity -UserAssignedManagedIdentityClientId $AppId

Add-PnPApp -Path provision-tenant/aces-packages/demo-ace-1.sppkg -Publish -SkipFeatureDeployment

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
