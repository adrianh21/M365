# JumpCloud API call
$jcApiKey = "jca_MkCju1SXueovuQnBPSAN1D55MzAfZVx3BhRO"
$jcHeaders = @{ "x-api-key" = $jcApiKey }
$jcUrl = "https://console.jumpcloud.com/api/v2/systeminsights/system_info"

$response = Invoke-RestMethod -Uri $jcUrl -Headers $jcHeaders
$devices = $response  # Adjust if paginated

# Microsoft Graph Auth
$tenantId = "3d08c29d-04bc-44cc-a867-3238f9d6e401"
$clientId = "7cbd9ee5-a7f7-4f70-9f96-ea8e6cead4e4"
$clientSecret = "your_client_secret"

$body = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $clientId
    client_secret = $clientSecret
}

$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $body
$accessToken = $tokenResponse.access_token
$graphHeaders = @{ Authorization = "Bearer $accessToken" }

# Loop through devices
foreach ($device in $devices) {
    $deviceName = $device.computer_name
    $cpu = $device.cpu_brand
    $model = $device.model
    $serial = $device.serial_number
    $vendor = $device.vendor

    # Search for existing item in Microsoft List
    $siteId = "your_site_id"
    $listId = "your_list_id"
    $filter = "fields/Title eq '$deviceName'"
    $searchUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items?\$filter=$filter"

    $existing = Invoke-RestMethod -Uri $searchUrl -Headers $graphHeaders

    if ($existing.value.Count -gt 0) {
        # Update existing item
        $itemId = $existing.value[0].id
        $updateUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items/$itemId"
        $updateBody = @{
            fields = @{
                CPU = $cpu
                Model = $model
                SerialNumber = $serial
                Vendor = $vendor
            }
        } | ConvertTo-Json -Depth 3

        Invoke-RestMethod -Uri $updateUrl -Headers $graphHeaders -Method PATCH -Body $updateBody
    } else {
        # Create new item
        $createUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items"
        $createBody = @{
            fields = @{
                Title = $deviceName
                CPU = $cpu
                Model = $model
                SerialNumber = $serial
                Vendor = $vendor
            }
        } | ConvertTo-Json -Depth 3

        Invoke-RestMethod -Uri $createUrl -Headers $graphHeaders -Method POST -Body $createBody
    }
}
