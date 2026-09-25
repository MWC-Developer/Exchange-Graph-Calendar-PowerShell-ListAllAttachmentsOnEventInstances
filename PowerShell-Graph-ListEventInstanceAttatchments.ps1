# PowerShell-Graph-ListEventInstanceAttatchments.ps1
# Lists the attachments on the second instance of a recurring meeting series.  
# Note: You can chanage the instance below - look for #$AnInstance = $Instances[1]

# =========================
# CONFIG
# =========================
$TenantId     = "<TENANT_ID>"       # TODO  Set to your tenant ID
$ClientId     = "<CLIENT_ID>"       # TODO  Set to your app registration's client ID
$ClientSecret = "<CLIENT_SECRET>"   # TODO  Set to your app registration's client secret
$Mailbox = "someuser@contoso.com"   # TODO  Set to the mailbox you want to access (app-only requires proper app permissions)
$SeriesMasterId = "AAMkAGEyMT..."   # TODO - set to the recurring event's series master ID (Graph ID, not the Outlook EntryID)

 
# =========================
# GET ACCESS TOKEN
# =========================
$TokenBody = @{
    client_id     = $ClientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
}

$TokenResponse = Invoke-RestMethod -Method POST `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body $TokenBody

$AccessToken = $TokenResponse.access_token

$Headers = @{
    Authorization = "Bearer $AccessToken"
}

# =========================
# STEP 1: GET INSTANCES
# =========================
$Start = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
$End   = (Get-Date).AddDays(14).ToString("yyyy-MM-ddTHH:mm:ssZ")

$InstancesUri = "https://graph.microsoft.com/v1.0/users/$Mailbox/events/$SeriesMasterId/instances?startDateTime=$Start&endDateTime=$End"
 
$InstancesResponse = Invoke-RestMethod -Method GET -Uri $InstancesUri -Headers $Headers
 

if (-not $InstancesResponse.value -or $InstancesResponse.value.Count -lt 2) {
    throw "Less than 2 instances found"
}

# =========================
# STEP 2: SORT + PICK AN INSTANCE
# =========================
$Instances = $InstancesResponse.value | Sort-Object { $_.start.dateTime }
$AnInstance = $Instances[1]  # This is zero based, so 1 is the second instance.

$InstanceId = $AnInstance.id

Write-Host "Instance ID:" $InstanceId
Write-Host "Start:" $AnInstance.start.dateTime

# =========================
# STEP 3: GET ATTACHMENTS
# =========================
$AttachmentsUri = "https://graph.microsoft.com/v1.0/users/$Mailbox/events/$InstanceId/attachments"
$AttachmentsUri 
$AttachmentsResponse = Invoke-RestMethod -Method GET -Uri $AttachmentsUri -Headers $Headers

if (-not $AttachmentsResponse.value -or $AttachmentsResponse.value.Count -eq 0) {
    Write-Host "No attachments found on this instance."
    return
}

# =========================
# STEP 4: OUTPUT ATTACHMENTS
# =========================
foreach ($att in $AttachmentsResponse.value) {
    Write-Host "-----------------------------------"
    Write-Host "Name :" $att.name
    Write-Host "Type :" $att."@odata.type"
    Write-Host "Size :" $att.size
    Write-Host "Id   :" $att.id
}

# Optional full JSON
$AttachmentsResponse.value | ConvertTo-Json -Depth 10
