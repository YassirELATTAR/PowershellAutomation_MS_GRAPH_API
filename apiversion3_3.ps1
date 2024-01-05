
####-----------Modules needed:
#Install-Module -Name AzureAD
#Import-Module -Name AzureAD
#Install-Module Microsoft.Graph
#Import-Module Microsoft.Graph

#Edit your credentials here:
$login = "USER@TENANT.onmicrosoft.com"
$Psswd = 'PASSWORD'
$securString = ConvertTo-SecureString $Psswd  -AsPlainText -Force
$UserCredential = New-Object System.Management.Automation.PSCredential ($login, $securString)
Connect-ExchangeOnline -Credential $UserCredential
Connect-MsolService  -Credential $UserCredential

Connect-AzureAD -Credential $UserCredential
Connect-MgGraph -Scopes "Application.ReadWrite.All","User.ReadWrite.All","Directory.ReadWrite.All"

####PERMISSIONS FOR BOUNCE RETRIEVING:  Mailbox


#Let it rest meskin:
Start-Sleep -Seconds 3
$myAPP = New-AzureADApplication -DisplayName "Yass_API"

$mysecretkey = New-AzureADApplicationPasswordCredential -ObjectId $myapp.ObjectID -CustomKeyIdentifier "Yass_Keys" -StartDate $startDate -EndDate $endDate


#Show App after creation:
Write-Host "The API details:" -ForegroundColor Yellow
Write-Host $myAPP -ForegroundColor Green
Write-Host "The API ID is:" -ForegroundColor Yellow
Write-Host $myAPP.AppId -ForegroundColor Green



#Show SecretKEY Value:
Write-Host "The secret key value for use:" -ForegroundColor Yellow
Write-Host $mysecretkey.Value -ForegroundColor Green


#Let it rest meskin:
Start-Sleep -Seconds 5


# Find the Microsoft Graph Service Principal:
$GraphServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"

# Find the User.ReadWrite.All permission
$UserReadWriteAllPermission = $GraphServicePrincipal.AppRoles | Where-Object { $_.Value -eq 'User.ReadWrite.All' }

# Find the Directory.ReadWrite.All permission:
$DirectoryReadWriteAllPermission = $GraphServicePrincipal.AppRoles | Where-Object { $_.Value -eq 'Directory.ReadWrite.All' }

# Find the Application.ReadWrite.All Permission:
$ApplicationReadWriteAllPersmission = $GraphServicePrincipal.AppRoles | Where-Object { $_.Value -eq 'Application.ReadWrite.All' }

# Create the RequiredResourceAccess object:
$RequiredResourceAccess = [PSCustomObject]@{
    resourceAppId = $GraphServicePrincipal.AppId
    resourceAccess = @(
        @{
            id = $UserReadWriteAllPermission.Id
            type = "Role"
        },
        @{
            id = $DirectoryReadWriteAllPermission.Id
            type = "Role"
        },
        @{
            id = $ApplicationReadWriteAllPersmission.Id
            type = "Role"
        }
    )
}



#Let it rest meskin:
Start-Sleep -Seconds 15

# Update the application:
$Uri = "https://graph.microsoft.com/v1.0/applications/$($myAPP.ObjectID)"
$Body = @{
    requiredResourceAccess = @($RequiredResourceAccess)
} | ConvertTo-Json -Depth 10


$retryCount = 0
$maxRetries = 5
while ($retryCount -lt $maxRetries) {
    try {
        $result = Invoke-MgGraphRequest -Method PATCH -Uri $Uri -Body $Body
        Write-Host "Sucess adding permissions"  -ForegroundColor Green
        break
    }
    catch {
        $result
        Write-Host "Failed to add permissions"  -ForegroundColor Red
        $retryCount++
        Start-Sleep -Seconds 5
    }
}

#Let it rest meskin:
Start-Sleep -Seconds 5


# Get the Tenant ID:
$TenantId = (Get-MgContext).TenantId

# Get APP ID
$AppId = $myAPP.AppId

# Grant admin consent:
$Scope = "https://graph.microsoft.com/.default"
$Uri = "https://login.microsoftonline.com/$TenantId/adminConsent?client_id=$AppId&scope=$Scope"

Write-Host "Please open the following URL in a web browser, and sign in with an admin account to grant admin consent:" -ForegroundColor Yellow
Write-Host $Uri  -ForegroundColor Cyan

Write-Host "Check on the following link if it is workign and if all works, go to next script ;)" -ForegroundColor Yellow


#To make sure the permissions and everythng went well in case of doubt:
$verificationURI = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($myAPP.AppId)/isMSAApp~/false?Microsoft_AAD_IAM_legacyAADRedirect=true"
Write-Host "Verify Here:" -ForegroundColor Yellow
Write-Host $verificationURI -ForegroundColor Cyan

#Let it rest meskin:
Start-Sleep -Seconds 2
Disconnect-AzureAD
Disconnect-MgGraph
Clear-AzContext -Force

Write-Host "DISCONNECTED FROM THE ACCOUNT! ;) SO LONG..." -ForegroundColor Yellow -BackgroundColor DarkRed