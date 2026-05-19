<#
.SYNOPSIS
    Bulk-create Microsoft 365 users from a CSV file using Microsoft Graph.

.DESCRIPTION
    This script reads a CSV file containing user details and creates
    Azure AD / Microsoft 365 accounts, assigns licenses, sets usage
    location, and forces password change on next login.

.CSV FORMAT (example):
    DisplayName,UserPrincipalName,GivenName,Surname,JobTitle,Department,UsageLocation,LicenseSku
    John Smith,john.smith@tenant.onmicrosoft.com,John,Smith,Sales Rep,Sales,AU,Microsoft365BusinessPremium

#>

# -----------------------------
# 1. Connect to Microsoft Graph
# -----------------------------
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All"
Select-MgProfile -Name "beta"

# -----------------------------
# 2. Import CSV
# -----------------------------
$csvPath = ".\users.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "CSV file not found at $csvPath" -ForegroundColor Red
    exit
}

$users = Import-Csv $csvPath
Write-Host "Loaded $($users.Count) users from CSV." -ForegroundColor Green

# -----------------------------
# 3. Loop through each user
# -----------------------------
foreach ($u in $users) {

    Write-Host "Creating user: $($u.DisplayName)..." -ForegroundColor Yellow

    # Generate a temporary password
    $tempPassword = "Temp!" + (Get-Random -Minimum 10000 -Maximum 99999)

    try {
        # Create user
        $newUser = New-MgUser -AccountEnabled $true `
            -DisplayName $u.DisplayName `
            -UserPrincipalName $u.UserPrincipalName `
            -GivenName $u.GivenName `
            -Surname $u.Surname `
            -JobTitle $u.JobTitle `
            -Department $u.Department `
            -UsageLocation $u.UsageLocation `
            -PasswordProfile @{ Password = $tempPassword; ForceChangePasswordNextSignIn = $true }

        Write-Host "User created: $($u.UserPrincipalName)" -ForegroundColor Green

        # Assign license
        if ($u.LicenseSku) {
            Write-Host "Assigning license: $($u.LicenseSku)" -ForegroundColor Cyan

            $sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $u.LicenseSku }

            if ($sku) {
                Set-MgUserLicense -UserId $newUser.Id -AddLicenses @{SkuId = $sku.SkuId} -RemoveLicenses @()
                Write-Host "License assigned." -ForegroundColor Green
            }
            else {
                Write-Host "License SKU not found: $($u.LicenseSku)" -ForegroundColor Red
            }
        }

    }
    catch {
        Write-Host "Error creating user $($u.UserPrincipalName): $_" -ForegroundColor Red
    }

    Write-Host "----------------------------------------"
}

Write-Host "Bulk user creation completed." -ForegroundColor Green
