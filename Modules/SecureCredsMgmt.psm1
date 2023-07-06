function New-MZZEncryptedPassword {
    param (
        [Parameter(Mandatory=$false)][string]$FileName,
        [Parameter(Mandatory=$false)][string]$AdminUser = $env:USERNAME
    )

    if (!$localRepoPath) {

        $localRepoPath = git rev-parse --show-toplevel

    }
    
    $secureCredsFolder = "$localRepoPath\.local\SecureCreds"

    if (!(Test-Path -Path $secureCredsFolder)) {

        New-Item -Path $secureCredsFolder -ItemType Directory

    }

    $SecureStringPassword = Read-Host "Please enter the password you would like to hash" -AsSecureString
    
    $PasswordHash = $SecureStringPassword | ConvertFrom-SecureString

    if ($FileName) {

        Set-Content -Path "$secureCredsFolder\$($FileName).txt" -Value $PasswordHash -Force

    }

    else {

        Set-Content -Path "$secureCredsFolder\$($adminUser).txt" -Value $PasswordHash -Force

    }

}

function Get-MZZSecureCreds {
    param (
        [Parameter(Mandatory=$false)][switch]$checkPassword,
        [Parameter(Mandatory=$false)][switch]$updatePassword,
        [Parameter(Mandatory=$false)][string]$FileName,
        [Parameter(Mandatory=$false)][string]$AdminUser = $env:USERNAME
    )

    if (!$localRepoPath) {

        $localRepoPath = git rev-parse --show-toplevel

    }

    $secureCredsFolder = "$localRepoPath\.local\SecureCreds"

    if (!(Test-Path -Path $secureCredsFolder)) {

        New-Item -Path $secureCredsFolder -ItemType Directory

    }

    if ($FileName) {

        if ($updatePassword) {

            New-MZZEncryptedPassword -fileName $FileName

        }

        if (!(Test-Path -Path ".local\SecureCreds\$FileName.txt")) {

            Write-Host "No password found for filename: $FileName..." -ForegroundColor Yellow

            . New-MZZEncryptedPassword -fileName $FileName

            . Get-MZZSecureCreds -fileName $FileName

        }

        else {

            $passwordEncrypted = Get-Content -Path "$localRepoPath\.local\SecureCreds\$($FileName).txt" | ConvertTo-SecureString

            if (!$passwordEncrypted) {

                . New-MZZEncryptedPassword -fileName $FileName

            }

            $global:passwordDecrypted = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordEncrypted))
            
            if ($checkPassword) {

                Write-Host "Decrypted password: $passwordDecrypted" -ForegroundColor Cyan

            }

        }

        Write-Host "Password is stored in `$passwordDecrypted variable and in clipboard!" -ForegroundColor Yellow

        $passwordDecrypted | Set-Clipboard

        return $passwordDecrypted > $null

    }

    else {

        if ($updatePassword) {

            New-MZZEncryptedPassword

        }


        if (!(Test-Path -Path "$LocalRepoPath\.local\SecureCreds\$adminUser.txt")) {

            Write-Host "No credentials found for user: $adminUser..." -ForegroundColor Yellow

            . New-MZZEncryptedPassword

            . Get-MZZSecureCreds

        }

        else {

            $adminPasswordEncrypted = Get-Content -Path "$localRepoPath\.local\SecureCreds\$($adminUser).txt" | ConvertTo-SecureString

            if (!$adminPasswordEncrypted) {

                . New-MZZEncryptedPassword

            }

            Write-Host $adminUser -ForegroundColor Green

            $global:secureCreds = New-Object System.Management.Automation.PSCredential -ArgumentList $adminUser,$adminPasswordEncrypted

            if ($checkPassword) {

                $adminPasswordDecrypted = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureCreds.Password))
                Write-Host "Decrypted password: $adminPasswordDecrypted" -ForegroundColor Cyan

            }

        }

        Write-Host "Credentials are stored in `$secureCreds variable!" -ForegroundColor Cyan

        return $secureCreds

    }

}