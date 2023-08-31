<#
    .SYNOPSIS
    Manage Self-Service Purchase Policies using the MSCommerce PowerShell module.

    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    This script connects to the MSCommerce module and manages self-service purchase policies for products based on predefined patterns. It disables self-service purchase for all products if not already disabled.

    .NOTES
    Requires MSCommerce PowerShell module
    Only works in PowerShell 5.1

    .EXAMPLE
    Run from within the repo root as working directory:
    . .\Scripts\Licensing\DisableSelfServicePurchasing.ps1

#>

#requires -Module MSCommerce

Connect-MSCommerce

$defaultPolicy = Get-MSCommercePolicy -PolicyId AllowSelfServicePurchase

$selfServicePolicies = Get-MSCommerceProductPolicies -PolicyId $defaultPolicy.PolicyId | Out-String

# Define a regular expression pattern to match ProductId values
$pattern = "CFQ7TTC0[A-Z0-9]{4}"

# Find all matches in the string using the regex pattern
$productIdMatches = [regex]::Matches($selfServicePolicies, $pattern)

foreach ($productId in $productIdMatches) {

    $policyValue = ($selfServicePolicies[($productId.Index - $productId.Length)..($productId.Index - 1)] -join "").Trim()

    $productNameIndex = $productId.Index + ($productId.Length + ($defaultPolicy.PolicyId | Out-string).Trim().Length) + 2

    $productName = ($selfServicePolicies[$productNameIndex..($productNameIndex + 99)] -join "").Split("`n")[0].Trim()

    $productId = $productId.Value.Trim()

    if ($policyValue -eq "Disabled") {

        Write-Host "Self-service purchase is already disabled for $($productId) ($productName)" -ForegroundColor Green

    }

    else {

        Write-Host "Self-service purchase is not yet disabled for $($productId) ($productName)" -ForegroundColor Yellow

        try {

            Update-MSCommerceProductPolicy -PolicyId $defaultPolicy.PolicyId -ProductId $productId -Value Disabled -ErrorAction Stop

            Write-Host "Disabled self-service purchase for $($productId) ($productName)" -ForegroundColor Green

        }
        catch {

            Write-Host "Failed to disable self-service purchase for $($productId) ($productName)" -ForegroundColor Red

            # $Error[-1]

        }

    }

}