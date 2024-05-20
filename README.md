# RFL.Microsoft.Azure
Automatic create an Azure (Entra) Documentation to simplify the life of admins and consultants.

# Azure Documentation
Automatic create an Azure (Entra) Documentation to simplify the life of admins and consultants.

# Usage
ExportAzureData.ps1 - On a device with internet connectivity, open a PowerShell 5.1 session (not PowerShell core), it will connect to the Azure Services and export the data to a word or html format. It uses the PScribo PowerShell module (https://github.com/iainbrighton/PScribo) and Microsoft Azure PowerShell modules Az.Accounts, Az.Resources, Az.Compute, Az.Network, Az.Storage, Az.Monitor, Az.Billing, Az.ResourceGraph, Az.RecoveryServices, Az.Reservations, Az.StorageSync, Az.PolicyInsights, Az.LogicApp, Az.KeyVault (https://learn.microsoft.com/en-us/powershell/module/?view=azps-11.6.0). It can be run from any Windows Device (Workstation, Server). As it uses external PowerShell module, it is recommended not to run from a Domain Controller.


# Pre-Requisites
Internet Access and the required PowerShell modules

# Examples
Example01: Exports all the Data in a word format for the Tenant tenant.com and will save the file to c:\temp folder

**.\ExportAzureData.ps1 -TenantId 'tenant.com' -OutputFolderPath "c:\temp" -ExportAll**

Example 02: Exports the Virtual Machines Data in a HTML format for the Tenant tenant.com and will save the file to c:\temp folder

**.\ExportAzureData.ps1 -BetaAPI -OutputFormat @('HTML') -TenantId 'tenant.com' -OutputFolderPath "c:\temp" -ExportVirtualMachines**

Example 03: Exports the Virtual Machines Data in a HTML format for the Tenant tenant.com without exporting Costs and Metric information and will save the file to c:\temp folder

**.\ExportAzureData.ps1 -BetaAPI -OutputFormat @('HTML') -TenantId 'tenant.com' -OutputFolderPath "c:\temp" -ExportVirtualMachines -ExportCostInformation $false -ExportWithMetrics $false**

# Documentation
Access our Wiki at RFL.Microsoft.Azure/wiki

# Issues and Support
Access our Issues at RFL.Microsoft.Azure/issues
