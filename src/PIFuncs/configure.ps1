
# check if this installed 
# Get-InstalledModule -name SharePointPnPPowerShell*
# if yes unistall
# UnInstall-Module -name <SharePointPnPPowerShell>

#Get-installedModule -name Pnp.PowerShell
# UnInstall-Module -name Pnp.PowerShell

#Check if this command run

##### install PnP.PowerShell 1.12.0 (as newer version has issues)
# Install-Module -Name "PnP.PowerShell" -RequiredVersion 1.12.0 -Force -AllowClobber

## check if this command runs
# Register-PnPAzureADApp


# execute Azure registration script

# prerequisite - need azure admin access to register a new azure function app


param (
	[Parameter(Mandatory=$true)]
	[string]
	$SiteUrl,
	[Parameter(Mandatory=$true)]
	[string]
	$Tenant,
	[Parameter(Mandatory=$true)]
	[string]
	$CertificatePassword,
	[string]$AzureADAppName = "PnP.Core.SDK.AzureFunctionSample",
	[string]$CertificateOutDir = ".\Certificates"
)

if (-not ( Test-Path -Path $CertificateOutDir -PathType Container )) {
	md $CertificateOutDir
}

$app = Register-PnPAzureADApp -ApplicationName $AzureADAppName -Tenant $Tenant -OutPath $CertificateOutDir `
	-CertificatePassword (ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force) `
	-Scopes "MSGraph.Group.ReadWrite.All","MSGraph.User.ReadWrite.All","SPO.Sites.FullControl.All","SPO.TermStore.ReadWrite.All","SPO.User.ReadWrite.All" `
	-Store CurrentUser -DeviceLogin

$localSettings = Get-Content local.settings.sample.json | ConvertFrom-JSON
$localSettings.Values.SiteUrl = $SiteUrl
$localSettings.Values.TenantId = $Tenant
$localSettings.Values.ClientId = $app.'AzureAppId/ClientId'
$localSettings.Values.CertificateThumbPrint = $app.'Certificate Thumbprint' 
$localSettings.Values.WEBSITE_LOAD_CERTIFICATES = $app.'Certificate Thumbprint'

Write-Host $localSettings

($localSettings | ConvertTo-JSON) > local.settings.json