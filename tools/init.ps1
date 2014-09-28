param($installPath, $toolsPath, $package, $project)

Write-Host "Hello $name!"

# Configure
$moduleName = "VSNugetProjectTools"

# Derived variables
$psdFileName = "$moduleName.psd1"
$psmFileName = "$moduleName.psm1"

Write-Host "ToolsPath $toolsPath"
$psd = (Join-Path $toolsPath $psdFileName)
$psm = (Join-Path $toolsPath $psmFileName)

Write-Host "psd $psd"

# Check if the NuGet_profile.ps1 exists and register the Mod1.psd1 module
if(!(Test-Path $profile)){
	mkdir -force (Split-Path $profile)
	New-Item $profile -Type file -Value "Import-Module $moduleName -DisableNameChecking"
}
else {
    if(!(Get-Content $profile | Select-String "Import-Module $moduleName" -quiet)){
	    Add-Content -Path $profile -Value "`r`nImport-Module $moduleName -DisableNameChecking"
    }
}



# Copy the files to the module in the profile directory
$profileDirectory = Split-Path $profile -parent
$profileModulesDirectory = (Join-Path $profileDirectory "Modules")
$moduleDir = (Join-Path $profileModulesDirectory $moduleName)
if(Test-Path $moduleDir){
    # If you install this package and the dir exists, then you're upgrading...
    Remove-Item -Recurse -Force $moduleDir
}
if(!(Test-Path $moduleDir)){
	mkdir -force $moduleDir
}
copy $psd (Join-Path $moduleDir $psdFileName)
copy $psm (Join-Path $moduleDir $psmFileName)

# Copy additional files
# copy "$toolsPath\*.xsd" $moduleDir
# copy "$toolsPath\*.xml" $moduleDir


. $profile


Write-Host "Module installed, I think again"
