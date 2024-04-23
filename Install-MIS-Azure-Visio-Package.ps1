#########################################################
#                                                       #
# Install Microsoft Integration & Azure Stencils Pack   #
# Author: Sandro Pereira                                #
#                                                       #
#########################################################

[String]$location = $psscriptroot
[String]$destination = Get-ChildItem HKCU:\Software\Microsoft\Office\ -Recurse | Where-Object { $_.PSChildName -eq 'Application' } | Get-ItemProperty -Name MyShapesPath | Select-Object -ExpandProperty MyShapesPath

$files = Get-ChildItem $location -Recurse -Force -Filter *.vssx
foreach ($file in $files) {
    if ($file.PSPath.Contains('Previous Versions') -eq $false) {
        Copy-Item -Path $file.PSPath -Destination $destination -Force
    }
}
