ipmo -force AdobeApplicationPowerShell

$AdobeApplication = New-Object -ComObject InDesign.Application
Invoke-AdobeApplicationJSX -AdobeApplicationName InDesign -AdobeApplicationCOMObject $AdobeApplication -JSXFileContent @"
alert("test");
"@
Invoke-AdobeApplicationJSX -AdobeApplicationName InDesign -JSXFileContent @"
alert("test");
"@

$AdobeApplication2 = New-Object -ComObject Illustrator.Application
Invoke-AdobeApplicationJSX -AdobeApplicationName Illustrator -AdobeApplicationCOMObject $AdobeApplication2 -JSXFileContent @"
alert("test");
"@
Invoke-AdobeApplicationJSX -AdobeApplicationName Illustrator -JSXFileContent @"
alert("test");
"@

$ind = new-object -comobject InDesign.Application.CC.2018

$AdobeApplicationName = "InDesign"
$InterfaceToAdobeApplication = "Server"
$BatchNumber = "20181115-1300"
Set-CustomyzerModuleEnvironment -Name Production