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

New-Object -ComObject InDesign.idScriptLanguage.idJavascript



$ind = new-object -comobject InDesign.Application.CC.2018
$ind.DoScript('C:\Users\c.magnuson\test.jsx', 1246973031);
$ind.DoScript('C:\Users\c.magnuson\AppData\Local\Temp\tmp2707.jsx', 1246973031);
