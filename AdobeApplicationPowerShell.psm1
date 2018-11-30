function Invoke-AdobeApplicationJSX {
    param (
        [ValidateSet("Illustrator","InDesign")][Parameter(Mandatory)]
        $AdobeApplicationName,

        [Parameter(Mandatory,ValueFromPipeline,ParameterSetName="JSXFileContent")]
        $JSXFileContent,

        $AdobeApplicationCOMObject
    )
    begin {
        if (-not $AdobeApplicationCOMObject) {
            $AdobeApplicationCOMObject = New-Object -ComObject "$AdobeApplicationName.Application"
            $AdobeApplicationOpenedWithinFunction = $True
            Start-Sleep -Seconds 1 #Might not be needed, InDesign will pop an alert over the splash screen and close without alert achnolweldgement
        }
    }
    process {
        if ($AdobeApplicationName -eq "InDesign") {
            $AdobeApplicationCOMObject.DoScript($JSXFileContent, 1246973031)
        } elseif ($AdobeApplicationName -eq "Illustrator") {
            $AdobeApplicationCOMObject.DoScript($JSXFileContent, 1246973031)
            # $AdobeApplicationCOMObject.DoJavaScriptFile($JSXFilePath)
        }
    }
    end {
        if ($AdobeApplicationOpenedWithinFunction) {
            $AdobeApplicationCOMObject.Quit()
        }
    }
}