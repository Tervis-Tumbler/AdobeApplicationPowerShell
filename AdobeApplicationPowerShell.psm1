$AdobeApplications = [PSCustomObject]@{
    Name = "InDesign"
},
[PSCustomObject]@{
    Name = "Illustrator"
}

function Invoke-AdobeApplicationJSX {
    param (
        [ValidateSet("Illustrator","InDesign")][Parameter(Mandator,ParameterSetName="AdobeApplicationName")]$AdobeApplicationName,
        [Parameter(Mandator,ParameterSetName="AdobeApplicationCOMObject")]$AdobeApplicationCOMObject,
        [Parameter(Mandatory,ValueFromPipeline,ParameterSetName="JSXFilePath")]$JSXFilePath,
        [Parameter(Mandatory,ValueFromPipeline,ParameterSetName="JSXFileContent")]$JSXFileContent
    )
    begin {
        if (-not $AdobeApplicationCOMObject) {
            $AdobeApplicationCOMObject = New-Object -ComObject "$AdobeApplicationName.Application"
            $AdobeApplicationOpenedWithinFunction = $True
        }
    }
    process {
        if (-not $JSXFilePath) {
            $JSXFilePath = [IO.Path]::GetTempFileName()
            $JSXFileContent | Out-File -FilePath $JSXFilePath
        }
        $AdobeApplicationCOMObject.DoJavaScriptFile($JSXFilePath)
    }
    end {
        if ($AdobeApplicationOpenedWithinFunction) {
            $AdobeApplicationCOMObject.Quit()
        }
    }

}