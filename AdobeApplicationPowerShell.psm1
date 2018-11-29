function Invoke-AdobeApplicationJSX {
    param (
        [ValidateSet("Illustrator","InDesign")][Parameter(Mandatory)]
        $AdobeApplicationName,
        
        [Parameter(Mandatory,ValueFromPipeline,ParameterSetName="JSXFilePath")]
        $JSXFilePath,
        
        [Parameter(Mandatory,ValueFromPipeline,ParameterSetName="JSXFileContent")]
        $JSXFileContent,

        $AdobeApplicationCOMObject
    )
    begin {
        if (-not $AdobeApplicationCOMObject) {
            $AdobeApplicationCOMObject = New-Object -ComObject "$AdobeApplicationName.Application"
            $AdobeApplicationOpenedWithinFunction = $True
            Start-Sleep -Seconds 1
        }
    }
    process {
        if (-not $JSXFilePath) {
            $JSXFilePath = [IO.Path]::GetTempFileName() -replace "\.tmp",".jsx" #InDesign errors if extension is not jsx
            $JSXFileContent | Out-File -FilePath $JSXFilePath
        }

        if ($AdobeApplicationName -eq "InDesign") {
            $AdobeApplicationCOMObject.DoScript($JSXFilePath, 1246973031)
        } elseif ($AdobeApplicationName -eq "Illustrator") {
            $AdobeApplicationCOMObject.DoJavaScriptFile($JSXFilePath)
        }
    }
    end {
        if ($AdobeApplicationOpenedWithinFunction) {
            $AdobeApplicationCOMObject.Quit()
        }
    }

}