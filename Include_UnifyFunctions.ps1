###
###  Include_UnifyFunctions.ps1
###
###  Written by Carol Wapshere
###
###  Functions used by the Document_FIMPortal and Document_SyncConfig scripts to add Unify-only functions. Must be in the same folder.
###  This include script is NOT to be shared outside Unify.
###
###  Requires the Unify.MIIS.Configuration.ToCSV.xslt stylesheet to be in the XSLT folder.
###


Function ConvertCFToCSV
{
    PARAM([string]$ScriptFldr,[string]$XSLTFile = "Unify.MIIS.Configuration.ToCSV.xslt",[string]$CFConfigFile)
    END
    {
        if (-not $CFConfigFile) {Throw "The full path to the Codeless Framework configuration file must be provided in parameter CFConfigFile"}
        if (-not (Test-Path $CFConfigFile)) {Throw "Failed to find Codeless Framework configuration file at $CFConfigFile"}
        $XSLTPath = $ScriptFldr + "\XSLT\" + $XSLTFile
        if (-not (Test-Path $XSLTPath)) {Throw "Failed to find Codeless Framework Stylesheet at $XSLTPath"}

        $xslt = new-object system.xml.xsl.XslTransform
        $xslt.load($XSLTPath)
        $TargetFile = (split-path $CFConfigFile) + "\CF.csv"
        $xslt.Transform($CFConfigFile,$TargetFile)   
        $CFCSV = import-csv $TargetFile -Delimiter ";"

        Remove-Item $TargetFile
        Return $CFCSV
    }
}


Function HashCFConstants
{
    PARAM([string]$ScriptFldr,[string]$CFConfigFile)
    END
    {
        if (-not $CFConfigFile) {Throw "The full path to the Codeless Framework configuration file must be provided in parameter CFConfigFile"}
        if (-not (Test-Path $CFConfigFile)) {Throw "Failed to find Codeless Framework configuration file at $CFConfigFile"}

        $hashCFConstants = @{}

        [xml]$CF = get-content $CFConfigFile
        foreach ($Constant in $CF.UnifyConfiguration."miis-global".constant)
        {
            $hashCFConstants.Add($Constant.name,$Constant.value)
        }
    
        Return $hashCFConstants
    }
}

