PARAM ( [string]$ServerConfigFolder,
        [string]$StylesScript,
        [string]$SRGuidCSV,
        [string]$EBExtensibilityFolder,
        [string]$CFConfigFile,
        [string]$SingleMA,
        [boolean]$Visible=$true,
        [boolean]$FlowMapOnly=$false,
        [string]$ErrorAction="Continue")

#Copyright (c) 2014, Unify Solutions Pty Ltd
#All rights reserved.
#
#Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#* Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#
#THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
#IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; 
#OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

###
###  FIM Synchronization Configuration documenter
###
###  Written by Carol Wapshere
###
###  Takes the folder location of a Sync Server configuration export and produces a Word document.
###  - If a Codeless Framework config file is also provided it will add Provisioning sections.
###  - If a SingleMA is specified it will only report on that MA.
###
###  Notes:
###  -- Must be run on a computer with Word installed.
###  -- Tested with FIM 2010 R2 and Word 2013.
###
###  Parameters:
###    -ServerConfigFolder     (Required) The full path to the folder contains the full Sync Server config export
###    -StylesScript           (Optional) The full path to a customised version of Include_STYLES.ps1. Uses the version in the script folder if not specified. Will fail if not found.
###    -SRGuidCSV              (Optional) If using Sync Rules, a CSV matching Sync Rule name to Metaverse GUID. 
###                                       The CSV must have columns "object_id" and displayName and must be comma separated.
###                                       You can export this information from the FIMSynchronizationService database using the following SQL query:
###                                             select object_id,displayName from dbo.mms_metaverse where object_type='synchronizationRule'
###    -EBExtensibilityFolder  (Optional) The full path to folder containing the FIM Event Broker configuration files.
###    -CFConfigFile           (Optional) Unify Codeless Framework configuration file
###    -SingleMA               (Optional) The name of a single MA to only report on.
###    -Visible           (Default=$true) If set to $false does not open Word until the document is finished, which runs faster than with Visible=$true.
###    -FlowMapOnly      (Default=$false) If set to $true only the end-to-end attribute flow map is produced.
###    -ErrorAction    (Default=Continue) Will continue past errors. Set to "Stop" if you want it to stop on any error.
###
###
###  Changelog:
###    CW 15/10/2013 - Added support for CS object types joining to different MV object types
###                  - Fixed a bug with join rules not showing properly
###    CW 15/12/2013 - Replaced CF and EB rendering with a stylesheet and CSV based approach.
###    CW 30/12/2013 - Detects script folder. 
###                  - EB parsing errors suppressed.
###    CW 25/02/2014 - Fixed a bug with finding XSLT files
###    CW 16/06/2014 - Added a stylesheet to convert Sync Rules IAF/EAFs to CSV
###    CW 11/08/2014 - Merged public script with internal Unify script. Codeless framework functions moved to Include_UnifyFunctions.ps1 (not public),
###                  - Changed Import Flow Precedence table to a full end-to-end flow mapping table,
###                  - Improved attribute flow XSLT and use the resultant CSVs in building the document rather than hashtables.
###                  - Added extra script paramters - Visible, FlowMapOnly and ErrorAction
###    CW 4/9/2014   - Fixed a bug around multiple Projection rules for same object type
###    CW 25/1/2016  - Fixed a reported bug by changing "$ExportFlowCSV = """ to "$ExportFlowCSV = @()"
###
###  To Do:
###   * Sync Rule provisioning logic
###   * Run Profile details
###


### Run Shared Functions script
$ErrorActionPreference = "Stop"

$ScriptFldr = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

if ($StylesScript) {. $StylesScript} else {. $ScriptFldr\Include_STYLES.ps1}
. $ScriptFldr\Include_CustomisedContent.ps1
. $ScriptFldr\Include_WordFunctions.ps1
if (Test-Path ($ScriptFldr + "\Include_UnifyFunctions.ps1")) {. $ScriptFldr\Include_UnifyFunctions.ps1}

$ErrorActionPreference = $ErrorAction

## Read SRGuidCSV if provided
if ($SRGuidCSV)
{
    $SRNames = import-csv $SRGuidCSV
}

## Collect Hash Tables with details about each MA and the rules configured in it

$hashMADetails = @{}
$hashConnFilter = @{}
$hashJoin = @{}
$hashPrj = @{}
$hashRP = @{}
$ObjectName = @{}

foreach ($MAFile in (Get-ChildItem $ServerConfigFolder | where {$_.Name -like "MA*.XML"}))
{
    [xml]$MA = get-content "$ServerConfigFolder\$MAFile"
    $MAid = $MA."saved-ma-configuration"."ma-data".id

    $hashMADetails.Add($MAid,@{})

    # MA Name
    $MAName = $MA."saved-ma-configuration"."ma-data".name
    $hashMADetails.($MAid).Add("Name",$MAName)
    $ObjectName.Add($MAid.replace("{","").replace("}",""),$MAName)

    # MA Type
    $hashMADetails.($MAid).Add("Type",@($MA."saved-ma-configuration"."ma-data".category))
    if ($MA."saved-ma-configuration"."ma-data".subtype -ne "") {$hashMADetails.($MAid).("Type") += $MA."saved-ma-configuration"."ma-data".subtype}

    # Connection Details
    $ConnectionDetails = @()

    $PrivateConfig = $MA."saved-ma-configuration"."ma-data"."private-configuration"
    if ($PrivateConfig."adma-configuration")
    {
        $ConnectionDetails = (("Forest: " + $PrivateConfig."adma-configuration"."forest-name"),
                                ("Domain: " + $PrivateConfig."adma-configuration"."forest-login-domain"),
                                ("Account: " + $PrivateConfig."adma-configuration"."forest-login-user"))
    }
    elseif ($PrivateConfig."oledbma-configuration"."connection-info")
    {
        $ConnectionDetails = (("Server: " + $PrivateConfig."oledbma-configuration"."connection-info".server),
                ("Database: " + $PrivateConfig."oledbma-configuration"."connection-info".databasename),
                ("Table: " + $PrivateConfig."oledbma-configuration"."connection-info".tablename))
        if ($PrivateConfig."oledbma-configuration"."connection-info"."delta-tablename" -ne "")
        {$ConnectionDetails += "Delta Table: " + $PrivateConfig."oledbma-configuration"."connection-info"."delta-tablename"}
        if ($PrivateConfig."oledbma-configuration"."connection-info"."multivalued-tablename" -ne "")
        {$ConnectionDetails += "Multivalued Table: " + $PrivateConfig."oledbma-configuration"."connection-info"."multivalued-tablename"}
        $ConnectionDetails += "Account: " + $PrivateConfig."oledbma-configuration"."connection-info".domain + 
                                            "\" + $PrivateConfig."oledbma-configuration"."connection-info".user
    }
    elseif ($PrivateConfig."MAConfig"."extension-config"."connection-info"."connect-to")
    {
        $ConnectionDetails = ("Connect To: " + $PrivateConfig."MAConfig"."extension-config"."connection-info"."connect-to")
        if ($PrivateConfig."MAConfig"."extension-config"."connection-info".user) 
        {$ConnectionDetails += "Account: " + $PrivateConfig."MAConfig"."extension-config"."connection-info".user}
    }
    elseif ($PrivateConfig."MAConfig"."extension-config".attributes.attribute.name)
    {
        foreach ($item in $PrivateConfig."MAConfig"."extension-config".attributes.attribute)
        {
            if ($text -eq "") {$ConnectionDetails += $item.name + ": " + $item."#text"}
            else {$ConnectionDetails += $item.name + ": " + $item."#text"}
        }
    }
    elseif ($PrivateConfig."MAConfig"."parameter-values".parameter)
    {
        $ConnectionDetails = $PrivateConfig."MAConfig"."parameter-values".parameter.InnerXml
    }

    $hashMADetails.($MAid).Add("Connection",$ConnectionDetails)


    #Password Sync
    if ($MA."saved-ma-configuration"."ma-data"."password-sync-allowed" -eq 1) 
    {$hashMADetails.($MAid).Add("PasswordSync","Enabled")} 
    else {$hashMADetails.($MAid).Add("PasswordSync","Disabled")}


    #Rules Extension
    $hashMADetails.($MAid).Add("RulesExtension",$MA."saved-ma-configuration"."ma-data".extension."assembly-name")


    #Deprovisioning
    $hashMADetails.($MAid).Add("Deprovisioning",$MA."saved-ma-configuration"."ma-data"."provisioning-cleanup".action)


    # Object Types
    $hashMADetails.($MAid).Add("ObjectTypes",@())
    foreach ($CSObjectType in $MA."saved-ma-configuration"."ma-data"."ma-partition-data".partition."filter-hints"."object-classes"."object-class")
    {
        if ($CSObjectType.included -eq "1" -and $hashMADetails.($MAid).("ObjectTypes") -notcontains $CSObjectType.name)
        {
            $hashMADetails.($MAid).("ObjectTypes") += $CSObjectType.name
        }
    }


    #Connector Filters
    if (-not $hashConnFilter.ContainsKey($MAid)) {$hashConnFilter.Add($MAid,@{})}

    if ($MA."saved-ma-configuration"."ma-data"."stay-disconnector")
    {
        foreach($FilterSet in $MA."saved-ma-configuration"."ma-data"."stay-disconnector"."filter-set")
        {
            $CSObjectType = $FilterSet."cd-object-type"
            if (-not $hashConnFilter.($MAid).ContainsKey($CSObjectType)) {$hashConnFilter.($MAid).Add($CSObjectType,@())}

            if ($FilterSet."filter-alternative")
            {
                foreach ($filterid in $FilterSet."filter-alternative".id)
                {
                    $filtertext = ""
                    foreach ($condition in $FilterSet.SelectNodes("//filter-alternative[@id='{0}']/condition" -f $filterid))
                    {
                        $text = $condition."cd-attribute" + " " + $condition.operator + " " + $condition.value
                        if ($filtertext -eq "") {$filtertext = $text}
                        else {$filtertext = $filtertext + " AND " + $text}
                    }
                    $hashConnFilter.($MAid).($CSObjectType) += $filtertext
                }
            }
            elseif ($FilterSet.type -eq "Scripted")
            {
                $hashConnFilter.($MAid).($CSObjectType) += "scripted"
            }
        }
    }


    # Join Rules
    if (-not $hashJoin.ContainsKey($MAid)) {$hashJoin.Add($MAid,@{})}

    foreach ($CSObjectRule in $MA."saved-ma-configuration"."ma-data".join."join-profile")
    {
        $CSObjectType = $CSObjectRule."cd-object-type"
        if (-not $hashJoin.($MAid).ContainsKey($CSObjectType)) {$hashJoin.($MAid).Add($CSObjectType,@{})}

        foreach ($JoinRule in $CSObjectRule."join-criterion")
        {
            $RuleID = $JoinRule.id
            $hashJoin.($MAid).($CSObjectType).Add($RuleID,@{})
            $hashJoin.($MAid).($CSObjectType).($RuleID).Add("MVObjectType",$JoinRule.search."mv-object-type")
            $hashJoin.($MAid).($CSObjectType).($RuleID).Add("MVAttrib",$JoinRule.search."attribute-mapping"."mv-attribute")

            if ($JoinRule.search."attribute-mapping"."direct-mapping"."src-attribute")
            {
                $hashJoin.($MAid).($CSObjectType).($RuleID).Add("Type","Direct")
                $CSAttrib = @()
                foreach ($attr in $JoinRule.search."attribute-mapping"."direct-mapping"."src-attribute")
                {
                    if ($attr."#text") {$CSAttrib += $attr."#text"}
                    else {$CSAttrib += $attr}
                }
                $hashJoin.($MAid).($CSObjectType).($RuleID).Add("CSAttrib",$CSAttrib)
            }
            elseif ($JoinRule.search."attribute-mapping"."scripted-mapping"."src-attribute")
            {
                $hashJoin.($MAid).($CSObjectType).($RuleID).Add("Type","Advanced")
                $CSAttrib = @()
                foreach ($attr in $JoinRule.search."attribute-mapping"."scripted-mapping"."src-attribute")
                {
                    if ($attr."#text") {$CSAttrib += $attr."#text"}
                    else {$CSAttrib += $attr}
                }
                $hashJoin.($MAid).($CSObjectType).($RuleID).Add("CSAttrib",$CSAttrib)
            }
        }
    }

    #Projection Rules
    if (-not $hashPrj.ContainsKey($MAid)) {$hashPrj.Add($MAid,@{})}

    foreach ($CSObjectRule in $MA."saved-ma-configuration"."ma-data".projection."class-mapping")
    {
        if ($CSObjectRule) {foreach ($PrjRule in $CSObjectRule)
        {
            $RuleID = $PrjRule.id
            $CSObjectType = $PrjRule."cd-object-type"
            $hashPrj.($MAid).Add($RuleID,@{})
            $hashPrj.($MAid).($RuleID).Add("CSObjectType",$CSObjectType)
            $hashPrj.($MAid).($RuleID).Add("MVObjectType",$PrjRule."mv-object-type")
            $hashPrj.($MAid).($RuleID).Add("Type",$PrjRule.type)

            if ($PrjRule.scoping)
            {
                $ScopeRule = ""
                Try
                {
                    $ScopeRule = $ScopeRule + $PrjRule.scoping.scope.csAttribute
                    $ScopeRule = $ScopeRule + " " + $PrjRule.scoping.scope.csOperator
                    $ScopeRule = $ScopeRule + " " + $PrjRule.scoping.scope.csValue
                } Catch {}
                $hashPrj.($MAid).($RuleID).Add("Scope",$ScopeRule)
            }
        }}
    }

    # Run Profiles
    $hashRP.Add($MAid,@{})
    foreach ($RP in $MA."saved-ma-configuration"."ma-data"."ma-run-data"."run-configuration")
    {
        $hashRP.($MAid).Add($RP.id,$RP.name)
        $ObjectName.Add(($RP.id).replace("{","").replace("}",""),$RP.name)
    }
}


## Metaverse Schema
[xml]$MV = ReadFileContents -FilePath "$ServerConfigFolder\MV.xml"
$MVSchema = @{}
foreach ($MVObjectType in $MV."saved-mv-configuration"."mv-data".schema.dsml."directory-schema".class)
{
    $MVSchema.Add($MVObjectType.name,@())
    foreach ($MVAttr in $MVObjectType.attribute)
    {
        $MVSchema.($MVObjectType.name) += ($MVAttr.ref).Substring(1)
    }
}


## Import Flow Rules
## Uses the stylesheet to convert MV.XML to CSV
$XSLTFile = $ScriptFldr + "\XSLT\FlowRules.ToCSV.xslt"
$xslt = new-object system.xml.xsl.XslTransform
$xslt.load($XSLTFile)
$TargetFile = $ServerConfigFolder + "\importflows.csv"
$xslt.Transform("$ServerConfigFolder\MV.xml",$TargetFile)   
$ImportFlowCSV = import-csv $TargetFile -Delimiter ";"
Remove-Item $TargetFile


## Export Flow Rules
## Uses the stylesheet to convert the EAFs from each MA file to CSV, then concatenates the CSVs
$ExportFlowCSV = @()
foreach ($MAFile in (Get-ChildItem $ServerConfigFolder | where {$_.Name -like "MA*.XML"}))
{
    [xml]$MA = get-content "$ServerConfigFolder\$MAFile"
    $MAid = $MA."saved-ma-configuration"."ma-data".id
    $MAName = $MA."saved-ma-configuration"."ma-data".name

    $TargetFile = $ServerConfigFolder + "\" + $MAName + "-exportflows.csv"
    $xslt.Transform($MAFile.FullName,$TargetFile)   
    $FlowCSV = import-csv $TargetFile -Delimiter ";"

    $ExportFlowCSV = $ExportFlowCSV + $FlowCSV

    Remove-Item $TargetFile
}


## UNIFY only - If CF file provided convert to a flowrules CSV and load the constants into a hashtable

if ($CFConfigFile)
{
    $hashCFConstants = @{}
    $hashCFConstants = HashCFConstants -ScriptFldr $ScriptFldr -CFConfigFile $CFConfigFile

    $CFCSV = ConvertCFToCSV -ScriptFldr $ScriptFldr -CFConfigFile $CFConfigFile
}


## If Event Broker Operations file provided, collect details about Operations
if ($EBExtensibilityFolder)
{
    ## Convert XML files to CSV
    foreach ($EBFile in (Get-ChildItem $EBExtensibilityFolder | where {$_.Name -match "config.xml"}))
    {
        $XSLTFile = $null
        if ($EBFile.Name.Contains("LoggingEnginePlugInKey")) {$CSVFile = "Logging.csv";$XSLTFile = $ScriptFldr + "\XSLT\Unify.Product.EventBroker.LoggingEnginePlugInKey.extensibility.config.xslt"}
        elseif ($EBFile.Name.Contains("AgentEnginePlugInKey")) {$CSVFile = "Agents.csv";$XSLTFile = $ScriptFldr + "\XSLT\Unify.Product.EventBroker.AgentEnginePlugInKey.extensibility.config.xslt"}
        elseif ($EBFile.Name.Contains("EventBrokerPlugInKey")) {$CSVFile = "Roles.csv";$XSLTFile = $ScriptFldr + "\XSLT\Unify.Product.EventBroker.EventBrokerPlugInKey.extensibility.config.xslt"}
        elseif ($EBFile.Name.Contains("GroupEnginePlugInKey")) {$CSVFile = "Groups.csv";$XSLTFile = $ScriptFldr + "\XSLT\Unify.Product.EventBroker.GroupEnginePlugInKey.extensibility.config.xslt"}
        elseif ($EBFile.Name.Contains("OperationEnginePlugInKey")) {$CSVFile = "Operations.csv";$XSLTFile = $ScriptFldr + "\XSLT\Unify.Product.EventBroker.OperationEnginePlugInKey.extensibility.config.xslt"}

        if ($XSLTFile)
        {
            $xslt = new-object system.xml.xsl.XslTransform
            $xslt.load($XSLTFile)
            $TargetFile = $EBExtensibilityFolder + "\" + $CSVFile
            $xslt.Transform($EBFile.FullName,$TargetFile)
        }
    }

    ## Read the generated CSV files
    $EBCSVLogging = Import-Csv ($EBExtensibilityFolder + "\Logging.csv") -Delimiter ";"
    $EBCSVAgents = Import-Csv ($EBExtensibilityFolder + "\Agents.csv") -Delimiter ";"
    $EBCSVRoles = Import-Csv ($EBExtensibilityFolder + "\Roles.csv") -Delimiter ";"
    $EBCSVGroups = Import-Csv ($EBExtensibilityFolder + "\Groups.csv") -Delimiter ";"
    $EBCSVOperations = Import-Csv ($EBExtensibilityFolder + "\Operations.csv") -Delimiter ";"

    ## Build the Operations hashtable, replacing GUIDs with names
    $hashEB = @{}

    foreach ($OpName in ($EBCSVOperations.name | Get-Unique))
    {
        ## NOTE: The following suppresses errors cause by duplicate Operation names, but means the second Operation will not be in the doc
        if ($OpName -and -not $hashEB.ContainsKey($OpName))
        {
            $hashEB.Add($OpName,@{})

            foreach ($property in ($EBCSVOperations | where {$_.name -eq $OpName -and $_.category -eq 'General'}))
            {
                if (-not $hashEB.($OpName).ContainsKey($property.category)) {$hashEB.($OpName).Add($property.category,@{})}
                $hashEB.($OpName).($property.category).Add($property.property,$property.value)
            }
 
            $hashEB.($OpName).Add("Schedules",@{})
            $count = 1
            foreach ($StepName in ($EBCSVOperations | where {$_.name -eq $OpName -and $_.category -eq 'Schedules'}).step | Get-Unique)
            {
                $hashEB.($OpName).Schedules.Add($count,@{})
                foreach ($property in ($EBCSVOperations | where {$_.name -eq $OpName -and $_.category -eq 'Schedules' -and $_.step -eq $StepName}))
                {
                    $hashEB.($OpName).Schedules.($count).Add($property.property,$property.value)
                }
                $count += 1
            }
       
            $hashEB.($OpName).Add("Steps",@{})
            $count = 1
            foreach ($StepName in ($EBCSVOperations | where {$_.name -eq $OpName -and $_.category -eq 'Operations'}).step | Get-Unique)
            {
                $hashEB.($OpName).Steps.Add($count,@{})
                foreach ($property in ($EBCSVOperations | where {$_.name -eq $OpName -and $_.category -eq 'Operations' -and $_.step -eq $StepName}))
                {
                    $value = $property.value
                    if ($property.property -eq 'MA' -or $property.property -eq 'Run Profile') {$value = $ObjectName.($value)}
                    $hashEB.($OpName).Steps.($count).Add($property.property,$value)
                }
                $count += 1
            }
        }
    }
}


###
### Create New Word Document
###

$doc = StartDoc -Orientation "Landscape" -Visible $Visible
TestStyles -selection $doc

$DocItem = "Sync-Main"
WriteLine -selection $doc -Style $Heading1 -Text $DocText.($DocItem).Heading
if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

if (-not $SingleMA)
{
    ###
    ### General Sync Server Details
    ###

    $DocItem = "Sync-Global"
    WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    $Headings = @("Server",$MV."saved-mv-configuration".server)
    $table = StartTable -selection $doc -TableStyle $TableStyleTwoColumn -FontSize $TableFontSize -Headings $Headings
    $table.ApplyStyleHeadingRows = 0
    AddTableRow -table $table -row 2 -ColumnText @("Date of configuration export",$MV."saved-mv-configuration"."export-date")

    if ($MV."saved-mv-configuration"."mv-data"."password-sync"."password-sync-enabled" -eq 1) 
    {AddTableRow -table $table -row 3 -ColumnText @("Password Synchronization","Enabled")} 
    else {AddTableRow -table $table -row 3 -ColumnText @("Password Synchronization","Disabled")} 

    AddTableRow -table $table -row 4 -ColumnText @("Provisioning Type",$MV."saved-mv-configuration"."mv-data".provisioning.type)
    AddTableRow -table $table -row 5 -ColumnText @("Provisioning Rules Extension",$MV."saved-mv-configuration"."mv-data".extension."assembly-name")

    ## UNIFY only - Constants from Codeless Framework configuration file
    if ($CFConfigFile -and $hashCFConstants.Count -gt 0)
    {
        AddTableRow -table $table -row 6 -ColumnText @("Constants","")

        #Create subtable for the CF contants
        $subtable = StartTable -selection $table.cell(6,2)  -Columns 2  -Subtable $true -TableStyle $SubTableStyle
        $sr = 0

        foreach ($item in $hashCFConstants.Keys | sort)
        {
            $sr += 1
            if ($sr -eq 1) {$NotFirstRow = $false} else {$NotFirstRow = $true} 
            AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($item, $hashCFConstants.($item))
        }

    }

    #Move to the end of table
    $doc.EndOf(15)
    $doc.EndOf(6)
    $doc.MoveDown()

    ###
    ### End-to-end flows on Metaverse attribute
    ###

    $DocItem = "Sync-FlowMap"
    WriteLine -selection $doc -Style $Heading2 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    foreach ($MVObjectType in ($MVSchema.Keys |sort))
    {
        if (($ExportFlowCSV | where {$_.mvobject -eq $MVObjectType}) -or ($ImportFlowCSV | where {$_.mvobject -eq $MVObjectType}))
        {
            WriteLine -selection $doc -Style $Heading3 -Text $MVObjectType
            $table = StartTable -selection $doc -Headings @("Source Value","Metaverse Attribute","Target Value")
            $table.PreferredWidthType = 2 #wdPreferredWidthPercent
            $table.PreferredWidth = 100
            $r = 1

            foreach ($MVAttr in $MVSchema.($MVObjectType) |sort)
            {
                $IAFs = ($ImportFlowCSV | where {$_.mvobject -eq $MVObjectType -and $_.mvattr -eq $MVAttr})
                $EAFs = ($ExportFlowCSV | where {$_.mvobject -eq $MVObjectType -and $_.mvattr -eq $MVAttr})

                if ($IAFs -or $EAFs)
                {
                    $r += 1
                    AddTableRow -table $table -row $r -ColumnText @($null,$MVAttr,$null)

                    # Add a sub-table in the first column for the IAFs
                    if ($IAFs)
                    {
                        $subtable = StartTable -selection $table.cell($r,0) -Columns 5 -Subtable $true -TableStyle $SubTableStyle -AutoFit $false
                        $subtable.PreferredWidthType = 2 #wdPreferredWidthPercent
                        $subtable.PreferredWidth = 100
                        $sr = 0
                        foreach ($IAF in $IAFs | sort "precedence")
                        {
                            [string]$MAName = $hashMADetails.($IAF.ma).Name

                            if ($IAF.flowtype -like "SyncRule*" -and $SRNames)
                            {
                                $SRid = $IAF.flowtype.Replace("SyncRule ","")
                                $FlowType = $IAF.flowtype.Replace($SRid,($SRNames | where {$SRid.Contains($_."object_id")}).displayName)
                            }
                            else {$FlowType = $IAF.flowtype}

                            if ($IAF.flowtype -eq "Constant") {$CSAttr = $IAF.value}
                            else {$CSAttr = $IAF.csattr}

                            $sr += 1
                            if ($sr -eq 1) {$NotFirstRow = $false} else {$NotFirstRow = $true}
                            AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($IAF.precedence,$MAName,$IAF.csobject,$CSAttr,$Flowtype)
                        }
                    }

                    # Add a sub-table in the last column for the EAFs
                    if ($EAFs)
                    {
                        $subtable = StartTable -selection $table.cell($r,3) -Columns 4  -Subtable $true -TableStyle $SubTableStyle -AutoFit $false
                        $subtable.PreferredWidthType = 2 #wdPreferredWidthPercent
                        $subtable.PreferredWidth = 100
                        $sr = 0
                        foreach ($EAF in $EAFs | sort "ma")
                        {
                            [string]$MAName = $hashMADetails.($EAF.ma).Name

                            if ($EAF.flowtype -like "SyncRule*" -and $SRNames)
                            {
                                $SRid = $EAF.flowtype.Replace("SyncRule ","")
                                $FlowType = $EAF.flowtype.Replace($SRid,($SRNames | where {$SRid.Contains($_."object_id")}).displayName)
                            }
                            else {$FlowType = $EAF.flowtype}

                            if ($EAF.flowtype -eq "Constant") {$CSAttr = $EAF.value}
                            else {$CSAttr = $EAF.csattr}

                            $sr += 1
                            if ($sr -eq 1) {$NotFirstRow = $false} else {$NotFirstRow = $true}
                            AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($MAName,$EAF.csobject,$CSAttr,$Flowtype)
                        }
                    }

                    # Move to the end of the Table
                    $doc.EndOf(15)
                    $doc.EndOf(6)
                    $doc.MoveDown()
                }
            }
        }
    }
}

###
### Write Section per MA
###

if (-not $FlowMapOnly) {foreach ($MAid in $hashMADetails.Keys)
{
    $MAName = $hashMADetails.($MAid).("Name")
 
    if (-not $SingleMA -or $SingleMA -eq $MAName)
    {
        ### General MA Details

        WriteLine -selection $doc -Style $Heading2 -Text ("Connector: " + $MAName)
        WriteLine -selection $doc -Style $Heading3 -Text "MA Settings"

        $table = StartTable -selection $doc -TableStyle $TableStyleTwoColumn -Headings @("Name",$MAName)
        AddTableRow -table $table -row 2 -ColumnText @("Type",$hashMADetails.($MAid).("Type"))
        AddTableRow -table $table -row 3 -ColumnText @("Connection Details",($hashMADetails.($MAid).("Connection") -join "`n"))
        AddTableRow -table $table -row 4 -ColumnText @("Password Synchronization",$hashMADetails.($MAid).("PasswordSync"))
        AddTableRow -table $table -row 5 -ColumnText @("Rules Extension",$hashMADetails.($MAid).("RulesExtension"))
        AddTableRow -table $table -row 6 -ColumnText @("Deprovisioning",$hashMADetails.($MAid).("Deprovisioning"))
        AddTableRow -table $table -row 7 -ColumnText @("Object Types",($hashMADetails.($MAid).("ObjectTypes") -join "`n"))
        # Move to the end of the Table
        $doc.EndOf(15)
        $doc.EndOf(6)
        $doc.MoveDown()


        ###
        ### Connector Filters
        ###

        $DocItem = "Sync-CF"
        WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading
        
        if ($hashConnFilter.ContainsKey($MAid) -and $hashConnFilter.($MAid).count -gt 0)
        {
            if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

            $table = StartTable -selection $doc -Headings @("CS Object Type","Filter Rule")
            $r = 1

            foreach ($CSObjectType in $hashConnFilter.($MAid).Keys | sort)
            {
                $ColumnText = @($CSObjectType)
                $ColumnText += $hashConnFilter.($MAid).($CSObjectType) -join "`n"
                $r += 1
                AddTableRow -table $table -row $r -ColumnText $ColumnText
            }
            # Move to the end of the Table
            $doc.EndOf(15)
            $doc.EndOf(6)
            $doc.MoveDown()
        }
        else
        {
            if ($DocText.($DocItem).None) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}
        }


        ###
        ### Join Rules
        ###

        $DocItem = "Sync-Join"
        WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading

        # Write join rules table
        if ($hashJoin.ContainsKey($MAid) -and $hashJoin.($MAid).count -gt 0)
        {
            if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

            $table = StartTable -selection $doc -Headings @("CS Object Type","CS Attributes","Join Type","MV Object Type","MV Attribute")
            $r = 1

            foreach ($CSObjectType in $hashJoin.($MAid).Keys | sort)
            {
                foreach ($RuleID in $hashJoin.($MAid).($CSObjectType).Keys)
                {
                    $ColumnText = @()
                    $ColumnText += $CSObjectType
                    $ColumnText += $hashJoin.($MAid).($CSObjectType).($RuleID).("CSAttrib") -join "`n"
                    $ColumnText += $hashJoin.($MAid).($CSObjectType).($RuleID).("Type")
                    $ColumnText += $hashJoin.($MAid).($CSObjectType).($RuleID).("MVObjectType")
                    $ColumnText += $hashJoin.($MAid).($CSObjectType).($RuleID).("MVAttrib") -join "`n"
                    $r += 1
                    AddTableRow -table $table -row $r -ColumnText $ColumnText
                }
            }
            # Move to the end of the Table
            $doc.EndOf(15)
            $doc.EndOf(6)
            $doc.MoveDown()
        }
        else
        {
            if ($DocText.($DocItem).None) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}
        }



        ###
        ### Projection Rules
        ###

        $DocItem = "Sync-Projection"
        WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading

        if ($hashPrj.ContainsKey($MAid) -and $hashPrj.($MAid).count -gt 0)
        {
            if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

            $table = StartTable -selection $doc -Headings @("CS Object Type","MV Object Type","Scoped Sync Rule")
            $r = 1

            foreach ($RuleID in $hashPrj.($MAid).Keys)
            {
                $r += 1
                $ColumnText = @($hashPrj.($MAid).($RuleID).("CSObjectType"),$hashPrj.($MAid).($RuleID).("MVObjectType"))
                if ($hashPrj.($MAid).($CSObjectType).("Scope")) {$ColumnText += $hashPrj.($MAid).($RuleID).("Scope") -join ","}
                AddTableRow -table $table -row $r -ColumnText $ColumnText
            }
            # Move to the end of the Table
            $doc.EndOf(15)
            $doc.EndOf(6)
            $doc.MoveDown()
        }
        else
        {
            if ($DocText.($DocItem).None) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}
        }



        ###
        ### Import Attribute Flows
        ###
        $DocItem = "Sync-IAF"
        WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading

        # Write IAF Table
        $IAFs = ($ImportFlowCSV | where {$_.ma -eq $MAid})
        if ($IAFs)
        {
            if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

            $table = StartTable -selection $doc -Headings @("CS Object Type","CS Attributes","Mapping Type","MV Object Type","MV Attribute","Advanced Rule")
            $r = 1

            foreach ($IAF in $IAFs | sort "csobject","csattr")
            {
                          
                if ($IAF.flowtype -like "SyncRule*" -and $SRNames)
                {
                    $SRid = $IAF.flowtype.Replace("SyncRule ","")
                    $FlowType = $IAF.flowtype.Replace($SRid,($SRNames | where {$SRid.Contains($_."object_id")}).displayName)
                }
                else {$FlowType = $IAF.flowtype}

                if ($IAF.flowtype -eq "Constant") {$CSAttr = $IAF.value}
                else {$CSAttr = $IAF.csattr}

                if ($IAF.value -ne $CSAttr) {$Value = $IAF.value}
                else {$Value = $null}

                ## Add row
                $r += 1
                AddTableRow -table $table -row $r -ColumnText @($IAF.csobject,$CSAttr,$Flowtype,$IAF.mvobject,$IAF.mvattr,$Value)

                ## UNIFY Only - Add Advanced flow rule details from CF config file
                if ($FlowType -eq "Advanced" -and $CFCSV)
                {
                    $CFRule = $CFCSV | where {$_.ma -eq $MAName -and $_.rulename -eq $IAF.value}
                    if ($CFRule)
                    {
                        $subtable = StartTable -selection $table.cell($r,6) -TableStyle $SubTableStyle -Columns 2 -Subtable $true
                        $sr = 0
                        $NotFirstRow = $false
                        foreach ($rule in $CFRule)
                        {
                            $sr += 1
                            $ruleconfig = ""
                            if ($rule.flowtype) {$ruleconfig = $rule.flowtype + "; "}
                            if ($rule.parameters) {$ruleconfig = $ruleconfig + $rule.parameters + "; "}
                            if ($rule.value) {$ruleconfig = $ruleconfig + $rule.value + "; "}
                            if ($rule.filter) {$ruleconfig = $ruleconfig + $rule.filter}
                            AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($rule.priority,$ruleconfig)
                            $NotFirstRow = $true
                        }
                    }
                }
            }
            # Move to the end of the Table
            $doc.EndOf(15)
            $doc.EndOf(6)
            $doc.MoveDown()
        }
        else
        {
            if ($DocText.($DocItem).None) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}
        }


        ###
        ### Export Attribute Flows
        ###
        $DocItem = "Sync-EAF"
        WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading

        # Write EAF Table
        $EAFs = ($ExportFlowCSV | where {$_.ma -eq $MAid})
        if ($EAFs)
        {
            if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

            $table = StartTable -selection $doc -Headings @("MV Object Type","MV Attributes","Mapping Type","CS Object Type","CS Attribute","Advanced Rule")
            $r = 1

            foreach ($EAF in $EAFs | sort "mvobject","mvattr")
            {

                if ($EAF.flowtype -like "SyncRule*" -and $SRNames)
                {
                    $SRid = $EAF.flowtype.Replace("SyncRule ","")
                    $FlowType = $EAF.flowtype.Replace($SRid,($SRNames | where {$SRid.Contains($_."object_id")}).displayName)
                }
                else {$FlowType = $EAF.flowtype}

                if ($EAF.flowtype -eq "Constant") {$CSAttr = $EAF.value}
                else {$CSAttr = $EAF.csattr}

                if ($FlowType -eq "Advanced" -and $CFCSV) {$Value = $null}
                else {$Value = $EAF.value}
                
 
                ## Add row
                $r += 1
                AddTableRow -table $table -row $r -ColumnText @($EAF.mvobject,$EAF.mvattr,$Flowtype,$EAF.csobject,$CSAttr,$Value)


                ## UNIFY Only - Replace "Source Value" column with Advanced flow rule details from CF config file
                if ($FlowType -eq "Advanced" -and $CFCSV)
                {
                    $CFRule = $CFCSV | where {$_.ma -eq $MAName -and $_.rulename -eq $EAF.value}
                    if ($CFRule)
                    {
                        $subtable = StartTable -selection $table.cell($r,6) -TableStyle $SubTableStyle -Columns 2 -Subtable $true
                        $sr = 0
                        $NotFirstRow = $false
                        foreach ($rule in $CFRule)
                        {
                            $sr += 1
                            $ruleconfig = ""
                            if ($rule.flowtype) {$ruleconfig = $rule.flowtype + "; "}
                            if ($rule.parameters) {$ruleconfig = $ruleconfig + $rule.parameters + "; "}
                            if ($rule.value) {$ruleconfig = $ruleconfig + $rule.value + "; "}
                            if ($rule.filter) {$ruleconfig = $ruleconfig + $rule.filter}
                            AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($rule.priority,$ruleconfig)
                            $NotFirstRow = $true
                        }
                    }
                }
            }
            # Move to the end of the Table
            $doc.EndOf(15)
            $doc.EndOf(6)
            $doc.MoveDown()
        }
        else
        {
            if ($DocText.($DocItem).None) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}
        } 
    }
}}


###
### Event Broker Operations
###

if ($EBExtensibilityFolder -and -not $SingleMA -and -not $FlowMapOnly)
{
    WriteLine -selection $doc -Style $Heading1 -Text "Event Broker Operations"
    WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text "Event Broker manages the running of tasks relating to synchronisation in the solution."
    
    $table = StartTable -selection $doc -Headings @("Operation","StartUp","Schedules","Steps")
    $r = 1

    foreach ($OpName in $hashEB.Keys | sort)
    {
       $ColumnText = @()
       $ColumnText += $OpName

       $StartUp = "Enabled: " + $hashEB.($OpName).General.Enabled + "`nRun On Startup: " + $hashEB.($OpName).General."Run On Startup" + "`nQueue Missed: " + $hashEB.($OpName).General."Queue Missed"
       $ColumnText += $StartUp

       $Schedules = @()
       foreach ($step in $hashEB.($OpName)."Schedules".Keys | sort)
       { 
            foreach ($item in $hashEB.($OpName)."Schedules".($step).Keys | sort) 
            { 
                $Schedules += ($item + ": " + $hashEB.($OpName)."Schedules".($step).($item)) 
            }
       }
       $ColumnText += $Schedules -join "`n"


       $r += 1
       AddTableRow -table $table -row $r -ColumnText $ColumnText

       #Create subtable for the steps
       $subtable = StartTable -selection $table.cell($r,4)  -Columns 2  -Subtable $true -TableStyle $SubTableStyle
       $sr = 0

       foreach ($step in $hashEB.($OpName)."Steps".Keys | sort)
       {
            $sr += 1
            if ($sr -eq 1) {$NotFirstRow = $false} else {$NotFirstRow = $true} 
            if ($hashEB.($OpName)."Steps".($step).Name) 
            {
                $text = $hashEB.($OpName)."Steps".($step).Name
            }
            else 
            {
                if ($hashEB.($OpName)."Steps".($step).MA)
                {
                    $text = $hashEB.($OpName)."Steps".($step).MA + ", " + $hashEB.($OpName)."Steps".($step)."Run Profile"
                }
                else
                {
                    $text = $hashEB.($OpName)."Steps".($step).Type
                }
            }
            AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($step,$text)
       }


       # Move to the end of the Table
       $doc.EndOf(15)
       $doc.EndOf(6)
       $doc.MoveDown()

   }

}


###
### Save and Display Word Document if running invisible
###

if (-not $Visible) 
{
    $Error.clear()
    Try
    {
        $doc.Application.ActiveWindow.Visible = $true
        [ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
        $SaveFilePath = $env:temp + "\SyncConfig.docx"
        $doc.Document.SaveAs([ref]$SaveFilePath,[ref]$saveFormat::wdFormatDocumentDefault)

        write-host "Document saved as $SaveFilePath"
    }
    Catch {write-error $Error[0]}
}
