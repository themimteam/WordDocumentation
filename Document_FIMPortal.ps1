PARAM ( [string]$SchemaFile,
        [string]$PolicyFile,
        [string]$StylesScript,
        [string]$CustomisationsScript,
        [boolean]$IncludeSchema = $true,
        [boolean]$IncludeMPRs = $true,
        [boolean]$IncludeWFs = $true,
        [boolean]$IncludeSets = $true,
        [boolean]$IncludeEmailTemplates = $true,
        [boolean]$IncludeUI = $true,
        [boolean]$IncludeOtherObjects = $false,
        [boolean]$Visible=$true,
        [string]$ErrorAction="Continue"
      )

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
###  Document_FIMPortal.ps1
###
###  Written by Carol Wapshere
###
###  Takes the policy.xml and schema.xml export files and produces a Word document covering Schema, Policy, Portal UI and optionally other object types.
###  Use the Include* switches if you would rather produce seperate documents.
###
###  Tested with FIM 2010 R2 and Word 2010/2013.
###
###  Parameters:
###    -SchemaFile             (Required) The full path to the XML file exported by ExportSchema.ps1
###    -PolicyFile             (Required) The full path to the XML file exported by ExportPolicy.ps1
###    -StylesScript           (Optional) The full path to a customised version of Include_STYLES.ps1. Uses the version in the script folder if not specified. Will fail if not found.
###    -IncludeSchema          (Default True) Include the Schema section in the document
###    -IncludeMPRs            (Default True) Include the MPR section in the document
###    -IncludeWFs             (Default True) Include the Workflows section in the document
###    -IncludeSets            (Default True) Include the Sets section in the document
###    -IncludeEmailTemplates  (Default True) Include the Email Templates section in the document
###    -IncludeUI              (Default True) Include the UI section in the document
###    -IncludeOtherObjects    (Default False) Include the other objects (eg., Groups, Custom Objects) in the document
###                                o To include other objects you must modify ExportPolicy.xml to include them in the policy export file,
###                                  and modify the $ReportObjects hashtable in Include_CustomisedContent.ps1 to include the required object types and attributes.
###    -Visible                (Default=True) If set to $false does not open Word until the document is finished, which runs faster than with Visible=$true.
###    -ErrorAction            (Default=Continue) Will continue past errors. Set to "Stop" if you want it to stop on any error.
###
###  BEFORE RUNNING THIS SCRIPT:
###
###    1. Export the schema and policy using the standard configuration migration scripts http://technet.microsoft.com/en-us/library/ff400275(v=ws.10).aspx
###        - Note if you want to include Groups or custom object types in the document you must modify the ExportPolicy.ps1 script, eg: 
###             $policy = Export-FIMConfig -policyConfig -portalConfig -customConfig ("/Group", "/Role") -MessageSize 9999999
###          You must also modify the $ReportObjects hashtable in the Include_CustomisedContent.ps1 script.
###
###    2. Modify the Include_STYLES.ps1 script to specify a different Word template and/or styles.
###
###    3. Modify the Include_CustomisedContent.ps1 script to change document wording, which objects and parameters are included, and any custom object types.
###
###    4. Run the script specifying the full paths to the Policy and Schema exports files, and any other optional parameters (see above).
###
###
###  CHANGES:
###    CW 25/02/2014 - Fixed a problem where an empty $ShowObjectTypes showed no object types instead of all.
###    CW 11/08/2014 - New switches Visible and ErrorAction
###                  - Fixed a bug with the use of the $SubSections hashtable from Include_CustomisedContent.ps1
###    BB 21/11/2014 - Fixed nested tables for RCDC section
###    CW 4/12/2014  - Fixed a bug that stopped MPRs being listed in the Workflow table
###    CW 25/1/2016  - Fixed a bug where MPRs were not listed if Disabled is null. Added CustomisationsScript parameter.
###


### Run Include scripts. Fail if any are not found.
$ErrorActionPreference = "Stop"

$ScriptFldr = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

if ($StylesScript) {. $StylesScript} else {. $ScriptFldr\Include_STYLES.ps1}
. $ScriptFldr\Include_WordFunctions.ps1
if ($CustomisationsScript) {. $CustomisationsScript} else {. $ScriptFldr\Include_CustomisedContent.ps1}

$ErrorActionPreference = $ErrorAction


###
### Analyse and collect info from the config file and store in hash tables
###


## Open the config file
[xml]$Policy = ReadFileContents -FilePath $PolicyFile
[xml]$Schema = ReadFileContents -FilePath $SchemaFile

## Hash Tables applying to all object types
$ObjectName = @{}
$ObjectDescription = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject"))
{
    $ObjectName.Add($ObjectNode.ResourceManagementObject.ObjectIdentifier,($ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'DisplayName'}).Value)
    $ObjectDescription.Add($ObjectNode.ResourceManagementObject.ObjectIdentifier,($ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'Description'}).Value)
}

## Hashtable of Attributes
$hashAttributes = @{}
foreach ($object in $schema.Results.ExportObject)
{
    if ($object.ResourceManagementObject.ObjectType -eq 'AttributeTypeDescription')
    {   
        $Name =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'Name'}).Value
        $DisplayName =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'DisplayName'}).Value
        $ObjectID =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'ObjectID'}).Value
        $DataType =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'DataType'}).Value
        $Description =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'Description'}).Value
        $Multivalued =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'Multivalued'}).Value
        $StringRegex =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'StringRegex'}).Value
        
        $hashAttributes.Add($ObjectID,@{})
        $hashAttributes.($ObjectID).Add("Name",$Name)
        $hashAttributes.($ObjectID).Add("DataType",$DataType)
        $hashAttributes.($ObjectID).Add("Description",$Description)
        $hashAttributes.($ObjectID).Add("Multivalued",$Multivalued)
        $hashAttributes.($ObjectID).Add("StringRegex",$StringRegex)
    }
}

## Hashtable of Object Types
$hashObjectTypes = @{}
foreach ($object in $schema.Results.ExportObject)
{
    if ($object.ResourceManagementObject.ObjectType -eq 'ObjectTypeDescription')
    {   
        $Name =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'Name'}).Value
        $DisplayName =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'DisplayName'}).Value
        $ObjectID =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'ObjectID'}).Value
        
        $hashObjectTypes.Add($ObjectID,@{})
        $hashObjectTypes.($ObjectID).Add("Name",$Name)
        $hashObjectTypes.($ObjectID).Add("DisplayName",$DisplayName)
    }
}

## Hashtable of Bindings - with extra info about the bound Attribute
$hashBindings = @{}
foreach ($object in $schema.Results.ExportObject)
{
    if ($object.ResourceManagementObject.ObjectType -eq 'BindingDescription')
    {   
        $DisplayName =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'DisplayName'}).Value
        $ObjectID =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'ObjectID'}).Value
        $BoundAttributeTypeID =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'BoundAttributeType'}).Value
        $BoundAttributeType = $hashAttributes.($BoundAttributeTypeID).Name
        $BoundObjectTypeID =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'BoundObjectType'}).Value
        $BoundObjectType = $hashObjectTypes.($BoundObjectTypeID).Name
        $Required =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'Required'}).Value

        ## Take the following from the Attribute if not set on the Binding
        $Description = $null
        $Description =  ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'Description'}).Value
        if (-not $Description) {$Description = $hashAttributes.($BoundAttributeTypeID).Description}
        $StringRegex = $null
        $StringRegex = ($object.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute | where {$_.AttributeName -eq 'StringRegex'}).Value
        if (-not $StringRegex) {$StringRegex = $hashAttributes.($BoundAttributeTypeID).StringRegex}
        
        if (-not $hashBindings.ContainsKey($BoundObjectType)) {$hashBindings.Add($BoundObjectType,@{})}
        $hashBindings.($BoundObjectType).Add($BoundAttributeType,@{})
        $hashBindings.($BoundObjectType).($BoundAttributeType).Add("DisplayName",$DisplayName)
        $hashBindings.($BoundObjectType).($BoundAttributeType).Add("Required",$Required)
        $hashBindings.($BoundObjectType).($BoundAttributeType).Add("StringRegex",$StringRegex)
        $hashBindings.($BoundObjectType).($BoundAttributeType).Add("DataType",$hashAttributes.($BoundAttributeTypeID).DataType)
        $hashBindings.($BoundObjectType).($BoundAttributeType).Add("Multivalued",$hashAttributes.($BoundAttributeTypeID).Multivalued)
        $hashBindings.($BoundObjectType).($BoundAttributeType).Add("Description",$Description)
    }
}

## Hash Tables of Set details
$hashSet = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject") | where {$_.ResourceManagementObject.ObjectType -eq 'Set'} )
{
    $Attributes = $ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute
    $SetName = ($Attributes | where {$_.AttributeName -eq 'DisplayName'}).Value
    $ObjectID =  ($Attributes | where {$_.AttributeName -eq 'ObjectID'}).Value

    $filter = ($Attributes | where {$_.AttributeName -eq 'Filter'}).Value  
    $filter = ReplaceGUIDs -text $filter -hashIDtoName $ObjectName 
    $filter = ($filter -Replace $FilterPrefix,"") -Replace $FilterEnd,""

    $members = @()
    foreach ($guid in ($Attributes | where {$_.AttributeName -eq 'ExplicitMember'}).Values.string)
    {
        $members += ReplaceGUIDs -Text $guid -hashIDtoName $ObjectName
    }

    $hashSet.Add($SetName,@{})
    $hashSet.($SetName).Add("ExplicitMembers",$members)
    $hashSet.($SetName).Add("Description",($Attributes | where {$_.AttributeName -eq 'Description'}).Values)
    $hashSet.($SetName).Add("Filter",$filter)
    #$hashSet.($SetName).Add("Bookmark","set" + $ObjectID.Replace("urn:uuid:","").Replace("-",""))
}


## Hash Tables of Email Templates
$hashET = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject") | where {$_.ResourceManagementObject.ObjectType -eq 'EmailTemplate'} )
{    
    $Attributes = $ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute
    $ETName = ($Attributes | where {$_.AttributeName -eq 'DisplayName'}).Value

    $hashET.Add($ETName,@{})
    $hashET.($ETName).Add("Subject",($Attributes | where {$_.AttributeName -eq 'EmailSubject'}).Value)
    $hashET.($ETName).Add("TemplateType",($Attributes | where {$_.AttributeName -eq 'EmailTemplateType'}).Value)
    $hashET.($ETName).Add("Body",($Attributes | where {$_.AttributeName -eq 'EmailBody'}).Value)
}

## Hash Tables of Workflow details
$hashWF = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject") | where {$_.ResourceManagementObject.ObjectType -eq 'WorkflowDefinition'} )
{
    $Attributes = $ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute
    $WFName = ($Attributes | where {$_.AttributeName -eq 'DisplayName'}).Value

    $hashWF.Add($WFName,@{})
    $hashWF.($WFName).Add("RequestPhase",($Attributes | where {$_.AttributeName -eq 'RequestPhase'}).Value)
    $hashWF.($WFName).Add("Description",($Attributes | where {$_.AttributeName -eq 'Description'}).Value)

    # Workflow Steps
    [xml]$XOML = ($Attributes | where {$_.AttributeName -eq 'XOML'}).Value
    $stepCount = 1
    $hashWF.($WFName).Add("Steps",@{})

    foreach ($wfStepNode in $XOML.SequentialWorkflow.ChildNodes)
    {    
        $activityType = $wfStepNode.LocalName
        $stepName = $stepCount.ToString() + ". " + $activityType
        $hashWF.($WFName).("Steps").Add($stepName,@())

        if ($activityType -and $activitySteps.ContainsKey($activityType))
        {
            foreach ($Attribute in $activitySteps.($activityType))
            {
                $text = ReplaceGUIDs -Text $wfStepNode.($Attribute) -hashIDtoName $ObjectName
                $hashWF.($WFName).("Steps").($stepName) += $Attribute + ": " + $text
            }
        }
        else
        {
            foreach ($text in $wfStepNode.Attributes."#text")
            {
                $text = ReplaceGUIDs -Text $text -hashIDtoName $ObjectName
                $hashWF.($WFName).("Steps").($stepName) += $text
            }
        }
        $stepCount += 1
    }
}


## Hash Tables of MPRs

$hashMPRGRSet = @{}
$hashMPRGRRel = @{}
$hashMPRWFReq = @{}
$hashMPRWFTrans = @{}
$hashMPRWFs = @{}
$hashMPRSets = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject") | where {$_.ResourceManagementObject.ObjectType -eq 'ManagementPolicyRule'} )
{    
    $Attributes = $ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute
    $MPRName = ($Attributes | where {$_.AttributeName -eq 'DisplayName'}).Value

    $Disabled = ($Attributes | where {$_.AttributeName -eq 'Disabled'}).Value

    if (-not $Disabled -or $Disabled -eq 'False')
    {
        $MPRType = $null
        $MPRType = ($Attributes | where {$_.AttributeName -eq 'ManagementPolicyRuleType'}).Value

        $Description = $null
        $Description = ($Attributes | where {$_.AttributeName -eq 'Description'}).Value

        $GrantRight = $null
        $GrantRight = ($Attributes | where {$_.AttributeName -eq 'GrantRight'}).Value

        $ActionParameters = $null
        $ActionParameters = (($Attributes | where {$_.AttributeName -eq 'ActionParameter'}).Values).string

        $ActionType = $null
        $ActionType = (($Attributes | where {$_.AttributeName -eq 'ActionType'}).Values).string

        $PrincipalRelativeToResource = $null
        $PrincipalRelativeToResource = ($Attributes | where {$_.AttributeName -eq 'PrincipalRelativeToResource'}).Value

        $PrincipalSet = $null
        $PrincipalSet = ($Attributes | where {$_.AttributeName -eq 'PrincipalSet'}).Value
        if ($PrincipalSet) 
        {
            $PrincipalSet = $ObjectName.($PrincipalSet)
            if (-not $hashMPRSets.ContainsKey($PrincipalSet)) {$hashMPRSets.Add($PrincipalSet,@())}
            $hashMPRSets.($PrincipalSet) += $MPRName
        }

        $ResourceCurrentSet = $null
        $ResourceCurrentSet = ($Attributes | where {$_.AttributeName -eq 'ResourceCurrentSet'}).Value
        if ($ResourceCurrentSet)
        {
            $ResourceCurrentSet = $ObjectName.($ResourceCurrentSet)
            if (-not $hashMPRSets.ContainsKey($ResourceCurrentSet)) {$hashMPRSets.Add($ResourceCurrentSet,@())}
            $hashMPRSets.($ResourceCurrentSet) += $MPRName
        }

        $ResourceFinalSet = $null
        $ResourceFinalSet = ($Attributes | where {$_.AttributeName -eq 'ResourceFinalSet'}).Value
        if($ResourceFinalSet)
        {
            $ResourceFinalSet = $ObjectName.($ResourceFinalSet)
            if (-not $hashMPRSets.ContainsKey($ResourceFinalSet)) {$hashMPRSets.Add($ResourceFinalSet,@())}
            $hashMPRSets.($ResourceFinalSet) += $MPRName
        }

        $WFs = $null
        $WFs = (($Attributes | where {$_.AttributeName -eq 'AuthorizationWorkflowDefinition'}).Values).string
        $WFAuthZs = @()
        if ($WFs) {foreach ($WF in $WFs)
        {
            $WFName = $ObjectName.($WF)
            $WFAuthZs = $WFName
            if (-not $hashMPRWFs.ContainsKey($WFName)) {$hashMPRWFs.Add($WFName,@())}
            $hashMPRWFs.($WFName) += $MPRName
        }}        

        $WFs = $null
        $WFs = (($Attributes | where {$_.AttributeName -eq 'ActionWorkflowDefinition'}).Values).string
        $WFActions = @()
        if ($WFs) {foreach ($WF in $WFs)
        {
            $WFName = $ObjectName.($WF)
            $WFActions = $ObjectName.($WF)
            if (-not $hashMPRWFs.ContainsKey($WFName)) {$hashMPRWFs.Add($WFName,@())}
            $hashMPRWFs.($WFName) += $MPRName
        }}

        if ($GrantRight -eq "True" -and $PrincipalSet -ne $null)
        {
            $SetName = $PrincipalSet
            if (-not $hashMPRGRSet.ContainsKey($SetName)) {$hashMPRGRSet.Add($SetName,@{})}
            $hashMPRGRSet.($SetName).Add($MPRName,@{})
            $hashMPRGRSet.($SetName).($MPRName).Add("Description",$Description)
            $hashMPRGRSet.($SetName).($MPRName).Add("ActionParameters",$ActionParameters)
            $hashMPRGRSet.($SetName).($MPRName).Add("PrincipalSet",$PrincipalSet)
            $hashMPRGRSet.($SetName).($MPRName).Add("ResourceCurrentSet",$ResourceCurrentSet)
            $hashMPRGRSet.($SetName).($MPRName).Add("ResourceFinalSet",$ResourceFinalSet)
            $hashMPRGRSet.($SetName).($MPRName).Add("ActionType",$ActionType)
        }

        if ($GrantRight -eq "True" -and $PrincipalRelativeToResource -ne $null)
        {
            if (-not $hashMPRGRRel.ContainsKey($PrincipalRelativeToResource)) {$hashMPRGRRel.Add($PrincipalRelativeToResource,@{})}
            $hashMPRGRRel.($PrincipalRelativeToResource).Add($MPRName,@{})
            $hashMPRGRRel.($PrincipalRelativeToResource).($MPRName).Add("Description",$Description)
            $hashMPRGRRel.($PrincipalRelativeToResource).($MPRName).Add("ActionParameters",$ActionParameters)
            $hashMPRGRRel.($PrincipalRelativeToResource).($MPRName).Add("ResourceCurrentSet",$ResourceCurrentSet)
            $hashMPRGRRel.($PrincipalRelativeToResource).($MPRName).Add("ResourceFinalSet",$ResourceFinalSet)
            $hashMPRGRRel.($PrincipalRelativeToResource).($MPRName).Add("ActionType",$ActionType)
            $hashMPRGRRel.($PrincipalRelativeToResource).($MPRName).Add("PrincipalRelativeToResource",$PrincipalRelativeToResource)
        }

        if (($WFAuthZs -ne $null -or $WFActions -ne $null) -and $MPRType -eq "Request")
        {
            foreach ($Attrib in $ActionParameters)
            {
                if (-not $hashMPRWFReq.ContainsKey($Attrib)) {$hashMPRWFReq.Add($Attrib,@{})}
                $hashMPRWFReq.($Attrib).Add($MPRName,@{})
                $hashMPRWFReq.($Attrib).($MPRName).Add("Description",$Description)
                $hashMPRWFReq.($Attrib).($MPRName).Add("WFAuthZs",$WFAuthZs)
                $hashMPRWFReq.($Attrib).($MPRName).Add("WFActions",$WFActions)
                $hashMPRWFReq.($Attrib).($MPRName).Add("ActionParameters",$ActionParameters)
                $hashMPRWFReq.($Attrib).($MPRName).Add("PrincipalSet",$PrincipalSet)
                $hashMPRWFReq.($Attrib).($MPRName).Add("ResourceCurrentSet",$ResourceCurrentSet)
                $hashMPRWFReq.($Attrib).($MPRName).Add("ResourceFinalSet",$ResourceFinalSet)
                $hashMPRWFReq.($Attrib).($MPRName).Add("ActionType",$ActionType)
                $hashMPRWFReq.($Attrib).($MPRName).Add("PrincipalRelativeToResource",$PrincipalRelativeToResource)
            }
        }        

        if (($WFAuthZs -ne $null -or $WFActions -ne $null) -and $MPRType -eq "SetTransition")
        {
            if ($ActionType -eq "TransitionIn") {$SetName = $ResourceFinalSet}
            else {$SetName = $ResourceCurrentSet}
            if (-not $hashMPRWFTrans.ContainsKey($SetName)) {$hashMPRWFTrans.Add($SetName,@{})}
            $hashMPRWFTrans.($SetName).Add($MPRName,@{})
            $hashMPRWFTrans.($SetName).($MPRName).Add("Description",$Description)
            $hashMPRWFTrans.($SetName).($MPRName).Add("WFAuthZs",$WFAuthZs)
            $hashMPRWFTrans.($SetName).($MPRName).Add("WFActions",$WFActions)
            $hashMPRWFTrans.($SetName).($MPRName).Add("ActionParameters",$ActionParameters)
            $hashMPRWFTrans.($SetName).($MPRName).Add("ResourceCurrentSet",$ResourceCurrentSet)
            $hashMPRWFTrans.($SetName).($MPRName).Add("ResourceFinalSet",$ResourceFinalSet)
            $hashMPRWFTrans.($SetName).($MPRName).Add("ActionType",$ActionType)
        }
    }
} 

## Hash Table of RCDC details
$hashRCDC = @{}
$hashRCDCStrings = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject") | where {$_.ResourceManagementObject.ObjectType -eq 'ObjectVisualizationConfiguration'} )
{
    $Attributes = $ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute

    $ObjectType = ($Attributes | where {$_.AttributeName -eq 'TargetObjectType'}).Value
    $DisplayName = ($Attributes | where {$_.AttributeName -eq 'DisplayName'}).Value
    
    # Remove following condition to show all RCDCs
    if ($ShowRCDCTypes -contains $ObjectType)
    {
        if (-not $hashRCDC.ContainsKey($ObjectType)) {$hashRCDC.Add($ObjectType,@{})}
        $hashRCDC.($ObjectType).Add($DisplayName,@{})
        $hashRCDC.($ObjectType).($DisplayName).Add("Description",($Attributes | where {$_.AttributeName -eq 'Description'}).Value)

        #Applies To Create, Edit or View
        $Actions = @()
        if (($Attributes | where {$_.AttributeName -eq 'AppliesToCreate'}).Value -eq "True") {$Actions += "Create"}
        if (($Attributes | where {$_.AttributeName -eq 'AppliesToEdit'}).Value -eq "True") {$Actions += "Edit"}
        if (($Attributes | where {$_.AttributeName -eq 'AppliesToView'}).Value -eq "True") {$Actions += "View"}
        $hashRCDC.($ObjectType).($DisplayName).Add("Actions",$Actions -join ",")
        
        #Target Object Type
        $hashRCDC.($ObjectType).($DisplayName).Add("TargetObjectType",$ObjectType)

        #String Resources
        if (($Attributes | where {$_.AttributeName -eq 'StringResources'}).Value)
        {
            $hashRCDCStrings.Add($DisplayName,@{})
            [xml]$StringResources = ($Attributes | where {$_.AttributeName -eq 'StringResources'}).Value
            foreach ($item in $StringResources.SymbolResourcePairs.SymbolResourcePair)
            {
                if ($item.Symbol -notmatch $RCDCStringsRegexNotMatch) {$hashRCDCStrings.($DisplayName).Add($item.Symbol,$item.ResourceString)}
            }
        }

        #Configuration XML
        [xml]$Config = ($Attributes | where {$_.AttributeName -eq 'ConfigurationData'}).Value
        $g = 0
        $hashRCDC.($ObjectType).($DisplayName).Add("Tabs",@{})
        
        foreach ($group in $Config.ObjectControlConfiguration.Panel.Grouping | where {$RCDCTabsHide -notcontains $_.Name} )
        {
            $g += 1
            $GrpID = $g.ToString("00") + " " + $group.Name
            $hashRCDC.($ObjectType).($DisplayName).Tabs.Add($GrpID,@{})

            $c = 0
            foreach ($control in $group.Control)
            {
                $c += 1
                $CtrlID = $c.ToString("00") + " " + $control.Name
                $hashRCDC.($ObjectType).($DisplayName).Tabs.($GrpID).Add($CtrlID,@())
                $hashRCDC.($ObjectType).($DisplayName).Tabs.($GrpID).($CtrlID) += "_Type = " + $control.TypeName
                foreach ($property in $control.Properties.Property) 
                {
                    if ($ReportProps.ContainsKey($control.TypeName))
                    {
                         if ($ReportProps.($control.TypeName) -contains $property.Name)
                         {$hashRCDC.($ObjectType).($DisplayName).Tabs.($GrpID).($CtrlID) += $property.Name + " = " + $property.Value}
                    }
                    else
                    {
                        $hashRCDC.($ObjectType).($DisplayName).Tabs.($GrpID).($CtrlID) += $property.Name + " = " + $property.Value
                    }
                }
            }
        }
    }
}


## Hash Tables of Navigation Bar Resource details
$hashNBR = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject") | where {$_.ResourceManagementObject.ObjectType -eq 'NavigationBarConfiguration'} )
{
    $Attributes = $ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute
    $DisplayName = ($Attributes | where {$_.AttributeName -eq 'DisplayName'}).Value
    
    # Store in hastable by ParentOrder
    [int]$ParentOrder = ($Attributes | where {$_.AttributeName -eq 'ParentOrder'}).Value   
    if (-not $hashNBR.ContainsKey($ParentOrder)) {$hashNBR.Add($ParentOrder,@{})}
     
    # Order - duplicates allowed so add suffixes until it's unique
    [int]$Order = ($Attributes | where {$_.AttributeName -eq 'Order'}).Value
    if (-not $hashNBR.($ParentOrder).ContainsKey($Order)) 
    {
        $OrderId = $Order
        $hashNBR.($ParentOrder).Add($OrderId,@{})
    }
    else
    {
        $i = 1
        $Added = $false
        while (-not $Added)
        {
            $OrderId = $Order + $i/100
            if (-not $hashNBR.($ParentOrder).ContainsKey($OrderId)) 
            {
                $hashNBR.($ParentOrder).Add($OrderId,@{})
                $Added = $true
            }
            $i += 1
        }
    }
    
    # DisplayName
    $hashNBR.($ParentOrder).($OrderId).Add($DisplayName,@{})

    # Other properties
    $hashNBR.($ParentOrder).($OrderId).($DisplayName).Add("Description",($Attributes | where {$_.AttributeName -eq 'Description'}).Value)
    $hashNBR.($ParentOrder).($OrderId).($DisplayName).Add("NavigationUrl",($Attributes | where {$_.AttributeName -eq 'NavigationUrl'}).Value)
    $hashNBR.($ParentOrder).($OrderId).($DisplayName).Add("UsageKeyword",($Attributes | where {$_.AttributeName -eq 'UsageKeyword'}).Values.string -join ", ")
}


## Hash Tables of Home Page Resource details
$hashHPR = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject") | where {$_.ResourceManagementObject.ObjectType -eq 'HomepageConfiguration'} )
{
    $Attributes = $ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute
    $DisplayName = ($Attributes | where {$_.AttributeName -eq 'DisplayName'}).Value
    $Region = ($Attributes | where {$_.AttributeName -eq 'Region'}).Value
    
    if ($Region -eq "1")
    {
        # Store in hastable by ParentOrder
        #[int]$ParentOrder = ($Attributes | where {$_.AttributeName -eq 'Region'}).Value + "." + ($Attributes | where {$_.AttributeName -eq 'ParentOrder'}).Value
        [int]$ParentOrder = ($Attributes | where {$_.AttributeName -eq 'ParentOrder'}).Value
        if (-not $hashHPR.ContainsKey($ParentOrder)) {$hashHPR.Add($ParentOrder,@{})}
     
        # Order - duplicates allowed so add suffixes until it's unique
        [int]$Order = ($Attributes | where {$_.AttributeName -eq 'Order'}).Value
        if (-not $hashHPR.($ParentOrder).ContainsKey($Order)) 
        {
            $OrderId = $Order
            $hashHPR.($ParentOrder).Add($OrderId,@{})
        }
        else
        {
            $i = 1
            $Added = $false
            while (-not $Added)
            {
                $OrderId = $Order + $i/100
                if (-not $hashHPR.($ParentOrder).ContainsKey($OrderId)) 
                {
                    $hashHPR.($ParentOrder).Add($OrderId,@{})
                    $Added = $true
                }
                $i += 1
            }
        }

        # DisplayName
        $hashHPR.($ParentOrder).($OrderId).Add($DisplayName,@{})

        # Other properties
        $hashHPR.($ParentOrder).($OrderId).($DisplayName).Add("Description",($Attributes | where {$_.AttributeName -eq 'Description'}).Value)
        $hashHPR.($ParentOrder).($OrderId).($DisplayName).Add("NavigationUrl",($Attributes | where {$_.AttributeName -eq 'NavigationUrl'}).Value)
        $hashHPR.($ParentOrder).($OrderId).($DisplayName).Add("UsageKeyword",($Attributes | where {$_.AttributeName -eq 'UsageKeyword'}).Values.string -join ", ")
    }
}


## Hash Tables of Search Scopes
$hashSS = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject") | where {$_.ResourceManagementObject.ObjectType -eq 'SearchScopeConfiguration'} )
{
    $Attributes = $ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute
    $DisplayName = ($Attributes | where {$_.AttributeName -eq 'DisplayName'}).Value

    $UsageKeywords = ($Attributes | where {$_.AttributeName -eq 'UsageKeyword'}).Values

    $Include = $false
    if ($SSUsageKeyword -and $SSUsageKeyword.count -gt 0)
    {
        foreach ($UsageKeyword in $SSUsageKeyword) {if ($UsageKeywords.string -contains $UsageKeyword) {$Include = $true}}
    }
    else {$Include = $true}

    ## Show Search Scopes base on Usage Keyword if supplied
    if ($Include)
    {
        $TargetObjectType = ($Attributes | where {$_.AttributeName -eq 'SearchScopeResultObjectType'}).Value
        if (-not $hashSS.ContainsKey($TargetObjectType)) {$hashSS.Add($TargetObjectType,@{})}
        
        if (-not $hashSS.($TargetObjectType).ContainsKey($DisplayName))
        {
            $hashSS.($TargetObjectType).Add($DisplayName,@{})
            $hashSS.($TargetObjectType).($DisplayName).Add("Description",($Attributes | where {$_.AttributeName -eq 'Description'}).Value)
            $hashSS.($TargetObjectType).($DisplayName).Add("Filter",($Attributes | where {$_.AttributeName -eq 'SearchScope'}).Value)

            $Columns = ($Attributes | where {$_.AttributeName -eq 'SearchScopeColumn'}).Value
            if ($Columns) {$hashSS.($TargetObjectType).($DisplayName).Add("Columns",$Columns.split(";"))}
            else {$hashSS.($TargetObjectType).($DisplayName).Add("Columns",$null)}
        
            $hashSS.($TargetObjectType).($DisplayName).Add("Order",($Attributes | where {$_.AttributeName -eq 'Order'}).Value)
            $hashSS.($TargetObjectType).($DisplayName).Add("TargetObjectType",($Attributes | where {$_.AttributeName -eq 'SearchScopeResultObjectType'}).Value)
            $hashSS.($TargetObjectType).($DisplayName).Add("UsageKeyword",$UsageKeywords.string -join ", ")
            $hashSS.($TargetObjectType).($DisplayName).Add("AdvancedFilter",($Attributes | where {$_.AttributeName -eq 'msidmSearchScopeAdvancedFilter'}).Value)
        }
    }
}

## Hash Tables of Other Objects
$hashObjType = @{}
$hashObjProperties = @{}

foreach ($ObjectNode in $Policy.SelectNodes("Results/ExportObject"))
{
    $Attributes = $ObjectNode.ResourceManagementObject.ResourceManagementAttributes.ResourceManagementAttribute
    $ObjID = $ObjectNode.ResourceManagementObject.ObjectIdentifier
    $ObjectType = ($Attributes | where {$_.AttributeName -eq 'ObjectType'}).Value

    if ($ReportObjects.ContainsKey($ObjectType))
    {
        $hashObjType.Add($ObjID,$ObjectType)
        if (-not $hashObjProperties.ContainsKey($ObjectType)) {$hashObjProperties.Add($ObjectType,@{})}

        # Nest hashtable down a further level based on $SubSection if configured for this ObjectType
        if ($SubSections.ContainsKey($ObjectType))
        {
            $Property = ($Attributes | where {$_.AttributeName -eq $SubSections.($ObjectType)}).value
            if (-not $hashObjProperties.($ObjectType).ContainsKey($Property)){$hashObjProperties.($ObjectType).Add($Property,@{})}
            $StoreIn = $hashObjProperties.($ObjectType).($Property)
        }
        else
        {
            $StoreIn = $hashObjProperties.($ObjectType)
        }
        $StoreIn.Add($ObjectName.($ObjID),@{})

        # Collect properties based on $ReportObjects
        foreach ($Attr in $Attributes)
        {
            if ($ReportObjects.($ObjectType) -contains $Attr.AttributeName)
            {
                
                if ($Attr.IsMultiValue -eq 'true') 
                {
                    $StoreIn.($ObjectName.($ObjID)).Add($Attr.AttributeName,@())
                    foreach ($value in $Attr.Values.string)
                    {
                        $StoreIn.($ObjectName.($ObjID)).($Attr.AttributeName) += (ReplaceGUIDs -Text $ObjectName.($value) -hashIDtoName $ObjectName)
                    }
                }
                elseif ($Attr.AttributeName -eq "Filter") 
                {
                    $value = ReplaceGUIDs -Text ($Attr.Value.Replace($FilterPrefix,"").Replace($FilterEnd,"")) -hashIDtoName $ObjectName
                    $StoreIn.($ObjectName.($ObjID)).Add($Attr.AttributeName,$value)
                }
                else 
                {
                    $value = ReplaceGUIDs -Text $Attr.Value -hashIDtoName $ObjectName
                    $Attr.AttributeName
                    $StoreIn.($ObjectName.($ObjID)).Add($Attr.AttributeName,$value)
                }
            }
        }
    }
}



###
### Create New Word Document
###

$doc = StartDoc -Orientation "Landscape" -Visible $Visible
TestStyles -selection $doc

### Schema

if ($IncludeSchema)
{
    $DocItem = "Schema"
    WriteLine -selection $doc -Style $Heading1 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    foreach ($ObjectType in $hashBindings.Keys | sort)
    {
        if ((-not $ShowObjectTypes -or -not $ShowObjectTypes.count) -or ($ShowObjectTypes -contains $ObjectType))
        {
            WriteLine -selection $doc -Style $Heading2 -Text $ObjectType
    
            $table = StartTable -selection $doc -Headings @("Binding","Description","Attribute","DataType","MultiValued","Required","Validation")
            $r = 1

            foreach ($Attribute in $hashBindings.($ObjectType).Keys | sort)
            {
                $ColumnText = @()
                $ColumnText += $hashBindings.($ObjectType).($Attribute).DisplayName
                $ColumnText += $hashBindings.($ObjectType).($Attribute).Description
                $ColumnText += $Attribute
                $ColumnText += $hashBindings.($ObjectType).($Attribute).DataType
                $ColumnText += $hashBindings.($ObjectType).($Attribute).MultiValued
                $ColumnText += $hashBindings.($ObjectType).($Attribute).Required
                $ColumnText += $hashBindings.($ObjectType).($Attribute).StringRegex
                $r += 1
                AddTableRow -table $table -row $r -ColumnText $ColumnText
            }
            # Move to the end of the table
            $doc.EndOf(15)
            $doc.EndOf(6)
            $doc.MoveDown()
        }
    }
}

### MPRs

if ($IncludeMPRs)
{
    $DocItem = "MPR"
    WriteLine -selection $doc -Style $Heading1 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    $DocItem = "MPR-AC"
    WriteLine -selection $doc -Style $Heading2 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}


    ## By Requestor Set

    $DocItem = "MPR-AC-ReqSet"
    WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    foreach ($SetName in $hashMPRGRSet.Keys | sort)
    {
        WriteLine -selection $doc -Style $Heading4 -Text "Set: $SetName"

        $table = StartTable -selection $doc -Headings @("MPR","Description","Actions","Attributes","Target Before Request","Target After Request")
        $r = 1

        foreach ($MPRName in $hashMPRGRSet.($SetName).Keys | sort)
        {
            $ColumnText = @()
            $ColumnText += $MPRName
            $ColumnText += $hashMPRGRSet.($SetName).($MPRName).("Description")
            $ColumnText += $hashMPRGRSet.($SetName).($MPRName).("ActionType") -join "`n"
            $ColumnText += $hashMPRGRSet.($SetName).($MPRName).("ActionParameters") -join ", "
            $ColumnText += $hashMPRGRSet.($SetName).($MPRName).("ResourceCurrentSet")
            $ColumnText += $hashMPRGRSet.($SetName).($MPRName).("ResourceFinalSet")
            $r += 1
            AddTableRow -table $table -row $r -ColumnText $ColumnText
        }
        # Move to the end of the table
        $doc.EndOf(15)
        $doc.EndOf(6)
        $doc.MoveDown()
    }


    ## Requestor Relative

    $DocItem = "MPR-AC-ReqRel"
    WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    foreach ($Attr in $hashMPRGRRel.Keys | sort)
    {
        WriteLine -selection $doc -Style $Heading4 -Text ("Relation to Target: {0}" -f $Attr)

        $table = StartTable -selection $doc -Headings @("MPR","Description","Actions","Attributes","Target Before Request","Target After Request")   
        $r = 1

        foreach ($MPRName in $hashMPRGRRel.($Attr).Keys | sort)
        {
            $ColumnText = @()
            $ColumnText += $MPRName
            $ColumnText += $hashMPRGRRel.($Attr).($MPRName).("Description")
            $ColumnText += $hashMPRGRRel.($Attr).($MPRName).("ActionType") -join "`n"
            $ColumnText += $hashMPRGRRel.($Attr).($MPRName).("ActionParameters") -join ", "
            $ColumnText += $hashMPRGRRel.($Attr).($MPRName).("ResourceCurrentSet")
            $ColumnText += $hashMPRGRRel.($Attr).($MPRName).("ResourceFinalSet")
            $r += 1
            AddTableRow -table $table -row $r -ColumnText $ColumnText
        }
        # Move to the end of the table
        $doc.EndOf(15)
        $doc.EndOf(6)
        $doc.MoveDown()
    }


    ### Policies that run Workflows

    $DocItem = "MPR-WF"
    WriteLine -selection $doc -Style $Heading2 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}


    ## Request-Based Workflow MPRs

    $DocItem = "MPR-WF-Req"
    WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    foreach ($Attr in $hashMPRWFReq.Keys | sort)
    {
        $heading = $Attr
        if ($Attr -eq '*') {$heading = $Attr + " (Any)"}

        WriteLine -selection $doc -Style $Heading4 -Text "Change to Attribute: $heading"

        $table = StartTable -selection $doc -TableStyle $TableStyleMultiColumn -Headings @("MPR","Description","Requestor","Actions","Target Before Request","Target After Request","Authorization Workflow","Action Workflow")   
        $r = 1

        foreach ($MPRName in $hashMPRWFReq.($Attr).Keys | sort)
        {                                                                                                                                                                    
            $ColumnText = @()
            $ColumnText += $MPRName
            $ColumnText += $hashMPRWFReq.($Attr).($MPRName).("Description")
            $ColumnText += $hashMPRWFReq.($Attr).($MPRName).("PrincipalSet")
            $ColumnText += $hashMPRWFReq.($Attr).($MPRName).("ActionType") -join "`n"
            $ColumnText += $hashMPRWFReq.($Attr).($MPRName).("ResourceCurrentSet")
            $ColumnText += $hashMPRWFReq.($Attr).($MPRName).("ResourceFinalSet")
            $ColumnText += $hashMPRWFReq.($Attr).($MPRName).("WFAuthZs") -join "`n"
            $ColumnText += $hashMPRWFReq.($Attr).($MPRName).("WFActions") -join "`n"
            $r += 1
            AddTableRow -table $table -row $r -ColumnText $ColumnText
        }
        $doc.EndOf(15)
        $doc.EndOf(6)
        $doc.MoveDown()
    }


    ## Set Transition Workflow MPRs

    $DocItem = "MPR-WF-Trans"
    WriteLine -selection $doc -Style $Heading3 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    foreach ($SetName in $hashMPRWFTrans.Keys | sort)
    {
        WriteLine -selection $doc -Style $Heading4 -Text "Transition Set: $setName"

        $table = StartTable -selection $doc -TableStyle $TableStyleMultiColumn -Headings @("MPR","Description","Transition Direction","Authorization Workflow","Action Workflow")
        $r = 1

        foreach ($MPRName in $hashMPRWFTrans.($SetName).Keys | sort)
        {                                                                                                                                                                    
            $ColumnText = @()
            $ColumnText += $MPRName
            $ColumnText += $hashMPRWFTrans.($SetName).($MPRName).("Description")
            $ColumnText += $hashMPRWFTrans.($SetName).($MPRName).("ActionType")
            $ColumnText += $hashMPRWFTrans.($SetName).($MPRName).("WFAuthZs") -join "`n"
            $ColumnText += $hashMPRWFTrans.($SetName).($MPRName).("WFActions") -join "`n"
            $r += 1
            AddTableRow -table $table -row $r -ColumnText $ColumnText
        }
        $doc.EndOf(15)
        $doc.EndOf(6)
        $doc.MoveDown()

    }
}


### Sets
if ($IncludeSets)
{
    $DocItem = "Set"
    WriteLine -selection $doc -Style $Heading1 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    $table = StartTable -selection $doc -Headings @("Set","Description","MPRs","Manual Members","Criteria")
    $r = 1

    foreach ($SetName in $hashSet.Keys | sort)
    {
        $ColumnText = @()
        $ColumnText += $SetName
        $ColumnText += $hashSet.($SetName).("Description")
        $ColumnText += $hashMPRSets.($SetName) -join "`n"
        $ColumnText += $hashSet.($SetName).("ExplicitMembers") -join "`n"
        $ColumnText += $hashSet.($SetName).("Filter")
        $r += 1
        AddTableRow -table $table -row $r -ColumnText $ColumnText

        # Add Bookmark
        #$doc.MoveDown()
        #$doc.Bookmarks.Add($hashSet.($SetName).("Bookmark"))
    }
    $doc.EndOf(15)
    $doc.EndOf(6)
    $doc.MoveDown()
}


### Workflows

if ($IncludeWFs)
{
    $DocItem = "WF"
    WriteLine -selection $doc -Style $Heading1 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    $table = StartTable -selection $doc -TableStyle $TableStyleMultiColumn -Headings @("Workflow","Description","Phase","MPRs","Steps")
    $r = 1

    foreach ($WFName in $hashWF.Keys | sort)
    {
        $ColumnText = @()
        $ColumnText += $WFName
        $ColumnText += $hashWF.($WFName).("Description")
        $ColumnText += $hashWF.($WFName).("RequestPhase")
        $ColumnText += $hashMPRWFs.($WFName) -join "`n"
        $r += 1
        AddTableRow -table $table -row $r -ColumnText $ColumnText

        #Create a sub-table for WF steps
        $subtable = StartTable -selection $table.cell($r,5) -Columns 2  -Subtable $true -TableStyle $SubTableStyle
        $sr = 0
        foreach ($wfStep in $hashWF.($WFName).("Steps").Keys | sort)
        {
            $sr += 1
            if ($sr -eq 1) {$NotFirstRow = $false} else {$NotFirstRow = $true}
            $steps = $hashWF.($WFName).("Steps").($wfStep) | sort
            AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($wfStep,($steps -join "`n"))
        }

    }
    $doc.EndOf(15)
    $doc.EndOf(6)
    $doc.MoveDown()
}



### Email Templates
## Note: This section uses a temporary file to save the HTML, which allows it to be included and rendered in the Word document.
[string]$TempFile = $ScriptFldr + "\et.html"
if ($IncludeEmailTemplates)
{
    $DocItem = "ET"
    WriteLine -selection $doc -Style $Heading1 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    ## Write Tables
    $table = StartTable -selection $doc -TableStyle $TableStyleMultiColumn -Headings @("Email Template","Template Type","Subject","Body")
    $r = 1

    foreach ($ETName in $hashET.Keys | sort)
    {
        $ColumnText = @()
        $ColumnText += $ETName
        $ColumnText += $hashET.($ETName).("TemplateType")
        $ColumnText += $hashET.($ETName).("Subject")
        #$ColumnText += $hashET.($ETName).("Body")
        $r += 1
        AddTableRow -Table $table -Row $r -ColumnText $ColumnText

        ## To insert the template contents as rendered HTML we need to save it to a file and then insert it as a text object.
        $hashET.($ETName).("Body") | Out-File -FilePath $TempFile -Encoding default
        $doc.EndOf(15)
        $doc.InsertFile($TempFile)
    }
    $doc.EndOf(15)
    $doc.EndOf(6)
    $doc.MoveDown()
}
Remove-Item $TempFile
# End Email Templates


### UI Resources
if ($IncludeUI)
{
    $DocItem = "UI"
    WriteLine -selection $doc -Style $Heading1 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    ## Navigation Bar Resources
    $DocItem = "UI-NBR"
    WriteLine -selection $doc -Style $Heading2 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    $table = StartTable -selection $doc -TableStyle $TableStyleMultiColumn -FontSize 8 -Headings @("Menu","SubMenu","Description","Usage Keywords","URL")
    $r = 1

    foreach ($ParentOrder in $hashNBR.Keys | sort)
    {
        $first = $true
        foreach ($Order in $hashNBR.($ParentOrder).Keys | sort)
        {
            foreach ($DisplayName in $hashNBR.($ParentOrder).($Order).Keys)
            {
                $ColumnText = @()
                if ($first) {$ColumnText += $DisplayName; $ColumnText += "" }
                else {$ColumnText += ""; $ColumnText += $DisplayName}
                $ColumnText += $hashNBR.($ParentOrder).($Order).($DisplayName).Description
                $ColumnText += $hashNBR.($ParentOrder).($Order).($DisplayName).UsageKeyword
                $ColumnText += $hashNBR.($ParentOrder).($Order).($DisplayName).NavigationUrl
     
                $r += 1
                AddTableRow -table $table -row $r -ColumnText $ColumnText
            }
            $first = $false
        }
    }

    # Move to end of table
    $doc.EndOf(15)
    $doc.EndOf(6)
    $doc.MoveDown()


    ## Home Page Resources

    $DocItem = "UI-HPR"
    WriteLine -selection $doc -Style $Heading2 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    $table = StartTable -selection $doc -TableStyle $TableStyleMultiColumn -FontSize 8 -Headings @("Menu","SubMenu","Description","Usage Keywords","URL")
    $r = 1

    foreach ($ParentOrder in $hashHPR.Keys | sort)
    {
        $first = $true
        foreach ($Order in $hashHPR.($ParentOrder).Keys | sort)
        {
            foreach ($DisplayName in $hashHPR.($ParentOrder).($Order).Keys)
            {
                $ColumnText = @()
                if ($first) {$ColumnText += $DisplayName; $ColumnText += "" }
                else {$ColumnText += ""; $ColumnText += $DisplayName}
                $ColumnText += $hashHPR.($ParentOrder).($Order).($DisplayName).Description
                $ColumnText += $hashHPR.($ParentOrder).($Order).($DisplayName).UsageKeyword
                $ColumnText += $hashHPR.($ParentOrder).($Order).($DisplayName).NavigationUrl
     
                $r += 1
                AddTableRow -table $table -row $r -ColumnText $ColumnText
            }
            $first = $false
        }
    }

    # Move to end of table
    $doc.EndOf(15)
    $doc.EndOf(6)
    $doc.MoveDown()



    ## Search Scopes

    $DocItem = "UI-SS"
    WriteLine -selection $doc -Style $Heading2 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}

    foreach ($ObjectType in $hashSS.Keys | sort)
    {
        WriteLine -selection $doc -Style $Heading3 -Text "Object Type: $ObjectType"

        $table = StartTable -selection $doc -TableStyle $TableStyleMultiColumn -FontSize 8 -Headings @("Search Scope","Filter","Columns","Usage Keywords","Advanced Filter")
        $r = 1
        foreach ($DisplayName in $hashSS.($ObjectType).Keys)
        { 
            $ColumnText = @()           
            $ColumnText += $DisplayName
            $ColumnText += $hashSS.($ObjectType).($DisplayName).Filter
            $ColumnText += $hashSS.($ObjectType).($DisplayName).Columns -join ", "
            $ColumnText += $hashSS.($ObjectType).($DisplayName).UsageKeyword
            $ColumnText += $hashSS.($ObjectType).($DisplayName).AdvancedFilter           
 
            $r += 1
            AddTableRow -table $table -row $r -ColumnText $ColumnText
        }
        
        # Move to end of table
        $doc.EndOf(15)
        $doc.EndOf(6)
        $doc.MoveDown()
    }




    ### RCDCs

    $DocItem = "UI-RCDC"
    WriteLine -selection $doc -Style $Heading2 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}


    ## Section for each Object type

    foreach ($ObjectType in $hashRCDC.Keys | sort)
    {
        WriteLine -selection $doc -Style $Heading2 -Text $ObjectType

        foreach ($RCDCName in $hashRCDC.($ObjectType).Keys | sort)
        {
            WriteLine -selection $doc -Style $Heading3 -Text $RCDCName

            $r = 1
            $table = StartTable -selection $doc -TableStyle $TableStyleTwoColumn -FontSize $TableFontSize -Headings @("Name",$RCDCName)

            $r += 1
            AddTableRow -table $table -row $r -ColumnText @("Applies To",($hashRCDC.($ObjectType).($RCDCName).Actions -join ', '))

            $r += 1
            AddTableRow -table $table -row $r -ColumnText @("String Resources","")
            $subtable = StartTable -selection $table.cell($r,2) -Subtable $true -TableStyle $SubTableStyle -Columns 2
            $sr = 0
            foreach ($item in $hashRCDCStrings.($RCDCName).Keys | sort)
            {
                $sr += 1
                if ($sr -eq 1) {$NotFirstRow = $false} else {$NotFirstRow = $true} 
                AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($item,$hashRCDCStrings.($RCDCName).($item))
            }

            foreach ($rcdcGroup in $hashRCDC.($ObjectType).($RCDCName).Tabs.Keys | sort)
            {
                $r += 1
                AddTableRow -table $table -row $r -ColumnText @(("RCDC Tab " + $rcdcGroup.split(" ")[1]),"")
                $subtable = StartTable -selection $table.cell($r,2) -TableStyle $SubTableStyle -Columns 2 -Subtable $true
                $sr = 0
                foreach ($item in $hashRCDC.($ObjectType).($RCDCName).Tabs.($rcdcGroup).Keys | sort)
                {
                    $sr += 1
                    if ($sr -eq 1) {$NotFirstRow = $false} else {$NotFirstRow = $true} 
 
                    $ColumnText = @()           
                    $ColumnText += $item.split(" ")[1]
                    $ColumnText += ($hashRCDC.($ObjectType).($RCDCName).Tabs.($rcdcGroup).($item) | sort) -join "`n"

                    AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText $ColumnText
                 }
            }

            # Move to end of table
            $doc.EndOf(15)
            $doc.MoveDown()
        }
        # Move to end of table
        $doc.EndOf(15)
        $doc.MoveDown()
    }
}
# End UI


### Other Objects
if ($IncludeOtherObjects)
{
    $DocItem = "Objects"
    WriteLine -selection $doc -Style $Heading1 -Text $DocText.($DocItem).Heading
    if ($DocText.($DocItem).Description) {WriteLine -selection $doc -Style $Normal -FontSize $NormalSize -Text $DocText.($DocItem).Description}


    foreach ($ObjectType in $hashObjProperties.Keys | sort)
    {
        WriteLine -selection $doc -Style $Heading2 -Text $ObjectType

        if ($SubSections.ContainsKey($ObjectType))
        {
            foreach ($Section in $hashObjProperties.($ObjectType).Keys)
            {
                WriteLine -selection $doc -Style $Heading3 -Text $Section
                WriteSection -selection $doc -hashtable $hashObjProperties.($ObjectType).($Section)
            }
        }

        else
        {
             WriteSection -selection $doc -hashtable $hashObjProperties.($ObjectType)
        }

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
        $SaveFilePath = $env:temp + "\FIMPortal.docx"
        $doc.Document.SaveAs([ref]$SaveFilePath,[ref]$saveFormat::wdFormatDocumentDefault)

        write-host "Document saved as $SaveFilePath"
    }
    Catch {write-error $Error[0]}
}
