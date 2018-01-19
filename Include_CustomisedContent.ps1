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
###  Include_CustomisedContent.ps1
###
###  Written by Carol Wapshere
###
###  Customises the content included in the Document_FIMPortal and Document_SyncConfig scripts. Must be in the same folder.
###  Modify this script to change headings, text and which objects and which parameters are shown.
###


### Section headings and descriptions

# Change Heading and Description text here to change what appears in the document.
# - Do not add or remove entries to this hashtable (without also changing the Document_FIMPortal script)
# - Do not change hashtable keys
# - The Description element may be deleted if not needed
$DocText = @{}
# Portal document
$DocText.Add("Schema",@{"Heading"="Portal Schema";"Description"="This section shows the Object Types defined in the FIM Portal schema, and lists their bound attributes."})
$DocText.Add("MPR",@{"Heading"="Management Policy Rules";"Description"="Management Policy Rules define both access permission in the Portal, and the Workflows that should be run when particular changes are made to objects."})
$DocText.Add("MPR-AC",@{"Heading"="Access Permission Policies";"Description"="This section lists Management Policy Rules that grant access to read or change data in the Portal."})
$DocText.Add("MPR-AC-ReqSet",@{"Heading"="Requestor Sets";"Description"="The following sub-sections show the Management Policy Rules which grant access to members of Sets."})
$DocText.Add("MPR-AC-ReqRel",@{"Heading"="Requestor Relative to Target";"Description"="The following sub-sections show the Management Policy Rules which grant access where a relational link exists between the Requestor and the Target of the request."})
$DocText.Add("MPR-WF",@{"Heading"="Workflow Policies";"Description"="This section lists Management Policy Rules that trigger Workflows."})
$DocText.Add("MPR-WF-Req",@{"Heading"="Workflows Triggered when an Attribute is Changed";"Description"="The following sections show the Management Policy Rules which trigger workflows when a particular attribute is changed."})
$DocText.Add("MPR-WF-Trans",@{"Heading"="Workflows Triggered when a Set Transition Occurs";"Description"="The following sections show the Management Policy Rules which trigger workflows when an object transitions in or out of a set."})
$DocText.Add("Set",@{"Heading"="Sets";"Description"="Sets are collections of objects that represent some state. They may be criteria-based or mannually-populated, and are used both in granting access permissions in the Portal, and in triggering actions when objects transition between Sets."})
$DocText.Add("WF",@{"Heading"="Workflows";"Description"="Workflows define activities run by the FIM Service in response to requests and changes. Each Workflow is defined as either 'Authentication' (used in identifying the Requestor), 'Authorization' (used to evaluate if an action is permitted) or 'Action' (run after a change to trigger other changes and events)."})
$DocText.Add("ET",@{"Heading"="Email Templates";"Description"="Email Templates are used by Notification and Approval Workflows to construct the contents of the emails sent. They include variables in the format [//Object/Parameter] which are resolved at run time."})
$DocText.Add("UI",@{"Heading"="Portal Interface";"Description"="The following sections show the configuration of objects used to customise the Portal interface."})
$DocText.Add("UI-NBR",@{"Heading"="Sidebar Menu Options";"Description"="The following resources are used to generate the sidebar menu available in the Portal. The Usage Keywords are used to determine who can see particular options. Those tagged 'BasicUI' are available to all users."})
$DocText.Add("UI-HPR",@{"Heading"="Home Page Options";"Description"="The following resources are used to generate the options available on the Portal home page. The Usage Keywords are used to determine who can see particular options. Those tagged 'BasicUI' are available to all users."})
$DocText.Add("UI-SS",@{"Heading"="Search Scopes";"Description"="Search Scopes filter the objects available in a Portal view. This section lists the customised Search Scopes added to this solution. The Usage Keywords are used to determine who can see particular options. Those tagged 'BasicUI' are available to all users."})
$DocText.Add("UI-RCDC",@{"Heading"="Portal RCDC Forms";"Description"="The following sections show the forms configured for each listed object type."})
$DocText.Add("Objects",@{"Heading"="Solution Objects";"Description"="The following objects represent significant aspects of the solution so are listed in full."})

# Sync document
$DocText.Add("Sync-Main",@{"Heading"="Synchronization Service Configuration";"Description"=""})
$DocText.Add("Sync-Global",@{"Heading"="Global Settings";"Description"="The following global settings apply to the configuration of the Sync Service."})
$DocText.Add("Sync-FlowMap",@{"Heading"="Metaverse Attribute Flows";"Description"="The following sections show the end-to-end flows based on Metaverse attribute. Details about advanced or SyncRule flow configurations are included in the MA sections below."})
$DocText.Add("Sync-CF",@{"Heading"="Connector Filters";"Description"="An object that matches a Connector Filter is prevented from joining to a Metaverse object. If already joined the connector space object will be disconnected.";"None"="This MA has no connector filters."})
$DocText.Add("Sync-Join",@{"Heading"="Join Rules";"Description"="A disconnected connector space object will be joined the first Metaverse object that matches a join rule.";"None"="This MA has no join rules."})
$DocText.Add("Sync-Projection",@{"Heading"="Projection Rules";"Description"="A Projection Rule allows a Metaverse object to be created for disconnectors in this connector space.";"None"="This MA has no projection rules."})
$DocText.Add("Sync-IAF",@{"Heading"="Import Attribute Flows";"Description"="Import Flow Rules update joined Metaverse objects based on the attribute values on the Connector Space object.";"None"="This MA has no import flow rules."})
$DocText.Add("Sync-EAF",@{"Heading"="Export Attribute Flows";"Description"="Export Flow Rules update joined Connector Space objects based on the attribute values on the Metaverse object.";"None"="This MA has no export flow rules."})


### Schema inclusions

# Object Types to include in the Schema documentation. 
# This array must exist but can be empty [@()] in which case ALL object types are included.
$ShowObjectTypes = @("Person","Group")
#$ShowObjectTypes = @()

### Workflow Activity settings to show. 

# Specify which activity parameters to include in the document.
# If an activity is not listed here all then all parameters are shown but not identified.
$activitySteps = @{}
$activitySteps.Add("FunctionActivity",@("Description","FunctionExpression","Destination"))
$activitySteps.Add("EmailNotificationActivity",@("To","EmailTemplate"))
$activitySteps.Add("PowerShellActivity",@("Script"))
$activitySteps.Add("ApprovalActivity",@("Approvers","Duration","Escalation","ApprovalEmailTemplate","EscalationEmailTemplate","ApprovalDeniedEmailTemplate","ApprovalCompleteEmailTemplate"))
$activitySteps.Add("PWResetActivity",@())


### RCDC inclusions and exclusions

# RCDC target object types to include in the document
$ShowRCDCTypes = @("Person")

# RCDC Tabs not included in the document
$RCDCTabsHide = @("summaryControl","summary","_caption","caption")

# Exclude RCDC String Resources matching the following regex
# This regex will ignore all the country codes: "([A-Z]{2}Caption)"
$RCDCStringsRegexNotMatch = "([A-Z]{2}Caption)"

# Properties of RCDC control to include. If a control type is not listed here then all its properties are included.
$ReportProps = @{}
$ReportProps.Add("UocListView",@("Description","RightsLevel","Required","Hint","ColumnsToDisplay","EnableSelection","SingleSelection","ListFilter"))
$ReportProps.Add("UocCheckBox",@("Description","RightsLevel","Required","Hint","DefaultValue","Text"))
$ReportProps.Add("UocCommonMultiValueControl",@("Description","RightsLevel","Required","Hint","Rows","Columns","Value"))
$ReportProps.Add("UocDropDownList",@("Description","RightsLevel","Required","Hint","ValuePath","ItemSource"))
$ReportProps.Add("UocTextBox",@("Description","RightsLevel","Required","Hint","ReadOnly","Required"))
$ReportProps.Add("UocLabel",@("Description","RightsLevel","Required","Hint"))
$ReportProps.Add("UocIdentityPicker",@("Description","RightsLevel","Required","Hint","AttributesToSearch","ColumnsToDisplay","ObjectTypes","UsageKeywords","Filter","Mode","ResultObjectType"))
$ReportProps.Add("UocDateTimeControl",@("Description","RightsLevel","Required","Hint"))



### Search Scopes

# Include only Search Scopes with the following usage keywords (set to empty array to include all)
$SSUsageKeyword = @("BasicUI","Custom")



### Object Types and their attributes to include (excluding DisplayName which is included by default).

# Add extra object types here if you want to include them. Note you must have modified the ExportPolicy.ps1 script to export the extra object types.
$ReportObjects = @{}
$ReportObjects.Add("Group",@("Owner","Filter","Email","Type","ExplicitMember"))

# (Optional) To break the object list down into sub-sections based on an attribute, add a line here specifying "ObjectType","Attribute" where the subsections will be based on the Attribute.
$SubSections = @{}
$SubSections.Add("Group","Type")


