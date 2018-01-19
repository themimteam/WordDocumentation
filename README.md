# MIMTeam.WordDocumentation
Automatic Word Documentation

# Author
Carol Wapshere

The scripts on this folder are used to produce two Word documents from exported FIM configuration files.

1. Document_FIMPortal.ps1 takes the policy.xml and schema.xml export files and produces a Word document covering Schema, Policy, Portal UI and optionally other object types.
1. Document_SyncConfig.ps1 takes the folder location of a Sync Server configuration export and produces a Word document showing MA and flow rule configuration.

**All of the scripts on this page are needed to produce a document.**

BEFORE RUNNING THIS SCRIPT:

* Export the schema and policy using the standard configuration migration scripts http://technet.microsoft.com/en-us/library/ff400275(v=ws.10).aspx
** Note if you want to include Groups or custom object types in the document you must modify the ExportPolicy.ps1 script, 
** Eg, $policy = Export-FIMConfig -policyConfig -portalConfig -customConfig ("/Group", "/Role") -MessageSize 9999999
** You must also modify the $ReportObjects hashtable in the Include_CustomisedContent.ps1 script.
* Export the Sync Server config using the option in the UI.
** Note: if you use Sync Rules see the comment in the Document_SynConfig script about producing a CSV that links Sync Rule name to Metaverse ObjectID.
* Modify the Include_STYLES.ps1 script to specify a different Word template and/or styles.
* Modify the Include_CustomisedContent.ps1 script to change document wording, which objects and parameters are included, and any custom object types.
* Run the scripts specifying the full paths to the exported config.