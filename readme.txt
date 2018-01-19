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
###  Document_FIMPortal.ps1, Document_SyncConfig.ps1
###
###  Written by Carol Wapshere
###
###  Document_FIMPortal.ps1 takes the policy.xml and schema.xml export files and produces a Word document covering Schema, Policy, Portal UI and optionally other object types.
###  Document_SyncConfig.ps1 takes the folder location of a Sync Server configuration export and produces a Word document showing MA and flow rule configuration.
###  Tested with FIM 2010 R2 and Word 2010/2013.
###
###  BEFORE RUNNING THIS SCRIPT:
###
###    1. Export the schema and policy using the standard configuration migration scripts http://technet.microsoft.com/en-us/library/ff400275(v=ws.10).aspx
###        - Note if you want to include Groups or custom object types in the document you must modify the ExportPolicy.ps1 script, eg: 
###             $policy = Export-FIMConfig -policyConfig -portalConfig -customConfig ("/Group", "/Role") -MessageSize 9999999
###          You must also modify the $ReportObjects hashtable in the Include_CustomisedContent.ps1 script.
###
###    2. Export the Sync Server config using the option in the UI.
###
###    3. Modify the Include_STYLES.ps1 script to specify a different Word template and/or styles.
###
###    4. Modify the Include_CustomisedContent.ps1 script to change document wording, which objects and parameters are included, and any custom object types.
###
###    5. Run the scripts specifying the full paths to the exported config.
