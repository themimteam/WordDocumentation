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
###  Include_STYLES.ps1
###
###  Written by Carol Wapshere
###
###  Sets Word template and styles used by the Document_FIMPortal and Document_SyncConfig scripts. Must be in the same folder or the full path specified as a script parameter.
###  This script may be modified to change the template and styles.
###  




###
### Template - If $TemplateFile is null then Normal template is used
###

#$TemplateFile = "C:\Users\Carol\Documents\Custom Office Templates\My Custom Template.dotx"



###
### Text Styles
###

$Heading1 = "Heading 1"
$Heading2 = "Heading 2"
$Heading3 = "Heading 3"
$Heading4 = "Heading 4"
$Normal = "Normal"
$NormalFontSize = 9
$Caption = "Caption"



###
### Table Styles
###

$TableStyleTwoColumn = "Medium Shading 1 - Accent 1"
$TableStyleMultiColumn = "Medium Shading 1 - Accent 1"
$SubTableStyle = "Table Grid Light"
$TableFontStyle = "Normal"
$TableFontSize = 8


