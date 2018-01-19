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
###  Include_WordFunctions.ps1
###
###  Written by Carol Wapshere
###
###  Functions used by the Document_FIMPortal and Document_SyncConfig scripts. Must be in the same folder.
###
###  Changelog:
###    CW 11/08/2014 - Added "Visible" option to the StartDoc function
###                  - Added "Autofit" option to the StartTable function 


Function StartDoc
{
    PARAM([string]$Orientation = "Portrait",[boolean]$Visible=$true)
    END
    {
        if ($Orientation.ToLower() -eq "landscape") {$o = 1}
        else {$o = 0}

        $word=new-object -ComObject "Word.Application"
        if ($TemplateFile -ne "")
        {
            if (-not $TemplateFile) {$document=$word.documents.Add()}
            elseif (-not (Test-Path $TemplateFile)) {Throw "$TemplateFile not found"}
            else {$document=$word.documents.Add($TemplateFile)}
        }
        else
        {
            $document=$word.documents.Add()
        }
        $word.Visible=$Visible
        $document.PageSetup.Orientation = $o
        $selection=$word.Selection
        #Move to end of document
        $selection.EndOf(6) | out-null #out-null prevents the output messing up the return value 
        Return $selection
    }
}

Function WriteLine
{
    PARAM(
          $selection,
          [string]$Style = $Normal,
          [int]$FontSize,
          [string]$Text
         )
    END
    {
        $selection.Style = $Style
        if ($FontSize) {$selection.Font.Size = $FontSize}
        $selection.TypeText($Text) | out-null
        $selection.TypeParagraph() | out-null
    }
}


Function ReadFileContents
{
    PARAM($FilePath)
    END
    {
        if (Test-Path -Path $FilePath)
        {
            $contents = get-content $FilePath
            Return $contents
        }
        else
        {
            write-host "File not found $FilePath"
            exit
        }
    }
}


Function StartTable
{
    PARAM(
          $Selection,
          [string]$TableStyle=$TableStyleMultiColumn,
          [string]$FontStyle=$TableFontStyle,
          [int]$FontSize=$TableFontSize,
          [int]$Columns,
          [system.array]$Headings,
          $Subtable = $false,
          $HeaderRow = $true,
          $AutoFit = $true
         )
    END
    {
        if ($Subtable) 
        {
            $range = $selection.range
        }
        else
        {
            $paragraph = $Selection.Paragraphs.Add()
            $range = $paragraph.Range
        }

        # Create table. If Headings provided use that for column count.
        if ($Headings)
        {
            $Table = $selection.Tables.add($range,1,$Headings.Count)
        }
        else
        {
            $Table = $selection.Tables.add($range,1,$Columns)
        }
        
        # Table style and font
        if ($AutoFit) {$Table.AutoFitBehavior(1) | out-null}
        if ($TableStyle) {$table.Style = $TableStyle}
        if ($FontStyle) {$table.Range.Style = $TableFontStyle}
        if ($FontSize) {$table.Range.Font.Size = $FontSize}
        $table.ApplyStyleHeadingRows = $HeaderRow
 
        $Table.style.ParagraphFormat.SpaceAfter = 0

        # Fill in headings if provided
        if ($Headings)
        {
            $c = 1
            foreach ($heading in $Headings)
            {
                $Table.cell(1,$c).range.text = $Headings[$c-1]
                $c += 1
            }
        }

        Return $Table
    }
}


Function AddTableRow
{
    PARAM(
          $Table,
          [int]$Row,
          [system.array]$ColumnText,
          [boolean]$AddRow = $true
         )
    END
    {           
        if ($AddRow) {$Table.Rows.Add()}
        $c = 1
        foreach ($Text in $ColumnText)
        {
            $Table.cell($Row,$c).range.text = $Text
            $c += 1
        }
    }
}


Function ReplaceGUIDs
{
    PARAM([string]$Text,$hashIDtoName,$prefix="urn:uuid:")
    END
    {      
        do
        {
            if ($matches)
            {
                $name = "GUID"
                $guid = $matches[0]
                if ($hashIDtoName.ContainsKey($guid)) {$name = $hashIDtoName.Item($guid)}
                $Text = $Text.Replace($matches[0],$name)
            }
        } while ($Text -match $regexUrnGUID)

        $matches = $null
        do
        {
            if ($matches)
            {
                $name = "GUID"
                $guid = $prefix + $matches[0]
                if ($hashIDtoName.ContainsKey($guid)) {$name = $hashIDtoName.Item($guid)}
                $Text = $Text.Replace($matches[0],$name)
            }
        } while ($Text -match $regexGUID)

        Return $Text
    }
}

Function WriteSection
{
    PARAM($selection,$hashtable)
    END
    {
        $table = StartTable -selection $selection -TableStyle $TableStyleMultiColumn -Headings @("DisplayName","Properties")
        $r = 1
    
        foreach ($ObjName in $hashtable.Keys | sort)
        {
            $r += 1
            AddTableRow -Table $table -Row $r -ColumnText @($ObjName,"")

            if ($hashtable.Item($ObjName).count -gt 0) {
                #Create a sub-table for object properties
                $subtable = StartTable -selection $table.cell($r,2) -TableStyle $SubTableStyle -Columns 2 -Subtable $true
                $sr = 0

                foreach ($property in $hashtable.Item($ObjName).Keys | sort)
                {           
                    $sr += 1
                    if ($sr -eq 1) {$NotFirstRow = $false} else {$NotFirstRow = $true} 
                    $value = $hashtable.Item($ObjName).Item($property)
                
                    if ($value.GetType().BaseType.Name -eq "Array")
                    {
                        AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($property,($value -join "`n"))
                    }
                    else
                    {
                        AddTableRow -Table $subtable -Row $sr -AddRow $NotFirstRow -ColumnText @($property,$value)
                    }
                }

            }

        }

        $selection.EndOf(15)
        $selection.MoveDown()
    }
}



## This function tests the styles set in the STYLES script exist in this instance of Word.
Function TestStyles
{
    PARAM(
          $selection
         )
    END
    {
        ## Font Styles

        # Heading1
        if (-not $Heading1) {Throw "Heading1 style must be set in STYLES.ps1"}
        Try {$selection.Style = $Heading1}
        Catch {Throw "Failed to set style $Heading1"}

        # Heading2
        if (-not $Heading1) {Throw "Heading2 style must be set in STYLES.ps1"}
        Try {$selection.Style = $Heading2}
        Catch {Throw "Failed to set style $Heading2"}

        # Heading3
        if (-not $Heading3) {Throw "Heading3 style must be set in STYLES.ps1"}
        Try {$selection.Style = $Heading3}
        Catch {Throw "Failed to set style $Heading3"}

        # Heading4
        if (-not $Heading4) {Throw "Heading4 style must be set in STYLES.ps1"}
        Try {$selection.Style = $Heading4}
        Catch {Throw "Failed to set style $Heading4"}

        # Normal
        if (-not $Normal) {Throw "Normal style must be set in STYLES.ps1"}
        Try {$selection.Style = $Normal}
        Catch {Throw "Failed to set style $Normal"}

        # Caption
        if (-not $Caption) {Throw "Caption style must be set in STYLES.ps1"}
        Try {$selection.Style = $Caption}
        Catch {Throw "Failed to set style $Caption"}

        # TableFontStyle
        if (-not $TableFontStyle) {Throw "TableFontStyle style must be set in STYLES.ps1"}
        Try {$selection.Style = $TableFontStyle}
        Catch {Throw "Failed to set style $TableFontStyle"}


        ## Table Styles
        $paragraph = $Selection.Paragraphs.Add()
        $range = $paragraph.Range
        $TestTable = $selection.Tables.add($range,1,2)

        # TableStyleMultiColumn
        Try {$TestTable.Style = $TableStyleMultiColumn}
        Catch {Throw "Failed to set table style $TableStyleMultiColumn"}

        # SubTableStyle
        Try {$TestTable.Style = $SubTableStyle}
        Catch {Throw "Failed to set table style $SubTableStyle"}

        $TestTable.Delete()
    }
}


###
### Find and Replace constants
###

$regexGUID = "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
$regexUrnGUID = "urn:uuid:[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
$regexSyncGUID = "\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}"
$FilterPrefix = '<Filter xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" Dialect="http://schemas.microsoft.com/2006/11/XPathFilterDialect" xmlns="http://schemas.xmlsoap.org/ws/2004/09/enumeration">'
$FilterEnd = '</Filter>'




