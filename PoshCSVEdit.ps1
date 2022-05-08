<#
.SYNOPSIS
    PowerShell CSV Editor. Allows GUI editing of a .csv based on a pre-defined schema.

.DESCRIPTION

.PARAMETER CSVPath
    The path of the CSV to edit.

.PARAMETER SchemaPath
    The path of the schema for the CSV.

.EXAMPLE
    .\PoshCSVEdit.ps1 -CSVPath "C:\Test\test.csv" -SchemaPath "C:\Test\test.json"

.LINK
    https://github.com/ProjectThySoul/PoshCSVEdit

#>

Param (
    [Parameter(Mandatory)]
    [string]$CSVPath,
    [Parameter(Mandatory)]
    [string]$SchemaPath,
    [switch]$ShowLockFile,
    [switch]$HideConsole
)

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓ START OF FUNCTIONS ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
# .Net methods for hiding/showing the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
Function Show-Console {
    <#
    .SYNOPSIS
    Shows the PowerShell console window.

    .DESCRIPTION
    Uses .Net methods to show the PowerShell console window.

    .EXAMPLE
    Show-Console
    #>

    $consolePtr = [Console.Window]::GetConsoleWindow()

    # Hide = 0,
    # ShowNormal = 1,
    # ShowMinimized = 2,
    # ShowMaximized = 3,
    # Maximize = 3,
    # ShowNormalNoActivate = 4,
    # Show = 5,
    # Minimize = 6,
    # ShowMinNoActivate = 7,
    # ShowNoActivate = 8,
    # Restore = 9,
    # ShowDefault = 10,
    # ForceMinimized = 11

    [void] [Console.Window]::ShowWindow($consolePtr, 1)
} # Function Show-Console
Function Hide-Console {
    <#
    .SYNOPSIS
    Hides the PowerShell console window.

    .DESCRIPTION
    Uses .Net methods to hide the PowerShell console window.

    .EXAMPLE
    Hide-Console
    #>
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [void] [Console.Window]::ShowWindow($consolePtr, 0)
} # Function Hide-Console
Function Get-FromJson {
    <#
    .SYNOPSIS
    Reads in a JSON file from the specified path and returns it as a hashtable.

    .DESCRIPTION
    Reads in the specified JSON file, and converts its data to a hashtable.
    This makes it easier to iterate through 'keys' and values.
    ConvertFrom-JSON has an -AsHashTable parameter from version 6, however this function means that we are compatible with lower versions, and also returns the columns in the correct order.
    Adapted from https://stackoverflow.com/questions/40495248/create-hashtable-from-json.

    .PARAMETER Path
    The path of the JSON file to process.

    .EXAMPLE
    Get-FromJson -Path "C:\Test\Test.json"

    .OUTPUTS
    A hashtable containing the data found in the JSON file.
    #>

    Param(
        [Parameter(Mandatory = $True, Position = 1)]
        [string]$Path
    )

    Function Get-Value {

        Param(
            $Value
        )

        $Result = $Null
        If ( $Value -is [System.Management.Automation.PSCustomObject] ) {
            Write-Verbose "Get-Value: value is PSCustomObject"
            $Result = [ordered]@{}
            $Value.PSObject.Properties | ForEach-Object { 
                $Result[$_.Name] = Get-Value -Value $_.Value 
            }
        }
        ElseIf ($Value -is [System.Object[]]) {
            $List = New-Object System.Collections.ArrayList
            Write-Verbose "Get-Value: value is Array"
            $Value | ForEach-Object {
                $List.Add((Get-Value -Value $_)) | Out-Null
            }
            $Result = $List
        }
        Else {
            Write-Verbose "Get-Value: Value is type: $($Value.GetType())"
            $Result = $Value
        }
        Return $Result
    }

    If (Test-Path $Path) {
        $JSON = Get-Content $Path -Raw
    }
    Else {
        # $JSON = '{}'
        Return $Null
    }

    Try {
        $JSON = $JSON | ConvertFrom-JSON
    }
    Catch {
        Return $Null
    }

    $Hashtable = Get-Value -Value $JSON

    Return $Hashtable
}
Function Test-CSVColumns {
    <#
    .SYNOPSIS
    Checks whether the columns in the input CSV match the expected columns in the schema.

    .DESCRIPTION
    Loops through the columns in the passed CSV, and checks whether each has a corresponding entry in the schema.
    Loops through the schema to check that all columns are in the CSV.
    If there is a mismatch, return $False otherwise return $True.

    .PARAMETER CSVData
    The imported CSV data.

    .PARAMETER Schema
    The schema for this CSV.

    .EXAMPLE
    Test-CSVColumns -CSVData $CSVData -Schema $Schema
    #>

    Param (
        [Parameter(Mandatory)]
        [object]$CSVData,
        [Parameter(Mandatory)]
        [hashtable]$Schema
    )

    Try {
        $CSVColumns = $CSVData[0].PSObject.Properties.Name

        # Initial return value
        $ColumnsOK = $True

        # Check CSV columns against Schema to make sure all CSV columns are in the schema.
        ForEach ($CSVColumn In $CSVColumns) {
            If (!($Null -ne $Schema.Columns.$CSVColumn)) {
                # Column is not defined in schema.
                $ColumnsOK = $False
            }
        }

        # Check schema columns against CSV to make sure no columns are missing from the CSV.
        ForEach ($SchemaColumn in $Schema.Columns.Keys) {
            If ($CSVColumns -notcontains $SchemaColumn) {
                # Schema column is not in the CSV.
                $ColumnsOK = $False
            }
        }

        Return $ColumnsOK

    }
    Catch {
        Return $False
    }
        
} # Function Test-CSVColumns
Function Get-FormattedValue {
    <#
    .SYNOPSIS
    Formats the supplied value and returns the formatted version.

    .DESCRIPTION
    Trims leading and trailing whitespace from a supplied value.
    Formats it according to the SetCase parameter.

    .PARAMETER InputValue
    The Value to format.

    .PARAMETER SetCase
    Optional. Sets the case of the InputValue to Title, Upper, Lower.

    .OUTPUTS
    If the InputVaue contains a real string, returns the formatted version of this.
    If the InputValue is Null, or just whitespace then returns $Null.

    .EXAMPLE
    $Value = Get-FormattedValue -InputValue "   Banana " -SetCase Upper
    will return "BANANA"
    #>

    Param (
        [string]$InputValue,
        [ValidateSet("Title", "Upper", "Lower")]
        [string]$SetCase
    )

    $TextInfo = (Get-Culture).TextInfo

    # If the Input Value is not $Null
    If ($InputValue) {

        # Convert to string
        $InputValue = [string]$InputValue

        # Remove leading and trailing whitespace.
        $InputValue = $InputValue.Trim()

        # If the Input Value isn't just whitespace (will be an empty string now after the trim).
        If ($InputValue -ne "") {
            # Input Value is a non-empty string

            # Set Case
            Switch ($SetCase) {
                "Title" {
                    $OutputValue = $TextInfo.ToTitleCase($InputValue.ToLower())   
                }
                "Upper" {
                    $OutputValue = $InputValue.ToUpper()
                }
                "Lower" {
                    $OutputValue = $InputValue.ToLower()
                }
                Default {
                    # Otherwise just return the trimmed value.
                    $OutputValue = $InputValue
                }
            }
        }
        Else {
            # Input Value is an empty string.
            $OutputValue = $Null

        }
    }
    Else {
        # Input Value is $Null
        $OutputValue = $Null
    }

    Return $OutputValue

} # Function Get-FormattedValue
Function Compare-DataTable {
    <#
    .SYNOPSIS
    Compares two data tables to see if they are the same.

    .DESCRIPTION
    Compares a reference data table to a comparison table using a row-by-row check.
    If they're different, return $False, else return $True.

    .PARAMETER ReferenceTable
    The data table to compare against.

    .PARAMETER CompareTable
    The data table to compare.

    .OUTPUTS
    $False if the data tables are different.
    $True if the data tables are the same.

    .EXAMPLE
    Compare-DataTable -ReferenceTable $ReferenceTable -CompareTable $CompareTable
    #>

    Param (
        [System.Data.DataTable]$ReferenceTable,
        [System.Data.DataTable]$CompareTable
    )

    # If the number of rows in the Reference table is different to the number of rows in the Compare table.
    If ($ReferenceTable.Rows.Count -ne $CompareTable.Rows.Count) {
        Return $False
    }

    # If the number of columns in the Reference table is different to the number of columns in the Compare table.
    If ($ReferenceTable.Columns.Count -ne $CompareTable.Columns.Count) {
        Return $False        
    }

    # Reset the RowIndex
    $RowIndex = 0

    # For each row in the reference table.
    ForEach ($Row in $ReferenceTable.Rows) {

        # Reset the ColumnIndex
        $ColumnIndex = 0

        # For each column in the row
        ForEach ($ReferenceColumn in $Row.ItemArray) {

            # Get corresponding column in the compare table.
            $CompareColumn = $CompareTable.Rows[$RowIndex][$ColumnIndex];

            # If the reference column is different to the compare column.
            If ($ReferenceColumn -ne $CompareColumn) {
                # The table is different, so return false.
                Return $False
            }
            $ColumnIndex++
        }
        $RowIndex++
    }

    # If we get here, the tables are the same, so return true.
    Return $True
} # Function Compare-DataTable
Function Get-AdjustedColor {
    <#
    .SYNOPSIS
    Based on the provided colour, return a colour adjusted by the specified correction factor.

    .DESCRIPTION
    For the color provided, adjust it by the specified correction factor and return the adjusted color.
    Adapted from here: https://www.pvladov.com/2012/09/make-color-lighter-or-darker.html
        
    .PARAMETER Color
    The colour to adjust.

    .PARAMETER CorrectionFactor
    The brightness correction factor. Must be between -1 and 1. 
            
    .OUTPUTS
    The adjusted color.

    .EXAMPLE
    Get-AdjustedColor -Color $Color -CorrectionFactor -0.2
    #>

    Param (
        [System.Drawing.Color]$Color,
        [single]$CorrectionFactor
    )

    # Get the RGB values for the provided colour.
    [single]$Red = [single]$Color.R
    [single]$Green = [single]$Color.G
    [single]$Blue = [single]$Color.B
    
    If ($CorrectionFactor -lt 0) {
        $CorrectionFactor = 1 + $CorrectionFactor
        $Red *= $CorrectionFactor
        $Green *= $CorrectionFactor
        $Blue *= $CorrectionFactor
    }
    Else {
        $Red = (255 - $Red) * $CorrectionFactor + $Red
        $Green = (255 - $Green) * $CorrectionFactor + $Green
        $Blue = (255 - $Blue) * $CorrectionFactor + $Blue
    }
    
    Return [System.Drawing.Color]::FromArgb([int]$Red, [int]$Green, [int]$Blue)
} # Function Get-AdjustedColor
Function Set-GridProperties {
    <#
    .SYNOPSIS
    Sets the properties of the DataGridView based on the provided schema.

    .DESCRIPTION
    Sets formatting and read-only status of rows and cells in the DataGridView based on the configuration defined in the Schema.

    .PARAMETER CSVData
    The imported CSV data.

    .PARAMETER Schema
    The schema for this CSV.

    .EXAMPLE
    Test-CSVColumns -CSVData $CSVData -Schema $Schema
    #>

    # Loop through each row in the DataGridView.
    ForEach ($Row In $DataGridView.Rows) {

        # Loop through each column in the schema.
        ForEach ($Key In $Schema.Columns.Keys) {

            # If this isn't a new row.
            If (!($Row.IsNewRow)) {

                # Get the cell
                $Cell = $Row.Cells[$Key]

                # Get the cell value.
                $CellValue = $Row.Cells[$Key].Value

                # If the schema defines a row colour if the value in this column matches a specified value.
                If ($Null -ne $Schema.Columns.$Key.RowColourIfValue.$CellValue) {
     
                    # Set the row colour to the defined colour for the matching value.
                    $Row.DefaultCellStyle.BackColor = $Schema.Columns.$Key.RowColourIfValue.$CellValue

                }

                If ($Schema.Columns.$Key.ColumnReadOnly -eq $True) {
                    # For read only columns, darken the colour slightly.

                    # If the colour isn't nothing (i.e. white)
                    If ($Row.DefaultCellStyle.BackColor.Name -ne "0") {
                        # Darken the colour.
                        $Cell.Style.BackColor = Get-AdjustedColor -Color $Row.DefaultCellStyle.BackColor -CorrectionFactor -0.125
                    }
                    Else {
                        # If we try to darken white, we get black so just set it to light gray.
                        $Cell.Style.BackColor = "LightGray"
                    }
                }

                # If the schema defines that the row should be read only if the value in this column matches a specified value.
                If ($Null -ne $Schema.Columns.$Key.RowReadOnlyIfValue) {
                    # If the value of this cell matches the defined value for which to set this row to read only.
                    If ($Schema.Columns.$Key.RowReadOnlyIfValue -Contains $CellValue) {
                        # Set the row to read only and set the column headers to dark gray.
                        $Row.ReadOnly = $True
                        $Row.HeaderCell.Style.BackColor = "DarkGray"
                    }
                }
            }
        }
    }

    # We have to do a second loop through now that read only stuff has been set per row and column.
    ForEach ($Row in $DataGridView.Rows) {
        # Loop through all cells in the row
        ForEach ($Cell in $Row.Cells) {
            # If it's a Combo Box Cell.
            If ($Cell.GetType().Name -eq "DataGridViewComboBoxCell") {
                # If the cell is read only
                If ($Cell.ReadOnly -eq $True) {
                    # Set the display style to Nothing. This removes the drop-down arrow for read-only combo boxes.
                    # https://docs.microsoft.com/en-us/dotnet/api/System.Windows.Forms.datagridviewcomboboxcell.displaystyle?view=windowsdesktop-6.0#system-windows-forms-datagridviewcomboboxcell-displaystyle
                    $Cell.DisplayStyle = "Nothing"
                }
            }
        }
    }

} # Function Set-GridProperties
Function Set-HelpLabels {
    <#
    .SYNOPSIS
    Displays the help label and text for the current column.

    .DESCRIPTION
    Displays help for the current column based on the values defined in the Schema.
    Allows for basic 'forum-style' formatting of text by enclosing within simple tags - 

    [b]Bold[/b]
    [i]Italic[/i]
    [u]Underline[/u]
    [s]Strikeout[/s]

    This formating is based on an idea from here: https://www.sysnative.com/forums/threads/how-to-make-selected-text-bold-net.23289/post-188071
    Note: You can only apply 1 tag at a time (i.e not nest them). Applying multiple formats to the same text does strange things (like set point size to 7.5). No idea why, and single format is good enough.

    .PARAMETER Cell
    The cell (column) for which to display help.

    .EXAMPLE
    Set-HelpLabels -Cell $Cell
    #>

    Param(
        [System.Windows.Forms.DataGridViewCell]$Cell
    )

    $ColumnName = $Cell.OwningColumn.Name
    
    $ColumnHelpText = $Schema.Columns.$ColumnName.HelpText

    # If the cell is read only.
    If ($Cell.ReadOnly) {
        # Display a warning to that effect.
        $ColumnHelpText = "[b]This cell is read only.[/b]`r`n`r`n$ColumnHelpText" 
    }
    Else {
        # Otherwise, the cell is editable.
        # If we're not allowing blanks in the column.
        If ($Schema.Columns.$ColumnName.AllowBlank -eq $False) {
            # Display a warning to that effect.
            $ColumnHelpText = "[b]This column is mandatory.[/b]`r`n`r`n$ColumnHelpText" 
        }
    }

    If ($ColumnHelpText) {
        $Label_ColumnHelpTitle.text = "Help for '$ColumnName' column:"
        $Label_ColumnHelpTitle.Visible = $True

        $RichTextBox_ColumnHelpText.text = $ColumnHelpText

        # Define the tags to find, and the formatting.
        $Tags = @{
            Bold      = @{
                StartTag  = "[b]";
                EndTag    = "[/b]";
                FontStyle = [Drawing.FontStyle]::Bold
            };
            Italic    = @{
                StartTag  = "[i]";
                EndTag    = "[/i]";
                FontStyle = [Drawing.FontStyle]::Italic
            };
            Underline = @{
                StartTag  = "[u]";
                EndTag    = "[/u]";
                FontStyle = [Drawing.FontStyle]::Underline
            };
            Strikeout = @{
                StartTag  = "[s]";
                EndTag    = "[/s]";
                FontStyle = [Drawing.FontStyle]::Strikeout
            }
        }

        # Reset formatting
        $ControlFont = $RichTextBox_ColumnHelpText.Font
        $RichTextBox_ColumnHelpText.SelectAll()
        $RichTextBox_ColumnHelpText.SelectionFont = $ControlFont
        $RichTextBox_ColumnHelpText.DeSelectAll()

        # For each of the tags
        ForEach ($TagKey In $Tags.Keys) {
  
            # Get the length of the start tag (used for selection calculation below)
            $StartTagLength = $Tags.$TagKey.StartTag.Length

            # Set the initial values for this pass
            [int]$StartTag = 0
            [int]$EndTag = $RichTextBox_ColumnHelpText.Text.Length;

            # Get the location of the start tag
            $StartTag = $RichTextBox_ColumnHelpText.Text.IndexOf($Tags.$TagKey.StartTag, $StartTag)
            # If we didn't find the start tag
            If ($StartTag -eq -1) {
                # Set the end tag to not found
                $EndTag = -1
            }
            Else {
                # Otherwise find the end tag starting from the end of the start tag.
                $EndTag = $RichTextBox_ColumnHelpText.Text.IndexOf($Tags.$TagKey.EndTag, $StartTag)
            }

            # Whilst a start tag and and end tag are in the string.
            While (($StartTag -ne -1) -and ($EndTag -ne -1)) {

                # Get the location of the start tag and end tag
                $StartTag = $RichTextBox_ColumnHelpText.Text.IndexOf($Tags.$TagKey.StartTag, $StartTag)
                $EndTag = $RichTextBox_ColumnHelpText.Text.IndexOf($Tags.$TagKey.EndTag, $StartTag)

                # Select the tag and text
                $RichTextBox_ColumnHelpText.Select($StartTag + $StartTagLength, $EndTag - $StartTag - $StartTagLength);

                # Get the current font for the selection
                $CurrentFont = $RichTextBox_ColumnHelpText.SelectionFont
                $NewFont = New-Object Drawing.Font($CurrentFont.FontFamily, $CurrentFont.Size, $($Tags.$TagKey.FontStyle))

                # Set the font of the selection to the new font
                $RichTextBox_ColumnHelpText.SelectionFont = $NewFont

                # Set the location of the start tag to that of the end tag, to continue searching from that point.
                $StartTag = $EndTag;

                # Get the location of the next start tag
                $StartTag = $RichTextBox_ColumnHelpText.Text.IndexOf($Tags.$TagKey.StartTag, $StartTag)
                # If we didn't find the start tag
                If ($StartTag -eq -1) {
                    # Set the end tag to not found
                    $EndTag = -1
                }
                Else {
                    # Otherwise find the end tag starting from the end of the start tag.
                    $EndTag = $RichTextBox_ColumnHelpText.Text.IndexOf($Tags.$TagKey.EndTag, $StartTag)
                }
            }

            # remove the start and end tags from the text
            $RichTextBox_ColumnHelpText.Rtf = $RichTextBox_ColumnHelpText.Rtf.Replace($Tags.$TagKey.StartTag, "");
            $RichTextBox_ColumnHelpText.Rtf = $RichTextBox_ColumnHelpText.Rtf.Replace($Tags.$TagKey.EndTag, "");
        }

        $RichTextBox_ColumnHelpText.Visible = $True
    }
    Else {
        $Label_ColumnHelpTitle.Visible = $False
        $RichTextBox_ColumnHelpText.Visible = $False
    }

} # Function Set-HelpLabels
Function Test-RowRequiresValidation {
    <#
    .SYNOPSIS
    Checks whether the specified row requires validation.

    .DESCRIPTION
    A row required validation if - 
    - It's a new row which contains data.
    - It's not a new row.

    .PARAMETER RowIndex
    The index of the row to check.

    .EXAMPLE
    Test-RowRequiresValidation -RowIndex 5
    #>

    Param (
        [int]$RowIndex
    )

    If ($DataGridView.Rows[$RowIndex].IsNewRow) {
        $RowHasData = $False
        ForEach ($Cell in $DataGridView.Rows[$RowIndex].Cells ) {
            If ($Cell.Value) {
                $RowHasData = $True
                Break
            }
        }

        If (!($RowHasData)) {
            ForEach ($Cell in $DataGridView.Rows[$RowIndex].Cells ) {
                $Cell.Value = $Null
                $Cell.ErrorText = $Null
            }
            $DataGridView.Rows[$RowIndex].ErrorText = $Null
            $RowRequiresValidation = $False
        }
        Else {
            $RowRequiresValidation = $True
        }

    }
    Else {
        $RowRequiresValidation = $True
    }

    Return $RowRequiresValidation

} # Function Test-RowRequiresValidation
Function Test-CellValidation {
    <#
    .SYNOPSIS
    Validates the value in the specified cell.

    .DESCRIPTION
    Tests whether the value of the specified cell passes validation, based on a validation regex defined in the schema and whether blanks are allowed.

    .PARAMETER Cell
    The cell to validate.

    .OUTPUTS
    $True if the cell value passes validation.
    $False if the cell value does not pass validation.

    .EXAMPLE
    Test-CellValidation -Cell $Cell
    #>

    Param (
        [System.Windows.Forms.DataGridViewCell]$Cell
    )

    # If the cell isn't set to read only.
    If (!($Cell.ReadOnly)) {

        # Get the name of the owning column.
        $OwningColumn = $Cell.OwningColumn.Name

        # If the cell has something in it.
        If (($Null -ne $Cell.EditedFormattedValue) -and ($Cell.EditedFormattedValue -ne "")) {

            # Get the validation RegEx for this column from the schema.
            $ValidationRegEx = $Schema.Columns.$OwningColumn.ValidationRegEx

            # If there's a validation RegEx specified in the schema.
            If ($ValidationRegEx) {
                # Evaluate validation status based on RegEx
                $ValidationPass = $Cell.EditedFormattedValue -Match $ValidationRegEx
            }
            Else {
                # No validation RegEx specified, default to true (pass)
                $ValidationPass = $True
            }
        }
        Else {
            # Cell is empty.
            
            # If we're allowing blanks for this column.
            If ($Null -ne $Schema.Columns.$OwningColumn.AllowBlank) {
                If (!($Schema.Columns.$OwningColumn.AllowBlank)) {
                    # Pass
                    $ValidationPass = $False
                }
                Else {
                    # Fail
                    $ValidationPass = $True
                }
            }
            Else {
                # Unspecified AllowBlank, so Pass
                $ValidationPass = $False
            }
        }

    }
    Else {
        # The cell is read only, so just return true.
        $ValidationPass = $True
    }

    # If we passed validation
    If ($ValidationPass) {
        # Reset errors and don't cancel.
        $Cell.ErrorText = $Null
    }
    Else {
        # Otherwise, set the error values and cancel.
        $ErrorText = $Schema.Columns.$OwningColumn.ErrorText
        If (!($ErrorText)) {
            $ErrorText = "The value you entered did not pass the cell validation rules."
        }

        $Cell.ErrorText = $ErrorText
    }

    Return $ValidationPass
} # Function Test-CellValidation
Function Test-RowValidation {
    <#
    .SYNOPSIS
    Validates the values in the specified Row.

    .DESCRIPTION
    Loops through the colums in the row and calls cell validation for each column.

    .PARAMETER RowIndex
    The index of the row for which to validate.

    .OUTPUTS
    $True if the row passes validation.
    $False if the row does not pass validation.

    .EXAMPLE
    Test-RowValidation -RowIndex 4
    #>

    Param (
        [int]$RowIndex
    )

    # Set the initial validation status for the row.
    $RowValidationPass = $True

    # Loop through each cell (column) in the row.
    ForEach ($Key In $Schema.Columns.Keys) {

        $Cell = $DataGridView.Rows[$RowIndex].Cells[$Key]

        # Run the cell validation.
        If (!(Test-CellValidation -Cell $Cell)) {
            # If we didn't pass validation, then set the validation status to false.
            $RowValidationPass = $False
        }
    }

    # Return the validation status for the row.
    Return $RowValidationPass

} # Function Test-RowValidation
Function Test-FileLock {
    <#
    .SYNOPSIS
    Test-FileLock - Checks whether the specified file is locked by another process.

    .DESCRIPTION
    For the specified file, attempts to get a handle. If successful, file is not locked. If not successful, file is locked.

    .PARAMETER Path
    The path of file to check for locks.

    .OUTPUTS
    A boolean value indicating whether the specified file is locked.

    .EXAMPLE
    Test-FileLock -Path "C:\Logs\logfile1.log"

    .NOTES
    Adapted from https://stackoverflow.com/a/24992975
    #>
        
    Param (
        [parameter(Mandatory = $True)][string]$Path
    )
    
    $oFile = New-Object System.IO.FileInfo $Path
        
    Try {
        $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    
        If ($oStream) {
            $oStream.Close()
        }
        Return $False
    }
    Catch {
        # File is locked by a process.
        Return $True
    }
} # Function Test-FileLock
Function Export-DataTable {
    <#
    .SYNOPSIS
    Exports the data table to CSV.

    .DESCRIPTION
    Forces validation on the data table. If it passes, export the Data Table to CSV.
        
    .OUTPUTS
    A CSV file.

    .EXAMPLE
    Export-DataTable
    #>

    # We need to end editing, just in case we're currently editing a row. This pushes the changes back to the datatable.
    [void] $DataGridView.EndEdit()

    # Set the focus somewhere else to trigger the row validation. This sticks an error against the rows / columns.
    [void] $TableLayoutPanel.Focus()

    # Determine whether the current datatable and the one since last save are the same.
    $Same = Compare-DataTable $DataTable $Script:OrigDataTable

    # If the data tables are NOT the same, OR if there's an error in the current row (meaning an uncommitted change)
    If (!($Same) -or ($DataGridView.CurrentRow.ErrorText)) {

        # We need to run validation, just in case were were editing a row when we attempted to close the form.
        $Validated = $True
        ForEach ($Row in $DataGridView.Rows) {
            If ($Row.ErrorText) {
                $Validated = $False
                Break
            }
        }

        If ($Validated) {
            $CancelAction = $False

            # Unlock the CSV file
            $Script:CSVFile.Close()
            $Script:CSVFile.Dispose()

            # Save changes
            $DataTable | Export-Csv -Path $CSVPath -NoTypeInformation

            # Lock the CSV file
            $Script:CSVFile = [System.IO.File]::Open($CSVPath, "Open", "ReadWrite", "None")

            # Take a copy of latest data for comparison later
            $Script:OrigDataTable.Dispose()
            $Script:OrigDataTable = $Null
            $Script:OrigDataTable = New-Object System.Data.DataTable
            $Script:OrigDataTable = $DataTable.Copy()

            [void] [System.Windows.Forms.MessageBox]::Show($Form, "Data saved.", "Data saved", "OK", "Info")
		
        }
        Else {
            # We didn't pass validation. Display an error message
            [void] [System.Windows.Forms.MessageBox]::Show($Form, "You'll have to correct the highlighted validation errors before you can save the data. `r`rAlternatively you can cancel the changes (click OK on this message, and then press ESC on your keyboard) and try again.", "Validation errors", "OK", "Warning")
            # Set the focus back to the DataGridView
            [void] $DataGridView.Focus()

            # And set the Cancel Action to True
            $CancelAction = $True
        }
    }
    Else {
        [void] [System.Windows.Forms.MessageBox]::Show($Form, "You haven't made any changes since the last save. `r`rThere's nothing to save!", "Data is unchanged", "OK", "Info")

        $CancelAction = $False
    }

    Return $CancelAction

} # Function Export-DataTable
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑ END OF FUNCTIONS ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓ START OF MAIN SCRIPT ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

Try {
    # If the HideConsole switch is specified, hide the console.
    If ($HideConsole) { Hide-Console }
    
    # Initialise
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # Check that the schema file exists
    Try {
        $SchemaPath = Resolve-Path -Path $SchemaPath -ErrorAction Stop
    }
    Catch {
        [void] [System.Windows.Forms.MessageBox]::Show($Form, "Couldn't find the specified Schema '$SchemaPath'.", "Schema not found", "OK", "Warning")
        Return
    }

    # Read in the schema
    $Schema = Get-FromJson -Path $SchemaPath
    If (!($Schema)) {
        [void] [System.Windows.Forms.MessageBox]::Show($Form, "Couldn't load the specified Schema '$SchemaPath'.`rCheck that the schema is a valid JSON file and is not malformed.", "Schema error", "OK", "Warning")
        Return
    }

    # Check that the CSV file exists.
    Try {
        $CSVPath = Resolve-Path -Path $CSVPath -ErrorAction Stop
    }
    Catch {
        [void] [System.Windows.Forms.MessageBox]::Show($Form, "Couldn't find the specified CSV '$CSVPath'.", "CSV not found", "OK", "Warning")
        Return
    }

    # Check whether the CSV file is locked.
    If (Test-FileLock -Path $CSVPath) {
        # File is locked. Display message and quit.
        [void] [System.Windows.Forms.MessageBox]::Show("The file '$CSVPath' is locked by another user, and can't be edited at this time.", "File is locked", "OK", "Warning")
        Return
    }

    # Check whether the file has any data
    If (!([String]::IsNullOrWhiteSpace((Get-Content $CSVPath)))) {
        # If it doesn't (i.e. it has some data)

        # Read in the CSV data.
        $CSVAll = Import-Csv $CSVPath

        # Check whether the columns in the CSV match those in the schema.
        If (!(Test-CSVColumns -CSVData $CSVAll -Schema $Schema)) {
            # Columns in CSV don't match the schema.
            [void] [System.Windows.Forms.MessageBox]::Show("The file '$CSVPath' doesn't match the schema defined in '$SchemaPath'!`r`rCan't continue.", "Schema mismatch", "OK", "Error")
            Return
        }
    }

    # Lock the file (assigned to a Script Scope variable). This is a static lock to prevent other people, or processes editing the file at the same time.
    $Script:CSVFile = [System.IO.File]::Open($CSVPath, "Open", "ReadWrite", "None")
    If ($ShowLockFile) {
        $LockFilePath = "$(Split-Path -Path $CSVPath -Resolve)\$(Split-Path -Path $CSVPath -Leaf -Resolve) IS LOCKED BY $([System.Environment]::UserName)" 
        New-Item -Path $LockFilePath -Force | Out-Null
    }

    ####################################################################################################
    # BEGIN BUILD DATATABLE                                                                            #
    #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#

    $DataTable = New-Object System.Data.DataTable

    # Get the headers from the CSV.
    $CSVHeaders = $CSVAll[0].PSObject.Properties | ForEach-Object { $_.Name }

    # First add the headers to the data table.
    ForEach ($Header in $CSVHeaders) {   
        [void] $DataTable.Columns.Add($Header)
    }

    # Then add the rows to the data table.
    ForEach ($CSVRow In $CSVAll) {
        $RowValues = ForEach ($Header in $CSVHeaders) {   
            $CSVRow.$Header
        }

        [void] $DataTable.Rows.Add($RowValues)
    }

    # Take a copy of original data for comparison later (to check whether the data has changed).
    $Script:OrigDataTable = New-Object System.Data.DataTable
    $Script:OrigDataTable = $DataTable.Copy()

    #↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#
    # END BUILD DATATABLE                                                                              #
    ####################################################################################################

    ####################################################################################################
    # BEGIN DEFINE FORM CONTROLS                                                                       #
    #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#

    # Set the form width and height to 60% of the available monitor size
    $FormWidth = ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width * 0.60)
    $FormHeight = ([System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height * 0.60)

    $Form = New-Object System.Windows.Forms.Form
    $Form.ClientSize = New-Object System.Drawing.Point($FormWidth, $FormHeight)

    # Get form title from schema
    If ($Null -ne $Schema.FormTitle) {
        If ($Schema.FormTitle) {
            $FormTitle = "$($Schema.FormTitle): $CSVPath"
        }
        Else {
            $FormTitle = "Data Entry: $CSVPath"
        }
    }
    Else {
        $FormTitle = "Data Entry: $CSVPath"
    }

    $Form.Text = $FormTitle
    $Form.TopMost = $False
    $Form.StartPosition = "CenterScreen"
    $Form.MinimumSize = New-Object System.Drawing.Point(800, 600) # Default minimum size. Can change if data to display is .. not very much.
    $Form.AutoValidate = "EnableAllowFocusChange" # This allows controls outside of the DataGridView to be clicked even if there's a validation failure in the current row / cell.

    # Form Icon
    # Icon from https://www.flaticon.com/authors/pancaza
    $IconBase64 = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAAdgAAAHYBTnsmCAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAKHSURBVDiNlZNLbIxRGIaf7///0ZnpbShKO5G4R2ajFggJFtKkFaHRCSENwliIuC00NtKEFbEkaUdEEWFBwmjVQkjLQtA2ISFpXHtJ3MbM1Eyn05n/s2inRjUR7/Kc93nfk3O+I/yHQo0h99zuvteudKrwlc+3ZNPh9Z9loqmwOrDIENmp6BqBWagkUfvQ1pPBDrs3crO+9fLGBV/6+FAyO/ZkToXPHCfXNVjFvgWnBbksYCHSpuhtQV7kzyh/Vjpv1YWU5dj8dL4PX/97yvrf5Tm/9+8fPYHfbxbFPTcNWK62URdta3yQzd33XN2RRKpD4izLrmWiMU5fOkXnwopzAlC0IXBClANIZmW05eK7XDhqD/egUsYPQCEynGYgntKpeY7d7bUFzVJQtXuGKdZ7FfbGWoLXs/Cuh+pMuod7VMRLeHIYwDANswYkHGsJ3shpdgxNSXWriHey5pcXj3YWVwXaAAyQFYI+GrWNKvI1dY2fLOY7YP+Gi53One21Bc22iQOh0us/4jLUZrqKfsnCdfcHZwrUYo+F5cCPt7iuAGTSI72AxIYSXgMhrCol2YBkekopgK1KOJlm4OefMIBYjhIAFf1hgXQJejC7+TWRKomMZDKqYlsGCU++40BHjetq7rAZGdYCHwdDTd9kWtVeb0bkrQ1bBluDdydO5l/y+83iuKcLuBdtDdYb4XsX+hTOGcj5/Or9s/7Fe+KeBqDcNuXM2CtATJ3HBf1kyUiHpzqwdDLQ6z/iKqoOnFU4BrptMNT0DWD8M5VW1uUnrbwmkG3AHZAWFe01BLdmWIGwAxDB2B5pbWwfv9CJTYVVe1YbYu4BXQeUAVHQN4LccpsSHAg1JXL9vwCZ0x9fLlzdKgAAAABJRU5ErkJggg=='
    $IconBytes = [Convert]::FromBase64String($IconBase64)
    # initialize a Memory stream holding the bytes
    $Stream = [System.IO.MemoryStream]::New($IconBytes, 0, $IconBytes.Length)
    $Form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

    # Save button
    $Button_Save = New-Object System.Windows.Forms.Button
    $Button_Save.Text = "Save"
    $Button_Save.Width = 100
    $Button_Save.Height = 30
    $Button_Save.Anchor = "Left,Bottom"
    $Button_Save.Font = New-Object System.Drawing.Font('Segoe UI', 10)

    # Quit button
    $Button_Quit = New-Object System.Windows.Forms.Button
    $Button_Quit.Text = "Quit"
    $Button_Quit.Width = 100
    $Button_Quit.Height = 30
    $Button_Quit.Anchor = "Left,Bottom"
    $Button_Quit.Font = New-Object System.Drawing.Font('Segoe UI', 10)

    # Data Grid View
    $DataGridView = New-Object System.Windows.Forms.DataGridView
    $DataGridView.Dock = "Fill"
    $DataGridView.ReadOnly = $False

    $DataGridView.MultiSelect = $False
    $DataGridView.AllowUserToDeleteRows = $True

    $DataGridView.DataSource = $DataTable
    $DataGridView.AutoGenerateColumns = $False
    $DataGridView.EnableHeadersVisualStyles = $False

    # Without VirtualMode = $True, the CancelRowEdit event doesn't fire.
    # https://docs.microsoft.com/en-us/dotnet/api/System.Windows.Forms.datagridview.cancelrowedit?view=windowsdesktop-6.0
    $DataGridView.VirtualMode = $True

    # https://generally.wordpress.com/2008/01/08/datagridviewcomboboxcolumn-requires-multiple-clicks-to-select-an-item/
    # https://web.archive.org/web/20080112002822/http://connect.microsoft.com/VisualStudio/feedback/ViewFeedback.aspx?FeedbackID=98504
    $DataGridView.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter

    $DataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $DataGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells

    # Context Menu
    # https://stackoverflow.com/questions/45574948/how-to-get-a-contextmenu-to-appear-on-right-click-of-a-datagridview-cell
    $ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip

    # Context Menu Option 1 (only option currently)
    $ToolStripItem1 = New-Object System.Windows.Forms.ToolStripMenuItem
    $ToolStripItem1.Text = "Delete Row"
    $ToolStripItem1.Name = "Delete Row"
    [void] $ContextMenuStrip.Items.Add($ToolStripItem1);

    # Column help
    $Panel_ColumnHelpTitle = New-Object System.Windows.Forms.Panel
    $Panel_ColumnHelpTitle.Dock = "Fill"

    $Label_ColumnHelpTitle = New-Object System.Windows.Forms.Label
    $Label_ColumnHelpTitle.Text = "Column Help:"
    $Label_ColumnHelpTitle.AutoSize = $True
    $Label_ColumnHelpTitle.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $Label_ColumnHelpTitle.Visible = $True

    $Panel_ColumnHelpText = New-Object System.Windows.Forms.Panel
    $Panel_ColumnHelpText.Dock = "Fill"

    $RichTextBox_ColumnHelpText = New-Object System.Windows.Forms.RichTextBox
    $RichTextBox_ColumnHelpText.Text = $Null
    $RichTextBox_ColumnHelpText.AutoSize = $True
    $RichTextBox_ColumnHelpText.Dock = "Fill"
    $RichTextBox_ColumnHelpText.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $RichTextBox_ColumnHelpText.Visible = $False
    $RichTextBox_ColumnHelpText.ReadOnly = $True
    $RichTextBox_ColumnHelpText.DetectUrls = $True

    #↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#
    # END DEFINE FORM CONTROLS                                                                         #
    ####################################################################################################

    ####################################################################################################
    # BEGIN ADD COLUMNS TO DATAGRIDVIEW                                                                #
    #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
    # Initialise hashtable to hold data tables for combo box options.
    $ComboBoxDataTables = @{}

    ForEach ($Key In $Schema.Columns.Keys) {

        Switch ( $Schema.Columns.$Key.Type ) {
            "ComboBox" {
                $Column = New-Object System.Windows.Forms.DataGridViewComboBoxColumn

                # Add a datatable to hold the combobox values to the ComboBoxDataTables Hash.
                $ComboBoxDataTables += @{$Key = New-Object System.Data.DataTable }

                [void] $ComboBoxDataTables.$Key.Columns.Add($Key)

                ForEach ($ComboBoxOption In $Schema.Columns.$Key.ComboBoxOptions) {
                    [void] $ComboBoxDataTables.$Key.Rows.Add($ComboBoxOption)
                }

                $Column.Name = $Key
                $Column.HeaderText = $Key
                $Column.DataSource = $ComboBoxDataTables.$Key
                $Column.ValueMember = $Key
                $Column.DisplayMember = $Key
                $Column.DataPropertyName = $Key
                $Column.FlatStyle = "Flat" # https://docs.microsoft.com/en-us/dotnet/api/System.Windows.Forms.flatstyle?view=windowsdesktop-6.0
                $Column.DisplayStyleForCurrentCellOnly = $True # https://stackoverflow.com/questions/1107069/how-can-i-hide-the-drop-down-arrow-of-a-datagridviewcomboboxcolumn-like-visual-s
            }
            "TextBox" {
                $Column = New-Object System.Windows.Forms.DataGridViewTextboxColumn

            }
        }

        # If the schema specifies that this column should be read only.
        If ($Schema.Columns.$Key.ColumnReadOnly) {
            # Set the column to read only.
            $Column.ReadOnly = $Schema.Columns.$Key.ColumnReadOnly
            # Set the header cell and cell colour.
            $Column.HeaderCell.Style.BackColor = "DarkGray"
            $Column.DefaultCellStyle.BackColor = "LightGray"
        }

        # If column wrap is specified
        If ($Schema.Columns.$Key.Wrap) {
            $Column.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
            $Column.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
        }

        # Common properties
        $Column.ToolTipText = $Schema.Columns.$Key.ToolTipText
        $Column.HeaderText = $Key
        $Column.DataPropertyName = $Key
        $Column.Name = $Key

        # Add column
        [void] $DataGridView.Columns.Add($Column)

    }
    #↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#
    # END ADD COLUMNS TO DATAGRIDVIEW                                                                  #
    ####################################################################################################

    ####################################################################################################
    # BEGIN ADD CONTROLS TO FORM                                                                       #
    #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
    [void] $Panel_ColumnHelpTitle.Controls.AddRange(@($Label_ColumnHelpTitle))
    [void] $Panel_ColumnHelpText.Controls.AddRange(@($RichTextBox_ColumnHelpText))

    #https://docs.microsoft.com/en-us/dotnet/api/System.Windows.Forms.tablelayoutpanel?view=windowsdesktop-6.0
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayoutPanel.Dock = "Fill"
    $TableLayoutPanel.AutoSize = $True

    [void] $TableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35))) # Buttons
    [void] $TableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60))) # DataGridView
    [void] $TableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25))) # Column Help Label
    [void] $TableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40))) # Column Help Text

    # Top panel has 2 columns, to achieve this we nest another TableLayoutPanel.
    $TableLayoutPanelTopButtons = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayoutPanelTopButtons.Dock = "Fill"
    $TableLayoutPanelTopButtons.AutoSize = $True
    [void] $TableLayoutPanelTopButtons.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) # Buttons
    [void] $TableLayoutPanelTopButtons.Controls.Add($Button_Save, 0, 0)
    [void] $TableLayoutPanelTopButtons.Controls.Add($Button_Quit, 1, 0)

    # We have 1 column (column 0), and 4 rows (0 - 3)
    [void] $TableLayoutPanel.Controls.Add($TableLayoutPanelTopButtons, 0, 0)
    [void] $TableLayoutPanel.Controls.Add($DataGridView, 0, 1)
    [void] $TableLayoutPanel.Controls.Add($Panel_ColumnHelpTitle, 0, 2)
    [void] $TableLayoutPanel.Controls.Add($Panel_ColumnHelpText, 0, 3)
 
    [void] $Form.Controls.AddRange(@($TableLayoutPanel))
    #↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#
    # END ADD CONTROLS TO FORM                                                                         #
    ####################################################################################################

    ####################################################################################################
    # BEGIN FORM CONTROL EVENTS                                                                        #
    #↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓#
    # Form Events
    [void] $Form.Add_Load({ 
            Set-GridProperties 

            # This bit is all cosmetic, to make the form look as good as possible upon first open, whilst still respecting minimum form size.
            # Get the current width of all columns in the datagridview.
            $DataGridViewWidth = 0    
            ForEach ($Column in $DataGridView.Columns) {
                $DataGridViewWidth += $Column.GetPreferredWidth([System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells, $True)
            }
            # Add the width of the Row Headers. We seem to need to add 24 to this total to make it fit ...
            $DataGridViewWidth = $DataGridViewWidth + $DataGridView.RowHeadersWidth + 24

            # If the width of the DataGridView is less than that of the form.
            If ($DataGridViewWidth -lt $Form.Width) {
        
                # Calculate a proportional height to match the width.
                $DataGridViewHeight = ($DataGridViewWidth / $Form.Width) * $Form.Height

                # Set the height and width of the form to match the data grid view.
                # NOTE: This won't override the minimum form size.
                $Form.Width = $DataGridViewWidth
                $Form.Height = $DataGridViewHeight

                # If we've have a vertical scrollbar (maybe caused by the height change above).
                If ( ($DataGridView.ScrollBars -and [System.Windows.Forms.ScrollBars]::Vertical) -ne [System.Windows.Forms.ScrollBars]::None) {
                    # Adjust the form width again (add width of vertical scroll bar) to compensate.
                    $Form.Width += [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
                }
            }

            # Scroll to last row
            $DataGridView.FirstDisplayedScrollingRowIndex = $DataGridView.RowCount - 1
			
            # This is a bit hacky, but it makes sure that the window is made active and on top of other windows (https://stackoverflow.com/a/33130322).
            $Form.WindowState = "Minimized"
            #[void] $Form.Show()
            $Form.WindowState = "Normal"
            [void] $Form.Activate()

        })

    [void]$Form.Add_Closing({

            # Set the focus somewhere else to trigger the row validation. This sticks an error against the rows / columns.
            $TableLayoutPanel.Focus()

            # Determine whether the current datatable and the one since last save are the same.
            $Same = Compare-DataTable $DataTable $Script:OrigDataTable

            # If the data tables are NOT the same, OR if there's an error in the current row (meaning an uncommitted change)
            If (!($Same) -or ($DataGridView.CurrentRow.ErrorText)) {

                # Prompt for action.
                $Result = [System.Windows.Forms.MessageBox]::Show($Form, "Do you want to save your changes?", "Save changes?", "YesNoCancel", "Question")

                Switch ( $Result ) {
                    "Yes" {
                        $CancelAction = Export-DataTable
                    }
                    "No" {
                        # Don't cancel. i.e. DO close the form.
                        $CancelAction = $False
                    }
                    "Cancel" {
                        # Cancel, i.e. DON'T close the form.
                        $CancelAction = $True
                    }
                }

                # Set the cancel state to the derived cancel state.
                $_.Cancel = $CancelAction
            }

        })

    # DataGridView Events
    [void]$DataGridView.Add_Sorted({
            Set-GridProperties
        })

    [void]$DataGridView.Add_CellEnter({
            Set-HelpLabels -Cell $args[0].CurrentCell
        })

    [void]$DataGridView.Add_CellClick({
            Set-HelpLabels -Cell $args[0].CurrentCell
        })

    [void]$DataGridView.Add_CellDoubleClick({
            Set-HelpLabels -Cell $args[0].CurrentCell
        })

    [void]$DataGridView.Add_CellContentClick({
            Set-HelpLabels -Cell $args[0].CurrentCell
        })

    [void]$DataGridView.Add_CellContentDoubleClick({
            Set-HelpLabels -Cell $args[0].CurrentCell
        })

    
    [void]$DataGridView.Add_EditingControlShowing({
            <# TODO: Add this functionality.
        # TODO: Don't want this for all columns, configure in the JSON. Might be useful for email addresses and stuff.
        #https://stackoverflow.com/questions/67165097/datagridview-with-datatable-as-source-and-suggestappend-combobox-columns-and-dyn
        # Not entirely sure how this works but it works lol
        # Basically the only column that can behave as WinForms ComboBox is
        # DataGridViewComboBoxColumn. When this is True we can enable DropDown and SuggestAppend
        # to the EditingControl. This code is emulated from C# and not sure if it's the right approach
        # but still it works.

        #https://docs.microsoft.com/en-us/dotnet/api/System.Windows.Forms.datagridview.editingcontrolshowing?view=windowsdesktop-6.0
        # [System.Windows.Forms.DataGridViewEditingControlShowingEventArgs]$e = $args[1]
        # $e is equivalent to $_

        if ($_.Control -as [System.Windows.Forms.ComboBox]) {   
            # Write-Host $_.Control.EditingControlDataGridView.CurrentCell.DisplayMember # <-- the column name
            $this.EditingControl.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
            $this.EditingControl.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
        }
        #>
        
            $_.CellStyle.BackColor = "white" # removes cell colour whilst editing.
        })

    [void]$DataGridView.Add_CellLeave({

            # Get the current cell.
            $Cell = $DataGridView.Rows[$_.RowIndex].Cells[$_.ColumnIndex]

            # If the cell isn't set to read only.
            If (!($Cell.ReadOnly)) {
                # This is a workaround. If we don't do this, if the user doesn't click on a value in the combobox column, then none is selected.
                # However, the Value passed to this function is NOT none, but the one highlighted but not actually clicked.
                # It's weird, however this sorts it by stamping the highlighted item in the cell as the value in the DataGridView cell.
                If ($Cell.OwningColumn.CellType -eq [System.Windows.Forms.DataGridViewComboBoxCell]) {
                    $DataGridView.Rows[$_.RowIndex].Cells[$_.ColumnIndex].Value = $Cell.EditedFormattedValue
                }
            }
        })

    [void]$DataGridView.Add_CancelRowEdit({
            # NOTE: this only fires if the data table is set to virtual mode.

            # Clear the error text for each cell in the row.
            ForEach ($Key In $Schema.Columns.Keys) {
                $Cell = $DataGridView.CurrentRow.Cells[$Key]
                $Cell.ErrorText = $Null
            }

            # Clear the error text for the row, and deselect the current cell.
            $DataGridView.CurrentRow.ErrorText = $Null
            $DataGridView.CurrentCell.Selected = $False

        })

    [void]$DataGridView.Add_RowLeave({
            Test-RowRequiresValidation -RowIndex $_.RowIndex
        })

    [void]$DataGridView.Add_RowValidating({

            # If this is a new row, and the row doesn't contain any data.
            If (!(Test-RowRequiresValidation -RowIndex $_.RowIndex)) {
                # Then we don't need to validate the data in the row.
                $RowValidationRequired = $False
            }
            Else {
                # This row is either not a new row, or it's a new row but contains data.
                $RowValidationRequired = $True
            }

            # If row validation is required.
            If ($RowValidationRequired) {
                # Run row validation.
                If (!(Test-RowValidation -RowIndex $_.RowIndex)) {
                    # If the row validation didn't pass, then add the row error text, and cancel the operation.
                    $DataGridView.Rows[$_.RowIndex].ErrorText = "One or more values in the row did not pass validation!"
                    $_.Cancel = $True
                }
                Else {
                    # Otherwise, we passed row validation. Clear the error text and allow the operation.
                    $DataGridView.Rows[$_.RowIndex].ErrorText = $Null
                    $_.Cancel = $False
                }
            }
        })

    [void]$DataGridView.Add_CellValidating({

            # Get the current cell.
            $Cell = $DataGridView.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
        
            # Run the cell validation.
            Test-CellValidation -Cell $Cell
        
        })

    [void]$DataGridView.Add_CellValidated({

            # Get the current cell.
            $Cell = $DataGridView.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
        
            # Run the cell transform.
            If ($Null -ne $Schema.Columns.$($Cell.OwningColumn.Name).Transform) {
                If (@("Title", "Upper", "Lower") -Contains $Schema.Columns.$($Cell.OwningColumn.Name).Transform) {
                    $Cell.Value = Get-FormattedValue -InputValue $Cell.Value -SetCase $Schema.Columns.$($Cell.OwningColumn.Name).Transform
                }
            }
        })

    [void]$DataGridView.Add_MouseDown({

            # Mouse down event on DataGridView and show menu when click
            #https://www.sapien.com/forums/viewtopic.php?t=11182

            If ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
                $RowIndex = ($DataGridView.HitTest($_.X, $_.Y)).RowIndex
                If ($DataGridView.Rows[$RowIndex].Selected -eq $False) {
                    $DataGridView.ClearSelection()
                    $DataGridView.Rows[$RowIndex].Selected = $True
                }
       
                If ($DataGridView.SelectedRows.Count -gt 1) {
                    $ContextMenuStrip.Items[$ToolStripItem1].Enabled = $False
                }
                Else {
                    $ContextMenuStrip.Items[$ToolStripItem1].Enabled = $True
                }
       
                $ContextMenuStrip.Show($DataGridView, $_.X, $_.Y)
            }
        })

    [void]$DataGridView.Add_DataError({
            # Error handling here
            # https://docs.microsoft.com/en-us/dotnet/api/System.Windows.Forms.datagridview.dataerror?view=net-5.0
            $_.Cancel = $True
        })

    # Context Menu Events
    [void] $ToolStripItem1.Add_Click({
            $RowIndex = $DataGridView.SelectedRows[0].Index
    
            If (!($DataGridView.Rows[$RowIndex].IsNewRow)) {
    
                $CellCount = 0
                ForEach ($Cell In $DataGridView.Rows[$RowIndex].Cells) {
                    If ($CellCount -le 5) {
                        $RowDetails += "$($Cell.OwningColumn.Name): $($Cell.Value)`r`n"
                    }
                    Else {
                        $RowDetails += "..."
                        Break
                    }
    
                    $CellCount++
                }
    
                $Result = [System.Windows.Forms.MessageBox]::Show($Form, "Are you sure you want to delete this row?`r`n`r`n$RowDetails`r`n`r`nYou can't undo this action." , "Confirm row deletion" , 4, 32)
                If ($Result -eq 'Yes') {
                    $DataGridView.Rows.RemoveAt($RowIndex)
                }
            }
        })
    

    # Column Help Text Events
    [void]$RichTextBox_ColumnHelpText.Add_LinkClicked({
            # Open the clicked link in the default browser for the system.
            Start-Process $_.LinkText
        })

    # Button events
    [void] $Button_Save.Add_Click({ Export-DataTable })
    [void] $Button_Quit.Add_Click({ $Form.Close() })
    #↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑#
    # END FORM CONTROL EVENTS                                                                          #
    ####################################################################################################

    # Show the form, and bring it to the front.
    [void] $Form.ShowDialog()
    [void] $Form.Activate()

}
Catch {
    [void] [System.Windows.Forms.MessageBox]::Show("Error at line $($_.InvocationInfo.ScriptLineNumber):`r$($_.Exception.Message)", "Error", "OK", "Warning")
}
Finally {
    # Clean up
    If ($DataTable) {
        [void] $DataTable.Dispose()
    }
    If ($ComboBoxDataTables) {
        ForEach ($ComboBoxTableKey in $ComboBoxDataTables.Keys) {
            [void] $ComboBoxDataTables.$ComboBoxTableKey.Dispose()
        }
    }

    # Unlock the CSV file
    If ($Script:CSVFile) {
        [void] $Script:CSVFile.Close()
        [void] $Script:CSVFile.Dispose()
    }

    If ($LockFilePath) {
        Remove-Item -Path $LockFilePath -Force -ErrorAction SilentlyContinue
    }

    # If the HideConsole switch was specified, now show the console.
    If ($HideConsole) { Show-Console }
}
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑ END OF MAIN SCRIPT ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
#↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑