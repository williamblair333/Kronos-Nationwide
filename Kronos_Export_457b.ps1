<# 
File:       kronos_export_457b.ps1
Date:       2021APR15
Author:     William Blair
Contact:    williamblair333@gmail.com
Note:       Runs from any folder but folder names with spaces will error if double-clicking to run


#>

<#
This script will do the following:
- save the kronos export xls as an xlsm
- import vba code into new xlsm file and create a module
- delete the first 7 rows
- merge 3 DEF  columns into one column, delete old columns, format as currency, add title "Record DEF"
- merge 2 Roth columns into one column, delete old columns, format as currency, add title "Record Roth"
- clear contents of remaining rows after last entry
- save the kronos export xlsm file into a csv

#>

<# 
Some links for review
Excel Saveas different formats 	https://stackoverflow.com/questions/6972494/how-save-excel-2007-in-html-format-using-powershell
Allow macros in excel 			https://stackoverflow.com/questions/35846996/running-excel-macro-from-windows-powershell-script 
Disable pop-ups					https://stackoverflow.com/questions/37979128/prevent-overwrite-pop-up-when-writing-into-excel-using-a-powershell
Run macros in powershell		https://www.excell-en.com/blog/2018/8/20/powershell-run-macros-copy-files-do-cool-stuff-with-power
Clean up user defined variables http://blog.insidemicrosoft.com/2017/05/28/how-to-clean-up-powershell-script-variables-without-restarting-powershell-ise/

#>

# Delete leftover xlsm and csv files
Add-Type -AssemblyName PresentationFramework 
#[System.Windows.MessageBox]::Show("Warning! Excel is going to close!  Please save all work before you click OK.")

Remove-Item * -Include *.xlsm*, *csv

# These registry keys will allow macro security - version 16
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\excel\Security" -Name AccessVBOM -PropertyType DWORD  -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\excel\Security" -Name VBAWarnings -PropertyType DWORD  -Value 1 -Force | Out-Null

# Kill all Excel processes
Stop-Process -Name "Excel"

# Get all files with xls extension
#Get-ChildItem $PSScriptRoot -Filter *.xls | 
Get-ChildItem $dir -File | Where-Object { $_.Extension -eq ".xls" } |

# Run this For loop for a files with xls extension
Foreach-Object {
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host

$excelFile = $_.FullName

# Variables to set file format to saveas
$formatXLSM = 52 	#xlOpenXMLWorkbookMacroEnabled
$formatCSV = 6 		#xlCSV

# Filler for the saveas process
$missing = [type]::Missing

# Cycle through each file with .xls extension in the script's directory

# Create excel object 
$excel = New-Object -ComObject Excel.Application

# disable visible updating of sheet
$excel.Visible = $false

# If you want to see what's going on while script runs, uncomment visible and displayalerts below
#$excel.Visible = $true
#$excel.DisplayAlerts = $true

$workBook = $excel.Workbooks.Open($excelFile)

# disables the pop up asking if it's ok to overwrite - just overwrite it
$excel.DisplayAlerts = $false;

# saveas xlsm
$excelFile = $excelFile + 'm'  
$excel.ActiveWorkbook.SaveAs($excelFile,$formatXLSM,$missing,$missing,$missing,$missing,$missing,$missing,$missing,$missing,$missing,$missing)

$excel.Quit()

# Create excel object 
$excel = New-Object -ComObject Excel.Application

Write-Host "Now processing file: " $excelFile
$workBook = $excel.Workbooks.Open($excelFile)

$excelModule = $workBook.VBProject.VBComponents.Add(1)

# This was supposed to load up the entire file as the macro from... the macro but I couldn't get it working
#$macroImport = [IO.File]::ReadAllText("R:\Nationwide\Nationwide_Export_Prep_Powershell.bas")

# This is the VBA Code 
$excelMacro = @"
'**************************************************************************

    Sub Nationwide_Export_Prep()
    
    'Reset the sheet.  Used for debugging and testing
    'Call Report_Restore
    
    Application.ScreenUpdating = False
    'Application.ScreenUpdating = True
    Sheets("report").Select
    
    Rows("1:6").Select
    Selection.Delete Shift:=xlUp
    
    'Testing files that already had a g column before it got removed 
	'Record DEF 2 457 Amount
	'Call Col_Del_G
    
    'Record Roth 457(b) Amount
    'Create this blank column concatenate columns
    'Record DEF 2, Record DEF 457 Amount, Record DEFCU 457 Amount
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("I2").Select
    'ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-2],RC[-1])"
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-2],RC[-1])"
    ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
    
   'While ActiveCell.Offset(rowOffset:=0, columnOffset:=-3) <> ""
   While ActiveCell.Offset(rowOffset:=0, columnOffset:=-3) <> ""
        'ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-3],RC[-2],RC[-1])"
        ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-2],RC[-1])"
        ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
   Wend
        
    'Create this blank column concatenate columns
    'Record Roth 457(b), Record Roth 457 CU Amount
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-2],RC[-1])"
    ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
    
   While ActiveCell.Offset(rowOffset:=0, columnOffset:=-3) <> ""
        ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-2],RC[-1])"
        ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
   Wend

    'Select Column Record Roth 457(b)
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Columns("I:I").Select
    Selection.Copy
    
    Columns("J:J").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Select Columns Record DEF 2 457 Amount, Record DEF 457 Amount, Record DEFCU 457 Amount
    Columns("G:I").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("G:G").Select
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "#,##0.00"
    
    Columns("G:G").Select
    Selection.NumberFormat = "$#,##0.00"

    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Columns("J:J").Select
    Selection.Copy
    
    Columns("K:K").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Columns("H:J").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    Columns("H:H").Select
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "$#,##0.00"
    
    Columns("H:H").Select
    Selection.NumberFormat = "$#,##0.00"
    
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Record DEF"
    
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Record Roth"
    
    Dim rownumber As Integer
    Dim rowstring As String


    Range("A7").Select

    While ActiveCell <> ""
        rownumber = ActiveCell.Row

        If ActiveCell <> "" Then
    
            ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
        Else
           
        End If

    Wend

    rownumber = rownumber + 1
    rowstring = CStr(rownumber) & ":" & CStr(Rows.Count)

    Rows(rowstring).ClearContents
    
End Sub
'*********************************************************************************

Sub Report_Restore()

    Application.DisplayAlerts = False
    Sheets("report").Select
    ActiveWindow.SelectedSheets.Delete
    
    Sheets("report_copy").Select
    Sheets("report_copy").Name = "report"
    
    Sheets("report").Select
    Sheets("report").Copy After:=Sheets(1)
    Sheets("report (2)").Select
    Sheets("report (2)").Name = "report_copy"
    Sheets("report").Select
    Range("A1").Select
    
End Sub
'*********************************************************************************
'*********************************************************************************

Sub Col_Del_G()

' Col_Del Macro
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
End Sub
'*********************************************************************************
"@

# This adds the macro to the xlsm file	
$excelModule.CodeModule.AddFromString($excelMacro)

# This runs the macro 
$excel.Run("Nationwide_Export_Prep")

# saveas csv
    $excelFile = $_.FullName
	$excelFile = [io.path]::GetFileNameWithoutExtension("$excelFile")
	$excelFile = $excelFile + ".csv"
	$excelFile = $PSScriptRoot + "\" + $excelFile
$excel.ActiveWorkbook.SaveAs($excelFile,$formatCSV,$missing,$missing,$missing,$missing,$missing,$missing,$missing,$missing,$missing,$missing)

$excel.Quit()

Stop-Process -Name "Excel" 
}

# These registry keys will disable macro security - version 16
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\excel\Security" -Name AccessVBOM -PropertyType DWORD  -Value 0 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\excel\Security" -Name VBAWarnings -PropertyType DWORD  -Value 0 -Force | Out-Null