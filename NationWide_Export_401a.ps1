<# 
File:       kronos-nationwide-401a-prep.ps1
Date:       2021MAY27
Author:     William Blair
Contact:    williamblair333@gmail.com
Note:       Runs from any folder but folder names with spaces will error if double-clicking to run

#>

<#
This script will do the following:
- save the kronos 401a xls as an xlsm
- import vba code into new xlsm file and create a module
- delete the first 7 rows
- merge rows that have the same SS# (Should only be two in this environment)
	- add data (money) contained in either Column G or H (Only one of them) \
	  with data (money) found in either Column I, J, or K (Only one of them).
	  
- Column G & H totals needs to have this formula applied. =MIN(Column*100%,125)
- clear contents of remaining rows after last entry
- merge two remaining columns together
- save the kronos export xlsm file into a csv

#>

<# 
Some links for review:
Excel Saveas different formats 	https://stackoverflow.com/questions/6972494/how-save-excel-2007-in-html-format-using-powershell
Allow macros in excel 			https://stackoverflow.com/questions/35846996/running-excel-macro-from-windows-powershell-script 
Disable pop-ups					https://stackoverflow.com/questions/37979128/prevent-overwrite-pop-up-when-writing-into-excel-using-a-powershell
Run macros in powershell		https://www.excell-en.com/blog/2018/8/20/powershell-run-macros-copy-files-do-cool-stuff-with-power
Clean up user defined variables http://blog.insidemicrosoft.com/2017/05/28/how-to-clean-up-powershell-script-variables-without-restarting-powershell-ise/
Variables overview				https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-variables#public-statement
Merging columns (didn't use)	https://stackoverflow.com/questions/28916523/vba-macro-merging-two-columns-into-1


#>

<#
TODO:
Refine the VBA so we can manipulate variables and cells without using .Select or .Activate 
								https://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba/10717999#10717999

#>

# Delete leftover xlsm and csv files
Add-Type -AssemblyName PresentationFramework 
[System.Windows.MessageBox]::Show("Warning! Excel is going to close!  Please save all work before you click OK.")

Remove-Item * -Include *.xlsm, *csv

# These registry keys will allow macro security - version 16
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\excel\Security" -Name AccessVBOM -PropertyType DWORD  -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\excel\Security" -Name VBAWarnings -PropertyType DWORD  -Value 1 -Force | Out-Null

# Kill all Excel processes
Stop-Process -Name "Excel"

# Get all files with xls extension
Get-ChildItem $PSScriptRoot -Filter *.xls | 

# Run this For loop for files with xls extension
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

$workBook = $excel.Workbooks.Open($excelFile)

# disables the pop up asking if it's ok to overwrite - just overwrite it
$excel.DisplayAlerts = $false;

# saveas xlsm
$excelFile = $excelFile + 'm'  
$excel.ActiveWorkbook.SaveAs($excelFile,$formatXLSM,$missing,$missing,$missing,$missing,$missing,$missing,$missing,$missing,$missing,$missing)

$excel.Quit()

# Create excel object 
$excel = New-Object -ComObject Excel.Application

# If you want to see what's going on while script runs, uncomment visible and displayalerts below
#$excel.Visible = $true
#$excel.DisplayAlerts = $true

Write-Host "Now processing file: " $excelFile
$workBook = $excel.Workbooks.Open($excelFile)

$excelModule = $workBook.VBProject.VBComponents.Add(1)

# This was supposed to load up the entire file as the macro from... the macro but I couldn't get it working
#$macroImport = [IO.File]::ReadAllText("R:\Nationwide\Nationwide_Export_Prep_Powershell.bas")

# This is the VBA Code 
$excelMacro = @"

'**************************************************************************

    Public Sheet As String
    
    'Strings to create the min formula
    Public MinHead As String
    Public MinTailG As String
    Public MinTailH As String
    
    'Concatenated strings for formula
    Public FunMinCVG As String
    Public FunMinFFH As String
        
    'LaborLevel determines which formula
    'to use
    Public LaborLevel As String
    
    'Used to compare dup SS# entries
    Public cellValue1 As Variant
    Public cellValue2 As Variant
    
    'All G through K columns togeter
    'and dup rows too
    Public AddUp As Currency

    'G through K 1st and 2nd row
    'values
    Public GOneValue As Currency
    Public HOneValue As Currency
    Public IOneValue As Currency
    Public JOneValue As Currency
    Public KOneValue As Currency
    
    Public GTwoValue As Currency
    Public HTwoValue As Currency
    Public ITwoValue As Currency
    Public JTwoValue As Currency
    Public KTwoValue As Currency

'**************************************************************************
    
Sub Nationwide_401a_Prep()
    
    Sheets("report").Select
    
    'Setting up formula variables

    MinHead = "=MIN("
    MinTailG = "*100%,125)"
    MinTailH = "*50%,125)"
    
    'Reset the sheet.  Used for debugging and testing
    'Call Report_Restore

    Application.ScreenUpdating = False
    
    'Clear out dashes in preparation of adding up cells G through K
    Range("G:K").Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    'Delete Header text
    Rows("1:7").Delete Shift:=x1Up
    
    'We start everything from B2 now
    Range("B2").Activate
    
    ' Loop through the rows and quit if empty cell
    While ActiveCell <> ""
        
        Call XVariable_Reset
        Call XRow1_Count
        Call XRow_Duplicate_Check
        
    Wend
    
    Call XConcatColumns
    Call XClean_Up
    
    
End Sub

'**************************************************************************

Sub XVariable_Reset()
        AddUp = 0
        GOneValue = 0
        HOneValue = 0
        IOneValue = 0
        JOneValue = 0
        KOneValue = 0
    
        GTwoValue = 0
        HTwoValue = 0
        ITwoValue = 0
        JTwoValue = 0
        KTwoValue = 0
        
        FunMinCVG = ""
        FunMinFFH = ""

End Sub

'**************************************************************************

Sub XGet_LaborLevel()
    
    LaborLevel = ActiveCell.Offset(rowOffset:=0, columnOffset:=18).Value
    
End Sub

'**************************************************************************

Sub XRow1_Count()

    'Go to G column (Record DEF 2 457 Amount, CV)
    ActiveCell.Offset(rowOffset:=0, columnOffset:=5).Select
    GOneValue = ActiveCell.Value
    AddUp = AddUp + ActiveCell.Value
    ActiveCell.Clear
    
    'Go to H column (Record DEF 457 Amount, FF)
    ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Activate
    HOneValue = ActiveCell.Value
    AddUp = AddUp + ActiveCell.Value
    ActiveCell.Clear
    
    'Go to I column (Record DEFCU 457 Amount)
    ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Activate
    IOneValue = ActiveCell.Value
    AddUp = AddUp + ActiveCell.Value
    ActiveCell.Clear
            
    'Go to J column (Record Roth 457(b)
    ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Activate
    JOneValue = ActiveCell.Value
    AddUp = AddUp + ActiveCell.Value
    ActiveCell.Clear
            
    'Go to K column (Record Roth 457 CU)
    ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Activate
    KOneValue = ActiveCell.Value
    AddUp = AddUp + ActiveCell.Value
    ActiveCell.Clear

End Sub

'**************************************************************************

Sub XRow_Duplicate_Check()

    ActiveCell.Offset(rowOffset:=0, columnOffset:=-9).Select
    cellValue1 = ActiveCell.Value
    ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Select
    cellValue2 = ActiveCell.Value
    
    If cellValue2 = cellValue1 Then
        Call XRow_Duplicate_Merge
    
    Else
        ActiveCell.Offset(rowOffset:=-1, columnOffset:=5).Select
        Call XRow_Formula_Totals
    
    End If

End Sub

'**************************************************************************

Sub XRow_Duplicate_Merge()
    
    ActiveCell.Offset(rowOffset:=0, columnOffset:=5).Select
    GTwoValue = ActiveCell.Value
    AddUp = AddUp + ActiveCell.Value
    ActiveCell.Clear
            
    ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Select
    HTwoValue = ActiveCell.Value
    AddUp = AddUp + ActiveCell.Value
    ActiveCell.Clear
    
    ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Select
    ITwoValue = ActiveCell.Value
    AddUp = AddUp + ActiveCell.Value
    ActiveCell.Clear
    
    ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Select
    JTwoValue = ActiveCell.Value
    AddUp = AddUp + ActiveCell.Value
    ActiveCell.Clear

    ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Select
    KTwoValue = ActiveCell.Value
    Rows(ActiveCell.Row).Delete
    
    ActiveCell.Offset(rowOffset:=-1, columnOffset:=-4).Select
    
    Call XRow_Formula_Totals


End Sub

'**************************************************************************

Sub XRow_Formula_Totals()

    FunMinCVG = MinHead & AddUp & MinTailG
    FunMinFFH = MinHead & AddUp & MinTailH
    
    LaborLevel = ActiveCell.Offset(rowOffset:=0, columnOffset:=13).Value
    
    If LaborLevel = "Firefighter" Or LaborLevel = "EMS" Or LaborLevel = "Reserve" Then
        ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Activate
        ActiveCell.Value = AddUp
        ActiveCell.Formula = FunMinFFH
        ActiveCell.Offset(rowOffset:=1, columnOffset:=-6).Activate
                
    Else
        ActiveCell.Value = AddUp
        ActiveCell.Formula = FunMinCVG
        ActiveCell.Offset(rowOffset:=1, columnOffset:=-5).Activate
            
    End If

End Sub

'**************************************************************************

Sub XConcatColumns()

    Call XGo_Home
        
        Do Until ActiveCell.Offset(rowOffset:=0, columnOffset:=5).Value = Empty And ActiveCell.Offset(rowOffset:=0, columnOffset:=6).Value = Empty
            
            Call XVariable_Reset
             
            'Go to G Column (Record DEF 2 457 Amount)
            ActiveCell.Offset(rowOffset:=0, columnOffset:=5).Select
            AddUp = ActiveCell.Value
            ActiveCell.Clear
            
            'Go to H column (Record DEF 457 Amount, FF)
            ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Activate
            AddUp = AddUp + ActiveCell.Value
            ActiveCell.Value = AddUp
            
            ActiveCell.Offset(rowOffset:=1, columnOffset:=-6).Select
    
        Loop
    
    'Delete summed totals at bottom in prep for csv export
    Call XRow_Delete
    Call XRow_Delete
    Call XRow_Delete
    
    Range("G:G,I:I,J:J,K:K").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Record Totals"


End Sub

'**************************************************************************

Sub XRow_Delete()

    Rows(ActiveCell.Row).EntireRow.Delete
    
End Sub

'**************************************************************************

Sub XClean_Up()
    
    With Selection.Font
        .Name = "Arial"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
        
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
        
   Call XGo_Home
   
End Sub

'**************************************************************************

Sub XGo_Home()

    Range("B2").Select

End Sub

'**************************************************************************

Sub XCell_Address()
    
    ActiveCell.Offset(rowOffset:=0, columnOffset:=0).Select

    cellAddress = ActiveCell.Offset(0, 0).Address(False, False)
    cellValue1 = ActiveCell.Value
 
    MsgBox cellAddress & ", " & cellValue1

End Sub

'**************************************************************************

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
"@

# This adds the macro to the xlsm file	
$excelModule.CodeModule.AddFromString($excelMacro)

# This runs the macro 
$excel.Run("Nationwide_401a_Prep")

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

