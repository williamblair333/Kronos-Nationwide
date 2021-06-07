Attribute VB_Name = "Module2"

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
    
    If LaborLevel = "Firefighter" Or LaborLevel = "EMS" Then
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

