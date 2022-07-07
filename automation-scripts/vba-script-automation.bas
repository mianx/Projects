Attribute VB_Name = "Module1"
Sub InsiderPivot()
Attribute InsiderPivot.VB_ProcData.VB_Invoke_Func = "I\n14"
'
' InsiderPivot Macro

'
Sheets("Insider").Select
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    'Sorting Data
    Range("A2:AD" & lastrow).Sort key1:=Range("P2:P" & lastrow), _
    order1:=xlDescending, Header:=xlNo
    Range("AD1").Select
'Computing Days Count
    ActiveCell.FormulaR1C1 = "Days Count"
    Dim last
    With ActiveSheet
        last = .Cells(.Rows.Count, "AC").End(xlUp).Row
    End With
    Range("AD2").Formula = "=IFERROR(TODAY() - P2,0)"
    Range("AD2").AutoFill Destination:=Range("AD2:AD" & last)
    
'Fill Data based on DaysCount
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "Periods"
    Range("AE2").Formula = "=IF(AD2<3,""Day""&AD2,IF(AD2<=7,""Wk1"",IF(AD2<=14,""Wk2"",IF(AD2<=21,""Wk3"",IF(AD2<=31,""Wk4"",IF(AD2<=61,""Mth2"",""Mth3""))))))"
    Range("AE2").AutoFill Destination:=Range("AE2:AE" & last)
    
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    
    Selection.FormulaR1C1 = "Example"
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A9"), Type:=xlFillDefault
    ActiveCell.Range("A1:A9").Select
    ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.FormulaR1C1 = "Promoters"
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A9"), Type:=xlFillDefault
    ActiveCell.Range("A1:A9").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.FormulaR1C1 = "Equity Shares"
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A9")
    ActiveCell.Range("A1:A9").Select
    ActiveCell.Offset(0, 6).Range("A1").Select
    Selection.FormulaR1C1 = "Buy"
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A9")
    ActiveCell.Range("A1:A9").Select
    
    ActiveCell.Offset(0, 7).Range("A1").Select
    Selection.FormulaR1C1 = "Market Purchase"
    
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A9"), Type:=xlFillDefault
    ActiveCell.Range("A1:A9").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mth3"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mth2"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Wk4"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Wk3"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Wk2"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Wk1"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Day2"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Day1"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.FormulaR1C1 = "Day0"
' Pivot Table
'Declare Variables
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim PCache As PivotCache
Dim PTable As PivotTable
Dim PRange As Range
Dim lastr As Long
Dim LastCol As Long

'Insert a New Blank Worksheet
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTable").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTable"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTable")
Set DSheet = Worksheets("Insider")

'Define Data Range
lastr = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(lastr, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), _
TableName:="InsiderPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="InsiderPivotTable")

    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "CATEGORY OF PERSON " & Chr(10) & "")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "TYPE OF SECURITY (PRIOR) " & Chr(10) & "")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "MODE OF ACQUISITION " & Chr(10) & "")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields("SYMBOL " & Chr(10) & "")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields("Periods")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("InsiderPivotTable").AddDataField ActiveSheet. _
        PivotTables("InsiderPivotTable").PivotFields( _
        "VALUE OF SECURITY (ACQUIRED/DISPLOSED) " & Chr(10) & ""), _
        "Count of VALUE OF SECURITY (ACQUIRED/DISPLOSED) " & Chr(10) & "", xlCount
    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "Count of VALUE OF SECURITY (ACQUIRED/DISPLOSED) " & Chr(10) & "")
        .Caption = "Sum of VALUE OF SECURITY (ACQUIRED/DISPLOSED) "
        .Function = xlSum
        .NumberFormat = "#\.00, "
    End With
    ActiveWindow.SmallScroll Down:=3
    ActiveSheet.PivotTables("InsiderPivotTable").AddDataField ActiveSheet. _
        PivotTables("InsiderPivotTable").PivotFields( _
        "NO. OF SECURITIES (ACQUIRED/DISPLOSED) " & Chr(10) & ""), _
        "Count of NO. OF SECURITIES (ACQUIRED/DISPLOSED) " & Chr(10) & "", xlCount
    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "Count of NO. OF SECURITIES (ACQUIRED/DISPLOSED) " & Chr(10) & "")
        .Caption = "Sum of NO. OF SECURITIES (ACQUIRED/DISPLOSED) "
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("InsiderPivotTable").DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
'Applying Pivot Filters
Sheets("PivotTable").Select
    ActiveSheet.PivotTables("InsiderPivotTable").PivotFields("CATEGORY OF PERSON " & Chr(10) & "" _
        ).CurrentPage = "(All)"
    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "CATEGORY OF PERSON " & Chr(10) & "")
        .PivotItems("-").Visible = False
        .PivotItems("Director").Visible = False
        .PivotItems("Employees/Designated Employees").Visible = False
        .PivotItems("Immediate relative").Visible = False
        .PivotItems("Key Managerial Personnel").Visible = False
        .PivotItems("Other").Visible = False
    End With
    ActiveSheet.PivotTables("InsiderPivotTable").PivotFields("CATEGORY OF PERSON " & Chr(10) & "" _
        ).EnableMultiplePageItems = True
    ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "TYPE OF SECURITY (PRIOR) " & Chr(10) & "").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "TYPE OF SECURITY (PRIOR) " & Chr(10) & "")
        .PivotItems("ADR/GDR/FCCB").Visible = False
        .PivotItems("Convertible Debenture").Visible = False
        .PivotItems("Debentures").Visible = False
        .PivotItems("Preference Shares").Visible = False
        .PivotItems("Warrants").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "TYPE OF SECURITY (PRIOR) " & Chr(10) & "").EnableMultiplePageItems = True
    ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "MODE OF ACQUISITION " & Chr(10) & "").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "MODE OF ACQUISITION " & Chr(10) & "")
        .PivotItems("-").Visible = False
        .PivotItems("Conversion of security").Visible = False
        .PivotItems("ESOP").Visible = False
        .PivotItems("Gift").Visible = False
        .PivotItems("Inter-se-Transfer").Visible = False
        .PivotItems("Invocation of pledge").Visible = False
        .PivotItems("Market Sale").Visible = False
        .PivotItems("Off Market").Visible = False
        .PivotItems("Others").Visible = False
        .PivotItems("Pledge Creation").Visible = False
        .PivotItems("Preferential Offer").Visible = False
        .PivotItems("Public Right").Visible = False
        .PivotItems("Revokation of Pledge").Visible = False
        .PivotItems("Scheme of Amalgamation/Merger/Demerger/Arrangement").Visible _
        = False
    End With
    ActiveSheet.PivotTables("InsiderPivotTable").PivotFields( _
        "MODE OF ACQUISITION " & Chr(10) & "").EnableMultiplePageItems = True
End Sub
Sub SummayMacro()
Attribute SummayMacro.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' SummayMacro
'

'SAST DATE COLUMN
    Sheets("SAST").Select
    Dim lr As Long
    lr = Cells(Rows.Count, 9).End(xlUp).Row
    Range("J1").Select
    Selection.Formula = "Date"
    Range("J2").Select
    ActiveCell.Formula = "=LEFT(I2,11)"
    Selection.AutoFill Destination:=Range("J2:J" & lr)
'   Copying Data from Pivot Table
'   Copy Symbol
    Sheets("PivotTable").Select
    Range("A8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Summary").Select
    Range("B2").Select
    ActiveSheet.Paste
    Sheets("Summary").Select
'Inserting Company Name
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("A2").Select
    Range("A2").Formula = "=+VLOOKUP(B2,Master!A:B,2,0)"
    Range("A2").AutoFill Destination:=Range("A2:A" & lastrow)
    
   
'Data From Pivot Table
    Range("C2").Select
    ActiveCell.Formula = "=IFERROR(GETPIVOTDATA(""Sum of VALUE OF SECURITY (ACQUIRED/DISPLOSED) "",PivotTable!R5C1,""SYMBOL " & Chr(10) & """,Summary!RC2,""Periods"",Summary!R1C)/100000,0)"
    Selection.AutoFill Destination:=Range("C2:K2"), Type:=xlFillDefault
    Range("C2:K2").Select
    Selection.AutoFill Destination:=Range("C2:K" & lastrow)
    Selection.NumberFormat = "_ * #,##0.0_ ;_ * -#,##0.0_ ;_ * ""-""??_ ;_ @_ "
    Selection.ColumnWidth = 9
    
'Computing Total
    Range("L2").Select
    ActiveCell.Formula = "=SUM(C2:K2)"
    Selection.AutoFill Destination:=Range("L2:L" & lastrow)
    Range("C2:L" & lastrow).Style = "Comma"         'Applying Comma style to columns
    Selection.ColumnWidth = 11.5
'Computing Avg Price
    Range("N2").Select
    ActiveCell.Formula = "=L2/PivotTable!U8*100000"
    Selection.AutoFill Destination:=Range("N2:N" & lastrow)
    Range("N2:N" & lastrow).Style = "Comma"
    Selection.NumberFormat = "0.00"
    
'CMP-Lookup  CLOSE_PRICE from the Price sheet
    Range("O2").Select
    ActiveCell.Formula = "=VLOOKUP(B2,Price!A:O,9)"
    Selection.AutoFill Destination:=Range("O2:O" & lastrow)
    
'Difference
    Range("P2").Select
    ActiveCell.Formula = "=IFERROR(+O2/N2-1,0)"
    Selection.AutoFill Destination:=Range("P2:P" & lastrow)
    Range("P2:P" & lastrow).Style = "Percent"
    

'Pledge% - Lookup (%) PLEDGE / DEMAT from Pledge sheet
    Range("S2").Select
    ActiveCell.Formula = "=IFERROR(VLOOKUP(A2,Pledge!A:M,12),0)/100"
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Selection.AutoFill Destination:=Range("S2:S" & lastrow)
    
'Sast Vlookup
    Range("T2").Select
    ActiveCell.Formula = "=IFERROR(VLOOKUP(B2,SAST!A:F,6,0)/100000,0)"
    Selection.AutoFill Destination:=Range("T2:T" & lastrow)
    Range("T2:T" & lastrow).Style = "Comma"
       
'Sast Date
    Range("U2").Select
    ActiveCell.Formula = "=IFERROR(VLOOKUP(B2,SAST!A:J,10,0),"""")"
    Selection.AutoFill Destination:=Range("U2:U" & lastrow)
    Selection.ColumnWidth = 11.4
'Prom % - Lookup(Column B) from Promoter% sheet
    Range("R2").Select
    ActiveCell.Formula = "=IFERROR(VLOOKUP(A2,'Promoter%'!A:I,2)/100,0)"
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Selection.AutoFill Destination:=Range("R2:R" & lastrow)
    
'Draging Formula Dynamically
    Range("V2").Select
    ActiveCell.Formula = "=(SUMIFS('Insider'!K:K,'Insider'!A:A,B2,'Insider'!L:L,""Sell"",'Insider'!E:E,""Promoters"")+SUMIFS('Insider'!K:K,'Insider'!A:A,B2,'Insider'!L:L,""Sell"",'Insider'!E:E,""Promoter Group""))/100000"
    Selection.NumberFormat = "0.00"
    Selection.AutoFill Destination:=Range("V2:V" & lastrow)
     Range("V2:V" & lastrow).Style = "Comma"
    
    Range("W2").Select
    ActiveCell.Formula = "=SUMIFS('Insider'!K:K,'Insider'!A:A,B2,'Insider'!L:L,""Sell"")/100000"
    Selection.NumberFormat = "0.00"
    Selection.AutoFill Destination:=Range("W2:W" & lastrow)
    Range("W2:W" & lastrow).Style = "Comma"
    
'Deleting End Row copied from PivotTable (GrandTotal)
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Delete Shift:=xlUp
    
'Applying filter
    Sheets("Summary").Select
    Columns("A:AAA").Select
    Selection.AutoFilter
    Range("B2").Select
        
' Adjustment
    Columns("A").ColumnWidth = 5
'Freezing Pane
    ActiveSheet.Range("C2").Select
    ActiveWindow.FreezePanes = True

'
End Sub
