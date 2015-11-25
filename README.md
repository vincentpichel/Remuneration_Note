# Remuneration_Note

Sub ClearWorksheets()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

On Error Resume Next

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Dim sh3 As Worksheet
Dim sh4 As Worksheet
Dim wkb As Workbook
Dim ws As Worksheet

Set wkb = ActiveWorkbook

With wkb

    Set sh1 = .Sheets("RP_ACTUL")
    Set sh2 = .Sheets("Supern Detail")
    Set sh3 = .Sheets("TaxData")
    Set sh4 = .Sheets("Allowances")


   sh1.Tab.Color = 9145219
   sh2.Tab.Color = 9145219
   sh3.Tab.Color = 9145219
   sh4.Tab.Color = 9145219
    
End With

Call AutoFilter_Off

Call Delete_SUPER_Dups_Sheets

Call Delete_DATACHK_Sheet


    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Note" Then
        
            If ws.Name <> "Status Codes" Then
            
                If ws.Name <> "Summary & Check" Then
                
                    If ws.Name <> "Audit Check" Then
            
                    ws.AutoFilterMode = False
                    ws.Range("A1:AW" & ws.Rows.Count).Clear
                End If
              End If
            
            End If
        End If
    Next ws

Sheets("Exec Cars").Activate

    With Worksheets("Exec Cars")
    
        .Range("A1").FormulaR1C1 = "Business"
        .Range("B1").FormulaR1C1 = "Employee Number"
        .Range("C1").FormulaR1C1 = "Car/Car Park Amount"
        .Range("F2").FormulaR1C1 = "* Data added on this tab is required to be in the following format:" & Chr(10) & "Column A = Business, Column B = Employee #, Column C = Amount"
    
        .Range("A1:C2").Font.Bold = True
        .Range("A1:C2").HorizontalAlignment = xlCenter
        .Range("A1:C2").VerticalAlignment = xlTop
        .Range("A1:C2").WrapText = True
      
      With Range("A1:C2").Interior
      
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10157978
        
      End With
       
      
    End With
    
Sheets("RP_ACTUL").Activate
    
With ActiveWorkbook.Sheets("RP_ACTUL").Tab
        .Color = 13487553 'Gray
        .TintAndShade = 0
    End With
    
Sheets("Supern Detail").Activate
    
With ActiveWorkbook.Sheets("Supern Detail").Tab
        .Color = 13487553 'Gray
        .TintAndShade = 0
    End With
    
Sheets("TaxData").Activate
    
With ActiveWorkbook.Sheets("TaxData").Tab
        .Color = 13487553 'Gray
        .TintAndShade = 0
    End With
    
Sheets("Allowances").Activate

With ActiveWorkbook.Sheets("Allowances").Tab
        .Color = 13487553 'Gray
        .TintAndShade = 0
    End With


Sheets("Summary & Check").Activate

Call Clear_Command_Button_Colours

 
   
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
   
    
End Sub

Private Sub Clear_Command_Button_Colours()

Dim cb As CommandButton
Dim i As Long

For i = 1 To 9
Set cb = Sheet8.Shapes("CommandButton" & i).OLEFormat.Object.Object
    With Sheet8.Range("B3").Cells(i, 1)
        If cb.BackColor = 65280 Then
            cb.BackColor = 12632256
         End If
    End With
Next i

End Sub
Private Sub AutoFilter_Off()

Application.ScreenUpdating = False
'Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.AutoFilterMode = False
        
    Next ws
    
Application.ScreenUpdating = True
'Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub Choose_ACTUL_File()

Application.ScreenUpdating = False
'Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Sheets("RP_ACTUL").Activate

    With Worksheets("RP_ACTUL")
    
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
 
        .AutoFilterMode = False
        .Range("A1:AZ" & LastRow).Clear
    
    End With

Dim Ret
    
    Ret = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")

    If Ret <> False Then
 With Worksheets("RP_ACTUL").QueryTables.Add(Connection:= _
        "FINDER;" & Ret, _
         Destination:=Range("$A$1"))
        .Name = "RawDataACTUL"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False

        End With
    End If
    
Call delete_Connections

Call ACTUL_Format_Update

With ActiveWorkbook.Sheets("RP_ACTUL").Tab
        .Color = 49407 'Orange
        .TintAndShade = 0
    End With

Sheets("Summary & Check").Activate


Application.ScreenUpdating = True
'Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

    
End Sub
Sub Choose_SUPER_File()

Application.ScreenUpdating = False
'Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Sheets("Supern detail").Activate

    With Worksheets("Supern detail")
    
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
 
        .AutoFilterMode = False
        .Range("A1:AZ" & LastRow).Clear
    
    End With

Dim Ret
   
    Ret = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")

    If Ret <> False Then
 With Worksheets("Supern detail").QueryTables.Add(Connection:= _
        "FINDER;" & Ret, _
         Destination:=Range("$A$1"))
        .Name = "RawDataSUP"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
        '.TextfileColumndatatypes =
        

        End With
    End If
    
Call delete_Connections

Call SUPER_Text_to_Col

With ActiveWorkbook.Sheets("Supern Detail").Tab
        .Color = 49407 'Orange
        .TintAndShade = 0
    End With

Sheets("Summary & Check").Activate

Application.ScreenUpdating = True
'Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

    
End Sub

Sub Choose_ALLOWANCES_File()

'OBIEE Import file nees to be updated so Employee ID comes out in Column A

Application.ScreenUpdating = False
'Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Sheets("Allowances").Activate

With Worksheets("Allowances")
    
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
 
        .AutoFilterMode = False
        .Range("A1:AZ" & LastRow).Clear
    
    End With

Dim Ret
   
    Ret = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")

    If Ret <> False Then
 With Worksheets("Allowances").QueryTables.Add(Connection:= _
        "FINDER;" & Ret, _
         Destination:=Range("$A$1"))
        .Name = "RawDataALLOW"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False

        End With
    End If
    
Call delete_Connections

Call ALLOW_Text_to_Col

With ActiveWorkbook.Sheets("Allowances").Tab
        .Color = 49407 'Orange
        .TintAndShade = 0
    End With

Sheets("Summary & Check").Activate

Application.ScreenUpdating = True
'Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

    
End Sub

Sub Choose_TAX_File()

Application.ScreenUpdating = False
'Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Sheets("TaxData").Activate

With Worksheets("TaxData")
    
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
 
        .AutoFilterMode = False
        .Range("A1:AZ" & LastRow).Clear
    
    End With

Dim Ret
   
    Ret = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")

    If Ret <> False Then
 With Worksheets("TaxData").QueryTables.Add(Connection:= _
        "FINDER;" & Ret, _
         Destination:=Range("$A$1"))
        .Name = "RawDataTAX"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False

        End With
    End If
    
Call delete_Connections

Call TAX_Text_to_Col

With ActiveWorkbook.Sheets("TaxData").Tab
        .Color = 49407 'Orange
        .TintAndShade = 0
    End With

Sheets("Summary & Check").Activate

Application.ScreenUpdating = True
'Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

    
End Sub
Sub FilterAndCopy()

Dim LastRow As Long
Dim vName As Variant
Dim rngName As Range
Set rngName = Sheets("Status Codes").Range("NameList")
vName = rngName.Value

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Sheets("Contracted employees").Activate

With Worksheets("Contracted employees")

    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
        
        .AutoFilterMode = False
        .Range("A4:AD" & LastRow).Clear

End With

'Call ACTUL_Format_Update (Don't think this is needed as it's already called at the time the
'file is uploaded

With Worksheets("RP_ACTUL")

     .AutoFilterMode = False
    
    ' Add code to ensure column headings are correct and in Row 1
    '.Range("A3").EntireRow.Delete
       

    'Array filter from NameList
    .Range("A3:AD3").AutoFilter Field:=12, Criteria1:=Application.Transpose(vName), _
                                Operator:=xlFilterValues
      

    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    .Range("A1:AD" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Contracted employees").Range("A1")
                    

End With

Sheets("Summary & Check").Activate

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic



End Sub

Private Sub ACTUL_Format_Update()

'look into stating each For with format change - see if that speeds thing up or rather just have on For
'that encompasses entire sheet (i.e. ActiveSheet.Columns("A:AD"))

'Create Similar Sub for Superannuation data

Sheets("RP_ACTUL").Activate

Dim LastRow As Long
Dim c As Range

With Worksheets("RP_ACTUL")

LastRow = .Range("A" & .Rows.Count).End(xlUp).Row

    .Range("A4:A" & LastRow).NumberFormat = "dd/mm/yyyy"

    .Range("J4:J" & LastRow).NumberFormat = "dd/mm/yyyy"

    .Range("K4:K" & LastRow).NumberFormat = "dd/mm/yyyy"
    
    .Range("L4:L" & LastRow).NumberFormat = "dd/mm/yyyy"
    
    .Range("N4:N" & LastRow).NumberFormat = "dd/mm/yyyy"
    
    .Range("O4:O" & LastRow).NumberFormat = "dd/mm/yyyy"
    
    .Range("S4:U" & LastRow).NumberFormat = "$#,##0.00;[Red]-$#,##0.00"
    
    .Range("AC4:AC" & LastRow).NumberFormat = "dd/mm/yyyy"
    
For Each c In Application.Intersect(ActiveSheet.UsedRange, ActiveSheet.Columns("A"))
  c.Value = c.Value
  
Next
    
For Each c In Application.Intersect(ActiveSheet.UsedRange, ActiveSheet.Columns("J"))
  c.Value = c.Value
  
Next
  
For Each c In Application.Intersect(ActiveSheet.UsedRange, ActiveSheet.Columns("K"))
  c.Value = c.Value
  
Next
    
For Each c In Application.Intersect(ActiveSheet.UsedRange, ActiveSheet.Columns("L"))
  c.Value = c.Value
  
Next

For Each c In Application.Intersect(ActiveSheet.UsedRange, ActiveSheet.Columns("N"))
  c.Value = c.Value
  
Next

For Each c In Application.Intersect(ActiveSheet.UsedRange, ActiveSheet.Columns("O"))
  c.Value = c.Value
  
Next

For Each c In Application.Intersect(ActiveSheet.UsedRange, ActiveSheet.Columns("S:U"))
  c.Value = c.Value
  
Next

For Each c In Application.Intersect(ActiveSheet.UsedRange, ActiveSheet.Columns("AC"))
  c.Value = c.Value
  
Next


End With

'Application.ScreenUpdating = True

End Sub

Sub Data_Integ_Check()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
'Application.EnableEvents = False

On Error Resume Next

    Test1 = Sheets("Contracted Employees").Cells(Rows.Count, "A").End(xlUp).Row
    Test2 = Sheets("RP_ACTUL").Cells(Rows.Count, "A").End(xlUp).Row

    If Test2 < 1 Then
    
        MsgBox "Not able to proceed, you first need to run the 'Upload RP_ACTUL' macro."
        
            End
    
        Else
        
            If Test1 < 1 Then
        
                MsgBox "Not able to proceed, you first need to run the 'Filter & Copy Data' macro."
                
                End
                
                Else
                
            End If
    End If

Call Delete_DATACHK_Sheet

ActiveWorkbook.Worksheets.Add(Before:=Sheet3).Name = "Data Checks"

Sheets("Data Checks").Select

    With ActiveWorkbook.Sheets("Data Checks").Tab
        .Color = 255
        .TintAndShade = 0
    End With

Sheets("RP_ACTUL").Activate

   'this part activates an input box requiring the user to enter the first day of Current financial year
    
    '  "01/Jul/2013" - Use for texting purposes  (i.e. bypass pop up message box)
    A1 = InputBox("Enter First Day of Current Fiancial Year (eg. 01/Jul/2014)", "DD/MMM/YYYY")
    B1 = DateAdd("yyyy", 5, A1)
   
   ' this part is a logic check to ensure a valid response is entered in the input box, if it is not the
    ' do / loop repeats until a valid response is provided

    Do Until A1 <> ""

    If A1 = "" Then
    
        A1 = InputBox("Enter First Day of Current Fiancial Year (eg. 01-JUL-2014)", "DD/MMM/YYYY")
    
    End If
    
    Loop

 
With Worksheets("RP_ACTUL")
    
    .AutoFilterMode = False
    
    Range("A4").Select
    
LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    
    Range("A4:AD" & LastRow).Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    Range("A3:AD3").AutoFilter Field:=12, Criteria1:="="
    
    Range("L4:L" & LastRow).SpecialCells(xlCellTypeVisible).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535 'Yellow
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    .Range("A1:AD" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Data Checks").Range("A1")
            
    Selection.AutoFilter Field:=12

End With

Sheets("Data Checks").Activate

    'Think about entering a comment for further explanation
    Range("A3").FormulaR1C1 = "The below entries do not have a Status Code entered (see yellow shaded cells) - They HAVE NOT been included in the Contracted Employees Tab"
    Range("A3").Font.Bold = True
    
        With Range("A3").Characters(Start:=86, Length:=8).Font
            .Underline = xlUnderlineStyleDouble
        End With
        
    Range("A3:AD3").Activate
    
        With Range("A3:AD3").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    
        With Range("A3:AD3").BorderAround([xlContinuous], _
            [xlThin], [xlColorIndexAutomatic])

        End With


Sheets("Contracted Employees").Activate

With Worksheets("Contracted Employees")

LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
LastRowA = Sheets("Data Checks").Cells(Rows.Count, "A").End(xlUp).Row
    
    .AutoFilterMode = False
    
    Range("A4:AD" & LastRow).Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    Range("A3:AV3").AutoFilter Field:=20, Criteria1:= _
        "=$0.00", Operator:=xlOr, Criteria2:="="
        
    
    Range("T4:T" & LastRow).SpecialCells(xlCellTypeVisible).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 49407 'Orange
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
 LastRowChk = Range("T4:T" & LastRow).SpecialCells(xlCellTypeVisible).Count
       
    If LastRowChk > 0 Then
    
    .Range("A4:AD" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Data Checks").Range("A" & LastRowA + 2)
            
        Else: End
        
     End If
    
    Selection.AutoFilter Field:=20

Sheets("Data Checks").Activate

    'Think about entering a comment for further explanation
    Range("A" & 1 + LastRowA).FormulaR1C1 = "The below entries do not have Salary Amount entered (see orange shaded cells)  - They HAVE been included in the Contracted Employees Tab"
    Range("A" & 1 + LastRowA).Font.Bold = True

        With Range("A" & 1 + LastRowA).Characters(Start:=87, Length:=4).Font
            .Underline = xlUnderlineStyleDouble
        End With
        
    Range("A" & 1 + LastRowA & ":AD" & 1 + LastRowA).Activate
    
        With Range("A" & 1 + LastRowA & ":AD" & 1 + LastRowA).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    
        With Range("A" & 1 + LastRowA & ":AD" & 1 + LastRowA).BorderAround([xlContinuous], _
            [xlThin], [xlColorIndexAutomatic])

        End With


Sheets("Contracted Employees").Activate

LastRowB = Sheets("Data Checks").Cells(Rows.Count, "A").End(xlUp).Row

    Range("O4").AutoFilter Field:=15, Criteria1:= _
        "<" & A1, Operator:=xlOr, Criteria2:="="
        
    
    Range("O4:O" & LastRow).SpecialCells(xlCellTypeVisible).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.399975585192419 'Red
            .PatternTintAndShade = 0
        End With
    
LastRowChk = Range("O4:O" & LastRow).SpecialCells(xlCellTypeVisible).Count
       
    If LastRowChk > 0 Then
 
    .Range("A4:AD" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Data Checks").Range("A" & LastRowB + 2)
            
        Else: End
        
    End If

    Selection.AutoFilter Field:=15

Sheets("Data Checks").Activate

    'Think about entering a comment for further explanation
    Range("A" & 1 + LastRowB).FormulaR1C1 = "The below entries do not have no contract end date or one that is pre " & A1 & " entered (see pink shaded cells)  - They HAVE been included in the Contracted Employees Tab"
    Range("A" & 1 + LastRowB).Font.Bold = True

        With Range("A" & 1 + LastRowB).Characters(Start:=123, Length:=4).Font
          .Underline = xlUnderlineStyleDouble
        End With
        
    Range("A" & 1 + LastRowB & ":AD" & 1 + LastRowB).Activate
    
        With Range("A" & 1 + LastRowB & ":AD" & 1 + LastRowB).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    
        With Range("A" & 1 + LastRowB & ":AD" & 1 + LastRowB).BorderAround([xlContinuous], _
            [xlThin], [xlColorIndexAutomatic])

        End With


Sheets("Contracted Employees").Activate

LastRowC = Sheets("Data Checks").Cells(Rows.Count, "A").End(xlUp).Row

    Range("O4").AutoFilter Field:=15, Criteria1:= _
            ">" & B1, Operator:=xlAnd
        
    
    Range("O4:O" & LastRow).SpecialCells(xlCellTypeVisible).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419 'Purple
            .PatternTintAndShade = 0
        End With
    
   
LastRowChk = Range("O4:O" & LastRow).SpecialCells(xlCellTypeVisible).Count
       
    If LastRowChk > 0 Then
        
    .Range("A4:AD" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Data Checks").Range("A" & LastRowC + 2)
            
      Else: End
        
    End If
    
    Selection.AutoFilter Field:=15
     
End With
          
Application.Sheets("Data Checks").Activate

    'Think about entering a comment for further explanation
    Range("A" & 1 + LastRowC).FormulaR1C1 = "The below entries have a contract end date greater than 5 years (i.e. beyond " & B1 & ") entered (see purple shaded cells)  - They HAVE been included in the Contracted Employees Tab"
    Range("A" & 1 + LastRowC).Font.Bold = True

        With Range("A" & 1 + LastRowC).Characters(Start:=132, Length:=4).Font
            .Underline = xlUnderlineStyleDouble
        End With
        
        Range("A" & 1 + LastRowC & ":AD" & 1 + LastRowC).Activate
    
        With Range("A" & 1 + LastRowC & ":AD" & 1 + LastRowC).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    
        With Range("A" & 1 + LastRowC & ":AD" & 1 + LastRowC).BorderAround([xlContinuous], _
            [xlThin], [xlColorIndexAutomatic])

        End With
        
        Range("A1").Select

    ActiveWindow.Zoom = 75

    Columns("A:A").ColumnWidth = 13
    Columns("B:AD").EntireColumn.AutoFit
    
Sheets("Summary & Check").Activate
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
'Application.EnableEvents = True

End Sub

Sub Super_Dups()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic

Test1 = Sheets("Supern Detail").Cells(Rows.Count, "A").End(xlUp).Row

    If Test1 < 1 Then
    
        MsgBox "Not able to proceed, you first need to run the 'Upload SUPER Data' macro."
    
            End
        
        Else
    
    End If

Call Delete_SUPER_Dups_Sheets

ActiveWorkbook.Worksheets.Add(Before:=Sheet9).Name = "Duplicates Part 1"

    With ActiveWorkbook.Sheets("Duplicates Part 1").Tab
            .Color = 9129488 'Blue
            .TintAndShade = 0
        End With
ActiveWorkbook.Worksheets.Add(Before:=Sheet9).Name = "Duplicates Part 2"

    With ActiveWorkbook.Sheets("Duplicates Part 2").Tab
            .Color = 9129488 'Blue
            .TintAndShade = 0
        End With
ActiveWorkbook.Worksheets.Add(Before:=Sheet9).Name = "Duplicates Part 3"

    With ActiveWorkbook.Sheets("Duplicates Part 3").Tab
            .Color = 9129488 'Blue
            .TintAndShade = 0
        End With
ActiveWorkbook.Worksheets.Add(After:=Sheet2).Name = "Super Cleansed"

    With ActiveWorkbook.Sheets("Super Cleansed").Tab
            .Color = 65535 'Yellow
            .TintAndShade = 0
        End With

Sheets("Supern detail").Activate

With Worksheets("Supern detail")

    .AutoFilterMode = False

    Columns("M:M").Clear

    Range("M1").FormulaR1C1 = "Count Emp Super"
    Range("M4").FormulaR1C1 = "=COUNTIFS(R4C3:R25000C3,RC[-10])"
    
'Copy formula and paste to final row of data
        
    Range("M4").Select
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    .Range("M4").SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Supern detail").Range("M5:M" & LastRow)
            
'Paste values from formulas

           
    .Range("M4:M" & LastRow).SpecialCells(xlCellTypeFormulas, 23).Copy
    .Range("M4:M" & LastRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    .Range("A3:M3").AutoFilter Field:=13, Criteria1:=">1"
    
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    .Range("A1:M" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Duplicates Part 1").Range("A1")
  End With
  
Sheets("Duplicates Part 1").Activate
  
With Worksheets("Duplicates Part 1")

Call SUPER_Text_to_Col
  
  
    Range("N1").FormulaR1C1 = "Current Sup Com Amount"
    Range("N4").FormulaR1C1 = "=IF(RC[-3]>0,0,RC[-5])"
    
     LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    .Range("N4").SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Duplicates Part 1").Range("N5:N" & LastRow)
            
    Columns("N:N").Copy
    Columns("N:N").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            
    Range("A3:N3").AutoFilter Field:=14, Criteria1:="<>"
    
    Range("A:N").SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Duplicates Part 2").Range("A1")

End With

Sheets("Duplicates Part 2").Activate

With Worksheets("Duplicates Part 2")

    Range("O1").FormulaR1C1 = "2nd Count Emp Super"
    Range("O4").FormulaR1C1 = "=COUNTIFS(R4C3:R10000C3,RC[-12])"
    
     LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    .Range("O4").SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Duplicates Part 2").Range("O5:O" & LastRow)
            
    Columns("O:O").Copy
    Columns("O:O").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            
    Range("A3:O3").AutoFilter Field:=15, Criteria1:=">1"
    
    Range("A:O").SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Duplicates Part 3").Range("A1")
  End With
  
Sheets("Duplicates Part 3").Activate

With Worksheets("Duplicates Part 3")

 .Rows("3:3").Delete Shift:=xlUp

    Cells.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(9, 14), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
        
    LastRow = .Range("I" & .Rows.Count).End(xlUp).Row
    .Range("A" & LastRow & ":O" & LastRow).Clear
    
    ActiveSheet.Outline.ShowLevels RowLevels:=2
       
    .Range("C:C").SpecialCells(xlCellTypeVisible).ClearContents
    
    ActiveSheet.Outline.ShowLevels RowLevels:=3
    
    .Range("C1").FormulaR1C1 = "Employee"
    
    LastRow = .Range("I" & .Rows.Count).End(xlUp).Row
    .Range("A1:H" & LastRow + 1).SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    
    End With
    
Call Combine_SUPER_Sheets

Sheets("Summary & Check").Activate

Application.ScreenUpdating = True

End Sub

Private Sub Delete_SUPER_Dups_Sheets()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets
         If ws.Name Like "Duplicates Part" & "*" Then
         
            '~~> This check is required to ensure that you don't get an error
            '~~> if there is only one sheet left and it matches the delete criteria
             If ThisWorkbook.Sheets.Count = 1 Then
                MsgBox "There is only one sheet left and you cannot delete it"
            Else
                '~~> This is required to supress the dialog box which excel shows
                '~~> When you delete a sheet. Remove it if you want to see the
                '~~~> Dialog Box
                Application.DisplayAlerts = False
                
                ws.Delete
                
                Application.DisplayAlerts = True
            End If
            
        End If
  Next
        
        For Each ws In ThisWorkbook.Sheets
         If ws.Name Like "Super Cleansed" Then
         
            '~~> This check is required to ensure that you don't get an error
            '~~> if there is only one sheet left and it matches the delete criteria
             If ThisWorkbook.Sheets.Count = 1 Then
                MsgBox "There is only one sheet left and you cannot delete it"
            Else
                '~~> This is required to supress the dialog box which excel shows
                '~~> When you delete a sheet. Remove it if you want to see the
                '~~~> Dialog Box
                Application.DisplayAlerts = False
                
                ws.Delete
                
                Application.DisplayAlerts = True
            End If
            
        End If
        
      
    Next
End Sub

Private Sub SUPER_Text_to_Col()
 
  Dim rng As Range
 
  For Each rng In Range("A:N").Columns
      
    If Application.CountA(rng) > 0 Then

        rng.Select
    
        Selection.TextToColumns Destination:=rng, DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Else
    
    End If
    
  Next rng
 
End Sub
Private Sub TAX_Text_to_Col()
 
  Dim rng As Range
 
  For Each rng In Range("A:E").Columns
      
    If Application.CountA(rng) > 0 Then

        rng.Select
    
        Selection.TextToColumns Destination:=rng, DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Else
    
    End If
    
  Next rng
 
End Sub
Private Sub ALLOW_Text_to_Col()
 
  Dim rng As Range
 
  For Each rng In Range("A:E").Columns
      
    If Application.CountA(rng) > 0 Then

        rng.Select
    
        Selection.TextToColumns Destination:=rng, DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Else
    
    End If
    
  Next rng
 
End Sub

Sub Combine_SUPER_Sheets()

Application.ScreenUpdating = False

Dim LastRow As Long
Dim LastRowA As Long

Sheets("Supern detail").Activate
 
With Worksheets("Supern detail")

    Range("M3").AutoFilter Field:=13, Criteria1:="1"
    
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    
    .Range("A1:M" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Super Cleansed").Range("A1")
    
End With

Sheets("Duplicates Part 2").Activate

With Worksheets("Duplicates Part 2")

    Range("O3").AutoFilter Field:=15, Criteria1:="1"
    
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    LastRowA = Sheets("Super Cleansed").Cells(Rows.Count, "A").End(xlUp).Row
    .Range("A4:O" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Super Cleansed").Range("A" & LastRowA + 1)
    
End With

Sheets("Duplicates Part 3").Activate

With Worksheets("Duplicates Part 3")

ActiveSheet.Outline.ShowLevels RowLevels:=2
     
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    LastRowA = Sheets("Super Cleansed").Cells(Rows.Count, "A").End(xlUp).Row
    .Range("A4:O" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Super Cleansed").Range("A" & LastRowA)
    
End With

End Sub
Sub Create_Formulas()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Test1 = Sheets("Contracted Employees").Cells(Rows.Count, "A").End(xlUp).Row

If Test1 < 1 Then
    
    MsgBox "Not able to proceed, you first need to run the 'Filter & Copy Data' macro"
        
            End
    
       Else
        
     Q1 = InputBox("Have all entries on 'Data Checks' sheet been" & Chr(10) & "investigated & any updates applied accordingly? (i.e. Yes/No)")
           
        If Q1 = "No" Then
           
                MsgBox "Please ensure all entries on 'Data Checks' sheet have been sorted before continuing."
                
                        End
                
                   Else
             
         End If
End If

' this part activates an input box requiring the user to enter the first day of Current financial year
    
    A1 = InputBox("Current Fiancial Year (eg. 2015)", "DD/MM/YYYY")
    B1 = (A1 + 1)
    C1 = (A1 - 1)
    
   
   ' this part is a logic check to ensure a valid response is entered in the input box, if it is not the
    ' do / loop repeats until a valid response is provided

       Do Until A1 <> ""

    If A1 = "" Then
    
        A1 = InputBox("Current Fiancial Year (eg. 2015)", "DD/MM/YYYY")
    
    End If
    
    Loop


Sheets("Contracted employees").Activate

With Worksheets("Contracted employees")

    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    
    .AutoFilterMode = False
    
    .Range("AE1:AV" & LastRow).Clear
    
    .Range("AE1") = "Super Fund" & Chr(10) & " Code"
    .Range("AF1") = "Super Fund" & Chr(10) & "Description"
    .Range("AG1") = "Super Fund" & Chr(10) & "Amt Type"
    .Range("AH1") = "Super Fund" & Chr(10) & "Amount"
    .Range("AI1") = "Annual Superannuation" & Chr(10) & "(employer super & sal sac super)"
    .Range("AJ1") = "Annual Allowances"
    .Range("AK1") = "Executives Car/Car Parking" & Chr(10) & "not in Salary Amount"
    .Range("AL1") = "Daily FBT Payable &" & Chr(10) & "Taxable Value"
    .Range("AM1") = "Total contract duration"
    .Range("AN1") = "Duration elapsed" & Chr(10) & "prior years'"
    .Range("AO1") = "Duration elapsed" & Chr(10) & "current year"
    .Range("AP1") = "Total Duration elapsed up to"
    .Range("AQ1") = "Days remaining"
    .Range("AR1") = "Year 1 Duration"
    .Range("AS1") = "Duration remaining" & Chr(10) & "(years 2 to 5)"
    .Range("AT1") = "Duration check"
    .Range("AU1") = "Year 1"
    .Range("AV1") = "Year 2 - 5"
    
    With Range("AE1:AV3").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599963377788629
        .PatternTintAndShade = 0
        
    End With
    
    .Range("AE1:AV1").Font.Bold = True
    .Range("AE1:AV3").HorizontalAlignment = xlCenter
    .Range("AE1:AV1").VerticalAlignment = xlTop
    .Range("AE1:AV1").WrapText = True
    
    .Range("AR2") = "06/30/" & B1 '2015
    .Range("AR3") = "07/01/" & A1 '2014
    .Range("AQ2") = "06/30/" & A1 '2014
    .Range("AP2") = "06/30/" & A1 '2014
    .Range("AN2") = "06/30/" & C1 '2013

    .Range("AE4:AE" & LastRow).FormulaR1C1 = "=VLOOKUP(RC[-25],'Super Cleansed'!R3C3:R50000C10,4,FALSE)"
    .Range("AF4:AF" & LastRow).FormulaR1C1 = "=VLOOKUP(RC[-26],'Super Cleansed'!R3C3:R50000C10,5,FALSE)"
    .Range("AG4:AG" & LastRow).FormulaR1C1 = "=VLOOKUP(RC[-27],'Super Cleansed'!R3C3:R50000C10,6,FALSE)"
    .Range("AH4:AH" & LastRow).FormulaR1C1 = "=VLOOKUP(RC[-28],'Super Cleansed'!R3C3:R50000C10,7,FALSE)"
    .Range("AI4:AI" & LastRow).FormulaR1C1 = "=IF(RC[-2]=""PERCENT"",(RC[-15]*(RC[-1]/100)),RC[-1]*26)"
    .Range("AJ4:AJ" & LastRow).FormulaR1C1 = "=IF(ISNA(VLOOKUP(RC[-30],'Allowances'!R2C1:R50000C2,2,FALSE)),0,(VLOOKUP(RC[-30],'Allowances'!R2C1:R50000C2,2,FALSE)))"
    .Range("AK4:AK" & LastRow).FormulaR1C1 = "=IF(ISNA(VLOOKUP(RC[-31],'Exec cars'!R3C2:R5000C3,2,FALSE)),0,(VLOOKUP(RC[-31],'Exec cars'!R3C2:R5000C3,2,FALSE)))"
    .Range("AL4:AL" & LastRow).FormulaR1C1 = "=IF(ISNA(VLOOKUP(RC[-32],'TaxData'!R2C2:R50000C6,5,FALSE)),0,(VLOOKUP(RC[-32],'TaxData'!R2C2:R50000C6,5,FALSE)))"
        .Range("AI4:AL" & LastRow).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    .Range("AM4:AM" & LastRow).FormulaR1C1 = "=RC[-24]-RC[-25]"
    .Range("AN4:AN" & LastRow).FormulaR1C1 = "=IF(RC[-26]<=R2C40,R2C40-RC[-26],0)"
    .Range("AO4:AO" & LastRow).FormulaR1C1 = "=RC[1]-RC[-1]"
    .Range("AP4:AP" & LastRow).FormulaR1C1 = "=IF(RC[-27]<=R2C42,RC[-3],R2C42-RC[-28])"
    .Range("AQ4:AQ" & LastRow).FormulaR1C1 = "=IF(RC[-28]<=R2C43,0,RC[-4]-RC[-1])"
    .Range("AR4:AR" & LastRow).FormulaR1C1 = "=IF(RC[-1]<=365,RC[-1],365)"
    .Range("AS4:AS" & LastRow).FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Range("AT4:AT" & LastRow).FormulaR1C1 = "=IF(RC[-2]+RC[-1]=RC[-3],""ok"",""error"")"
    .Range("AU4:AU" & LastRow).FormulaR1C1 = "=IF(RC[-3]=365,(RC[-27]+RC[-12]+RC[-11]+(RC[-9]*365)+RC[-10]),((RC[-27]/365*RC[-3])+(RC[-12]/365*RC[-3])+(RC[-11]/365*RC[-3])+(RC[-10]/365*RC[-3])+(RC[-9]*RC[-3])))"
    .Range("AV4:AV" & LastRow).FormulaR1C1 = "=IF(RC[-3]>0,((RC[-28]/365*RC[-3])+(RC[-13]/365*RC[-3])+(RC[-12]/365*RC[-3])+(RC[-11]/365*RC[-3])+(RC[-10]*RC[-3])),0)"
        .Range("AU4:AV" & LastRow).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"

 End With
   
Columns("AE:AV").EntireColumn.AutoFit

Call Final_Data_Intg_Chk

Call Create_Tax_Formulas

Sheets("Contracted Employees").Activate

With Worksheets("Contracted Employees")
        
    .AutoFilterMode = False
   
    .Range("AU" & 1 + LastRow).FormulaR1C1 = "=SUM(R4C:R[-1]C)"
    .Range("AV" & 1 + LastRow).FormulaR1C1 = "=SUM(R4C:R[-1]C)"
    
End With
    
Sheets("Note").Activate

With Worksheets("Note")

    LastRowA = Sheets("Contracted Employees").Cells(Rows.Count, "A").End(xlUp).Row
    LastRowB = LastRowA - 5
    LastRowC = LastRowB - 1
    
  .Range("I3").FormulaR1C1 = A1
  .Range("I6").FormulaR1C1 = "=ROUND('Contracted employees'!R[" & LastRowB & "]C[38]/1000,0)"
  .Range("I7").FormulaR1C1 = "=ROUND('Contracted employees'!R[" & LastRowC & "]C[39]/1000,0)"

End With

Call Generate_Audit_Check

Sheets("Summary & Check").Activate
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.EnableEvents = True


End Sub
Private Sub Create_Tax_Formulas()

'Can only be ran once RP_ACTUL data has been copied to "Contracted Employees" tab
'''This can only be ran via 'Create_Formulas' Sub and it contains a check to see whether RP_ACTUL has any data

Sheets("TaxData").Activate

With Worksheets("TaxData")

  LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
  LastRowA = LastRow + 1
     
    .AutoFilterMode = False
    
Application.Calculation = xlCalculationAutomatic
    
    .Range("E:F").Clear
    
    .Range("E1") = "Duration Elapsed " & Chr(10) & "Current Year"
    .Range("F1") = "DAILY TOTAL " & Chr(10) & "FBT PAYABLE & " & Chr(10) & "TAXABLE VALUE"
    
    With Range("E1:F1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599963377788629
        .PatternTintAndShade = 0
        
    End With
    
    .Range("E1:F1").Font.Bold = True
    .Range("E1:F1").HorizontalAlignment = xlCenter
    .Range("E1:F1").VerticalAlignment = xlTop
    .Range("E1:F1").WrapText = True
    
    .Range("E2:E" & LastRow).FormulaR1C1 = "=IF(ISNA(VLOOKUP(RC[-3],'Contracted employees'!R1C6:R10000C42,36,FALSE)),0,(VLOOKUP(RC[-3],'Contracted employees'!R1C6:R10000C42,36,FALSE)))"
    .Range("F2:F" & LastRow).FormulaR1C1 = "=IF(RC[-1]=0,0,RC[-2]/RC[-1])"
    
    .Range("E" & LastRowA).FormulaR1C1 = "=SUM(R[-" & LastRow & "]C:R[-1]C)"
    .Range("F" & LastRowA).FormulaR1C1 = "=SUM(R[-" & LastRow & "]C:R[-1]C)"
    
    .Range("D2:F" & LastRowA).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    
End With

Application.Calculation = xlCalculationManual

End Sub

Private Sub delete_Connections()

Do While ActiveWorkbook.Connections.Count > 0

ActiveWorkbook.Connections.Item(ActiveWorkbook.Connections.Count).Delete

Loop

End Sub

Sub Generate_Audit_Check()

'Required due to formulas created a pivot table based on them
Application.Calculation = xlCalculationAutomatic

Application.ScreenUpdating = False
Application.EnableEvents = False


'Generate Pivot Tables on sheets

Sheets("Allowances").Activate

With Worksheets("Allowances")

LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
LastRowPivot = .Range("G" & .Rows.Count).End(xlUp).Row

If LastRowPivot > 1 Then

    Columns("C:H").Delete
    
        Else
    
    End If

    .Range("C1").FormulaR1C1 = "In or Out"
    .Range("D1").FormulaR1C1 = "Variance" & Chr(10) & "of Included"
    
     
       With Range("C1:D1").Interior
        .Pattern = xlSolid
        .Color = 10156544
        .TintAndShade = 0.599963377788629
        .PatternTintAndShade = 0
        
    End With
    

    
    .Range("A1:D1").Font.Bold = True
    .Range("A1:D1").HorizontalAlignment = xlCenter
    .Range("A1:D1").VerticalAlignment = xlTop
    .Range("A1:D1").WrapText = True
    
   
    .Range("C2:C" & LastRow).FormulaR1C1 = "=IF(ISNA((VLOOKUP(RC[-2],'Contracted employees'!R4C6:R10000C6,1,FALSE)))=FALSE,""Included"",""Excluded"")"
    .Range("D2:D" & LastRow).FormulaR1C1 = "=IF(RC[-1]=""Included"",(RC[-2]-VLOOKUP(RC[-3],'Contracted employees'!R4C6:R10000C36,31,FALSE)),"""")"
        .Range("D2:D" & LastRow).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
        
    End With

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Allowances!R1C1:R50000C4", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="Allowances!R1C7", TableName:="PivotTable5", _
        DefaultVersion:=xlPivotTableVersion14
    Sheets("Allowances").Activate
    Cells(1, 7).Select
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("In or Out")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Entered DR")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Entered DR"), "Sum of Entered DR", xlSum
   
   Range("H2:H5").NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
   
Sheets("TaxData").Activate

With Worksheets("TaxData")

LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
LastRowPivot = .Range("J" & .Rows.Count).End(xlUp).Row

If LastRowPivot > 1 Then

    Columns("G:K").Delete
    
        Else
    
    End If

    .Range("G1").FormulaR1C1 = "In or Out"
    .Range("H1").FormulaR1C1 = "Amount"
    
     
       With Range("G1:H1").Interior
        .Pattern = xlSolid
        .Color = 10156544
        .TintAndShade = 0.599963377788629
        .PatternTintAndShade = 0
        
    End With
    
    
    .Range("G1:H1").Font.Bold = True
    .Range("G1:H1").HorizontalAlignment = xlCenter
    .Range("G1:H1").VerticalAlignment = xlTop
    .Range("G1:H1").WrapText = True
    
    
    .Range("G2:G" & LastRow).FormulaR1C1 = "=IF(ISNA((VLOOKUP(RC[-5],'Contracted employees'!R4C6:R10000C38,33,FALSE)))=TRUE,""Excluded"",""Included"")"
    .Range("H2:H" & LastRow).FormulaR1C1 = "=IF(RC[-1]=""Included"",VLOOKUP(RC[-6],'Contracted employees'!R4C6:R10000C38,33,FALSE),"""")"
        .Range("H2:H" & LastRow).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
        
    End With
   
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "TaxData!R1C1:R50000C8", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="TaxData!R1C10", TableName:="PivotTable2", _
        DefaultVersion:=xlPivotTableVersion14
    Sheets("TaxData").Select
    Cells(1, 10).Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("In or Out")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Amount")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Amount"), "Sum of Amount", xlSum
   
   
Range("K2:K5").NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"

Sheets("Audit Check").Activate

With Worksheets("Audit Check")

    .Range("B5").FormulaR1C1 = "=GETPIVOTDATA(""Entered DR"",Allowances!R1C7,""In or Out"",""Included"")"
    .Range("B7").FormulaR1C1 = "=GETPIVOTDATA(""Amount"",TaxData!R1C10,""In or Out"",""Included"")"
    
End With
             


End Sub

Private Sub Delete_DATACHK_Sheet()

Application.ScreenUpdating = False

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets
         If ws.Name Like "Data Checks" Then
                 
                Application.DisplayAlerts = False
                
                ws.Delete
                
                Application.DisplayAlerts = True
            
            End If
                    
  Next
        
       
End Sub

'The following can be called with any range as parameter:

Sub SetRangeBorder(poRng As Range)

    If Not poRng Is Nothing Then
    
        poRng.Borders(xlDiagonalDown).LineStyle = xlNone
        poRng.Borders(xlDiagonalUp).LineStyle = xlNone
        poRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        poRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        poRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        poRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        poRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        poRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
    End If
    
 'Examples of how to call:
    'Call SetRangeBorder(Range("C11"))
    'Call SetRangeBorder(Range("A" & result))
    'Call SetRangeBorder(DT.Cells(I, 6))
    'Call SetRangeBorder(Range("A3:I" & endRow))
    
End Sub

Private Sub Final_Data_Intg_Chk()

Application.ScreenUpdating = False

On Error Resume Next

With Worksheets("Contracted Employees")

    .AutoFilterMode = False

Application.Calculation = xlCalculationAutomatic

LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
LastRowA = Sheets("Data Checks").Cells(Rows.Count, "A").End(xlUp).Row

    Range("AE4:AV" & LastRow).Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

    .Range("A3:AV3").AutoFilter Field:=27, Criteria1:="VMO"
    .Range("A3:AV3").AutoFilter Field:=31, Criteria1:="#N/A"
        
    .Range("AE4:AH" & LastRow).SpecialCells(xlCellTypeVisible).Clear
    .Range("AE4:AE" & LastRow).SpecialCells(xlCellTypeVisible) = "VMO No Employer Superannuation"
    
    .Range("A3:AV3").AutoFilter Field:=27
    
   
        With Range("AE4:AI" & LastRow).SpecialCells(xlCellTypeVisible).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 10485588 ' sea green
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
LastRowChk = Range("AE4:AE" & LastRow).SpecialCells(xlCellTypeVisible).Count
       
    If LastRowChk > 0 Then
 
    .Range("A4:AI" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Data Checks").Range("A" & LastRowA + 2)
            
        Else: End
        
    End If
    
    .Range("AE1:AI1").Copy _
        Destination:=Sheets("Data Checks").Range("AE" & 1 + LastRowA)

    Selection.AutoFilter Field:=31

Sheets("Data Checks").Activate

    Range("A" & 1 + LastRowA).FormulaR1C1 = "The below entries are missing from the Superannuation data upload (see green shaded cells)  - They HAVE been included in the Contracted Employees Tab"
    Range("A" & 1 + LastRowA).Font.Bold = True

        With Range("A" & 1 + LastRowA).Characters(Start:=100, Length:=4).Font
          .Underline = xlUnderlineStyleDouble
        End With
        
    Range("A" & 1 + LastRowA & ":AI" & 1 + LastRowA).Activate
    
        With Range("A" & 1 + LastRowA & ":AI" & 1 + LastRowA).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 2037680
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    
        With Range("A" & 1 + LastRowA & ":AI" & 1 + LastRowA).BorderAround([xlContinuous], _
            [xlThin], [xlColorIndexAutomatic])

        End With
End With

Columns("AE:AI").ColumnWidth = 20
Range("A" & 1 + LastRowA).EntireRow.AutoFit
   
Application.Calculation = xlCalculationManual
    
End Sub

Sub remove_minister_salary_from_consolidated_sheet()




'Application.ScreenUpdating = False
'Application.EnableEvents = False


Call Delete_Minister_Sheet

ActiveWorkbook.Worksheets.Add(After:=Sheet2).Name = "Minister's Contract"

    With ActiveWorkbook.Sheets("Minister's Contract").Tab
            .Color = 65535 'Yellow
            .TintAndShade = 0
        End With




Sheets("Contracted Employees").Activate

With Worksheets("Contracted Employees")

            .AutoFilterMode = False

    .Range("A3:AV3").AutoFilter Field:=17, Criteria1:="*MIG*"
    
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    LastRowChk = .Range("A3:AV" & LastRow).SpecialCells(xlCellTypeVisible).Count
       
    If LastRowChk < 3 Then
    
        .Range("A1:AV" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
                Destination:=Sheets("Minister's Contract").Range("A1:AV" & LastRow)
            
        Else
      
      End If
            
    If LastRowChk > 3 Then
      
        .Range("A1:AV" & LastRow).SpecialCells(xlCellTypeVisible).Copy _
                Destination:=Sheets("Minister's Contract").Range("A1:AV" & LastRow)
        Else
     
    End If
    
End With
   

Rows(lastA_RW & ":" & lastA_RW).Select
    
    Selection.Delete Shift:=xlUp

    


'This part unfilters the transactions on column 'D'

 Selection.AutoFilter Field:=4
  
' This section returns the active sheet to 'Instructions'

    Sheets("Instructions").Select
    
'Application.ScreenUpdating = True
'Application.EnableEvents = True
'Application.StatusBar = False
         
End Sub

Private Sub Delete_Minister_Sheet()

Application.ScreenUpdating = False

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets
         If ws.Name Like "Minister's Contract" Then
                 
                Application.DisplayAlerts = False
                
                ws.Delete
                
                Application.DisplayAlerts = True
            
            End If
                    
  Next
        
       
End Sub
Sub Simple_Check()

Application.ScreenUpdating = False
Application.EnableEvents = False



ButtonChk = Worksheets("Contracted employees").Range("AU1").Value

If ButtonChk <> "Year 1" Then

    MsgBox "Not able to proceed, you first need to run the 'Finalize Commitment Note' macro."
    
        Else
        
   Call Generate_Audit_Check
        
 End If
 
Application.ScreenUpdating = True
Application.EnableEvents = True
 
End Sub
