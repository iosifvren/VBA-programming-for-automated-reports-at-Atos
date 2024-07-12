Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("XXXX")
    
    Dim i As Integer
    For i = 1 To 7
        ws.Columns("AY:AY").Insert Shift:=xlToRight
    Next i
    
    ws.Range("AY1").Value = "XXXX"
    ws.Range("AZ1").Value = "XXXX"
    ws.Range("BA1").Value = "XXXX"
    ws.Range("BB1").Value = "XXXX"
    ws.Range("BC1").Value = "XXXX"
    ws.Range("BD1").Value = "XXXX"
    ws.Range("BE1").Value = "XXXX"
    
    ws.Range("AX:AX").Copy
    ws.Range("AY:BE").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("AY2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),AB2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),AB2))"
ws.Range("AY2:AY" & ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row).FillDown

ws.Range("AY:AY").Value = ws.Range("AY:AY").Value
End Sub

Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("XXXX")
    
    With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("C").TextToColumns Destination:=.Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("O").TextToColumns Destination:=.Range("O1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AJ").TextToColumns Destination:=.Range("AJ1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AM").TextToColumns Destination:=.Range("AM1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        .Columns("J").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("K").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("W").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

.Range("AZ2").Formula = "=VLOOKUP(AY2,[XXXX.xlsx]Sheet1!$A:$G,7,0)"
.Range("BA2").Formula = "=VLOOKUP(AY2,[XXXX.xlsx]Sheet1!$A:$E,5,0)"
.Range("BB2").Formula = "=VLOOKUP(AY2,[XXXX.xlsx]Sheet1!$A:$F,6,0)"
.Range("BC2").Formula = "=VLOOKUP(AY2,[XXXX.xlsx]Sheet1!$A:$H,8,0)"
.Range("BD2").Formula = "=VLOOKUP(AY2,[XXXX.xlsx]Sheet1!$A:$J,10,0)"
.Range("BE2").Formula = "=VLOOKUP(AY2,[XXXX.xlsx]Sheet1!$A:$M,13,0)"

ws.Range("AZ2:AZ" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BA2:BA" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BB2:BB" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BC2:BC" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BD2:BD" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BE2:BE" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown

ws.Range("AZ2:BE" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).Value = ws.Range("AZ2:BE" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).Value
End With
End Sub

Sub CreatePivotTable()
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim LastRow As Long
    Dim LastCol As Long

    Set DSheet = Worksheets("XXXX")

    LastRow = DSheet.Cells(DSheet.Rows.Count, 1).End(xlUp).Row
    LastCol = DSheet.Cells(1, DSheet.Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

    Set PCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=PRange)

    Set PSheet = Worksheets.Add

    Set PTable = PCache.CreatePivotTable( _
        TableDestination:=PSheet.Cells(1, 1), _
        TableName:="PivotTable1")

    With PTable
    End With

    With PTable
        .PivotFields("XXXX").Orientation = xlRowField
        
    With .PivotFields("XXXX")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of XXXX"
    End With
End With
End Sub

' "XXXX" replacing data due to Company Private Information
