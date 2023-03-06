Attribute VB_Name = "Module11"
Sub PivotCopyFormatValues()
'Application.ScreenUpdating = False
'select pivot table cell first
Dim ws As Worksheet
Set ws = ActiveSheet
Dim pt As pivotTable
Set pt = ws.PivotTables("Summary")
Dim rngPT As Range
Dim rngPTa As Range
Dim rngCopy As Range
Dim lRowTop As Long
Dim lRowsPT As Long
Dim lRowPage As Long
Dim msgSpace As String
Dim tbl As ListObject
Dim tblrng As Range
On Error Resume Next

Set rngPTa = pt.PageRange

'On Error GoTo errHandler

If pt Is Nothing Then
    MsgBox "Could not copy pivot table for active cell"
    GoTo exitHandler
End If

If pt.PageFieldOrder = xlOverThenDown Then
  If pt.PageFields.Count > 1 Then
    msgSpace = "Horizontal filters with spaces." _
      & vbCrLf _
      & "Could not copy Filters formatting."
  End If
End If


'###################################Pivot layout changes and data extractions###################################
'set format of the table to get all needed values
pt.RowAxisLayout xlTabularRow
pt.RepeatAllLabels xlRepeatLabels
    
    With ActiveSheet.PivotTables("Summary").PivotFields("Brand")
        .PivotItems("(blank)").Visible = False
    End With
'remove subtotals
    For Each pf In pt.PivotFields
        On Error Resume Next
        pf.Subtotals(1) = False
    Next pf
  
  ActiveSheet.PivotTables("Summary").PivotFields("Media Property").ShowDetail = _
        True
'Copy and paste pivot
Set rngPT = pt.TableRange1
'Define paste DESTINATION as lrowTop = end of PivotRange + 100
lRowTop = rngPT.Rows(rngPT.Rows.Count + 100).row
lRowsPT = rngPT.Rows.Count
'Define SOURCE to copy
Set rngCopy = rngPT.Resize(lRowsPT - 1)

rngCopy.Copy
ws.Cells(lRowTop, 1).PasteSpecial _
    Paste:=xlPasteAll



'###################################END of Layout changes and data extraction###################################


'Sets new values' range to becaome a TABLE
Set tblrng = Range(Cells(lRowTop, 1), Cells(lRowTop + lRowsPT, 15))
Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, tblrng, , xlYes)
tbl.TableStyle = "TableStyleMedium2"

'Checks for red conditional formatting and sends the data to the RED VALUES sheet
Dim rowHeaders() As String
    ReDim rowHeaders(1 To tblrng.Rows.Count)
    For i = 1 To tblrng.Rows.Count
        rowHeaders(i) = tblrng.Cells(i, 1).Value
    Next i
    
    Dim colHeaders() As String
    ReDim colHeaders(1 To tblrng.Columns.Count)
    For i = 1 To tblrng.Columns.Count
        colHeaders(i) = tblrng.Cells(1, i).Value
    Next i
    
    Dim outputSheet As Worksheet
    Set outputSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    outputSheet.Name = "Red Cells"
    outputSheet.Range("A1").Value = "Campaign"
    outputSheet.Range("B1").Value = "Partner"
    outputSheet.Range("C1").Value = "Metric"
    
    Dim outputRow As Long
    outputRow = 2
    
    For i = 1 To tblrng.Rows.Count
        For j = 1 To tblrng.Columns.Count
            If tblrng.Cells(i, j).DisplayFormat.Font.Color = vbRed Then
                outputSheet.Cells(outputRow, 1).Value = tblrng.Cells(i, 3).Value
                outputSheet.Cells(outputRow, 2).Value = tblrng.Cells(i, 2).Value
                outputSheet.Cells(outputRow, 3).Value = colHeaders(j)
                
                
                outputRow = outputRow + 1
            End If
        Next j
    Next i



If Not rngPTa Is Nothing Then
    lRowPage = rngPTa.Rows(1).row
    rngPTa.Copy Destination:=ws.Cells(lRowPage, 1)
End If
    
ws.Columns.AutoFit
If msgSpace <> "" Then
  MsgBox msgSpace
End If



'Returns the pivot to it's original format
pt.RowAxisLayout xlCompactRow
    With ActiveSheet.PivotTables("Summary").PivotFields("Brand")
        .PivotItems("(blank)").Visible = True
    End With
    For Each pf In pt.PivotFields
        On Error Resume Next
        pf.Subtotals(1) = True
    Next pf
    ActiveSheet.PivotTables("Summary").PivotFields("Media Property").ShowDetail = _
        False
    
tbl.Delete

exitHandler:
    Exit Sub
'errHandler:
    'MsgBox "Could not copy pivot table for active cell"
    'Resume exitHandler
    

'Application.ScreenUpdating = True
    
End Sub


Sub makeTable()
  ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$25:$E$46"), , xlYes).Name = _
        "Table1"
End Sub

Sub CheckRedCells()
    
    Dim rowHeaders() As String
    ReDim rowHeaders(1 To tblrng.Rows.Count)
    For i = 1 To tblrng.Rows.Count
        rowHeaders(i) = tblrng.Cells(i, 1).Value
    Next i
    
    Dim colHeaders() As String
    ReDim colHeaders(1 To tblrng.Columns.Count)
    For i = 1 To tblrng.Columns.Count
        colHeaders(i) = tblrng.Cells(1, i).Value
    Next i
    
    Dim outputSheet As Worksheet
    Set outputSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    outputSheet.Name = "Red Cells"
    outputSheet.Range("A1").Value = "Campaign"
    outputSheet.Range("B1").Value = "Partner"
    outputSheet.Range("C1").Value = "Metric"
    
    Dim outputRow As Long
    outputRow = 2
    
    For i = 2 To tblrng.Rows.Count
        For j = 2 To tblrng.Columns.Count
            If tblrng.Cells(i, j).DisplayFormat.Font.Color = vbRed Then
                outputSheet.Cells(outputRow, 1).Value = tblrng.Cells(i, 2).Value
                outputSheet.Cells(outputRow, 2).Value = tblrng.Cells(i, 1).Value
                outputSheet.Cells(outputRow, 3).Value = colHeaders(j)
                
                
                outputRow = outputRow + 1
            End If
        Next j
    Next i
End Sub




