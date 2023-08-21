Sub AllSheetPrintPDF()

    Dim ws As Worksheet
    Dim i As Integer
    Dim msg As String
    i = 0
    For Each ws In Worksheets
        ws.Activate
        If ActiveWindow.Zoom <> 100 Then
         msg = msg & ws.Name & "：" & CStr(ActiveWindow.Zoom)
         i = i + 1
        End If
    Next ws
    If i > 0 Then
        MsgBox ("表示倍率(Zoom)がおかしい：" + msg)
        Exit Sub
    End If
    
    Worksheets.Select
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, fileName:="c:\temp\test.pdf" 'WorkbookをPDF出力
End Sub