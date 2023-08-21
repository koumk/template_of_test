Sub AllSheetPrintPDF()

    Dim ws As Worksheet
    Dim i As Integer, j As Integer
    Dim msg As String
    Dim msgA1 As String
    i = 0
    j = 0
    For Each ws In Worksheets
        ws.Activate
        If ActiveWindow.Zoom <> 100 Then
         msg = msg & vbCrLf & ws.Name & "：" & CStr(ActiveWindow.Zoom)
         i = i + 1
        End If
        
        addr = ActiveCell.Address
        If addr <> "$A$1" Then
         msgA1 = msgA1 & vbCrLf & ws.Name & "：" & addr
         j = j + 1
        End If
    Next ws
    If i > 0 Then
        MsgBox ("表示倍率(Zoom)がおかしい：" + msg)
    End If
    If j > 0 Then
         MsgBox ("A1セルに選択されていない：" + msgA1)
    End If
    
    If i > 0 Or j > 0 Then
        Exit Sub
    End If
    
    Worksheets.Select
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, fileName:="c:\temp\test.pdf" 'WorkbookをPDF出力
End Sub