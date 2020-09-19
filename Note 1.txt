'SplitWorkbook
Sub SplitWorkbook()
'by M.Rizmansyah Farizi
'Updateby20170815
Dim FileExtStr As String
Dim FileFormatNum As Long
Dim xWs As Worksheet
Dim xWb As Workbook
Dim xDir As String
Dim folder As FileDialog

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set xWb = Application.ThisWorkbook
Set folder = Application.FileDialog(msoFileDialogFolderPicker)
        If folder.Show <> -1 Then Exit Sub
For Each xWs In xWb.Worksheets
If xWs.Name <> "Data" And xWs.Name <> "Raw" Then
    xWs.Copy
    If Val(Application.Version) < 12 Then
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        Select Case xWb.FileFormat
            Case 51:
                FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If Application.ActiveWorkbook.HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56:
                FileExtStr = ".xls": FileFormatNum = 56
            Case Else:
                FileExtStr = ".xlsb": FileFormatNum = 50
        End Select
    End If
    xDir = folder.SelectedItems(1) & "\" & Application.ActiveWorkbook.Sheets(1).Name & FileExtStr
    Application.ActiveWorkbook.SaveAs xDir, FileFormat:=FileFormatNum
    Application.ActiveWorkbook.Close False
    End If
Next
MsgBox "Complete~~"
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub


