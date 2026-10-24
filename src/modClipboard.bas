'Convert all Order ID in clipboard to text
Public Sub PasteClipboardAsText()
    Dim startCell As Range
    Dim clipboardText As String
    Dim rowsArr As Variant
    Dim colsArr As Variant
    Dim r As Long, c As Long
    Dim targetCell As Range
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the starting cell first.", vbExclamation
        Exit Sub
    End If
    
    Set startCell = Selection.Cells(1, 1)
    
    clipboardText = GetClipboardText_PS()
    
    If Len(clipboardText) = 0 Then
        MsgBox "Clipboard is empty or cannot be read.", vbExclamation
        Exit Sub
    End If
    
    clipboardText = Replace(clipboardText, vbCrLf, vbLf)
    clipboardText = Replace(clipboardText, vbCr, vbLf)
    
    rowsArr = Split(clipboardText, vbLf)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo SafeExit
    
    For r = LBound(rowsArr) To UBound(rowsArr)
        If r = UBound(rowsArr) And rowsArr(r) = "" Then Exit For
        
        colsArr = Split(rowsArr(r), vbTab)
        
        For c = LBound(colsArr) To UBound(colsArr)
            Set targetCell = startCell.Offset(r, c)
            targetCell.NumberFormat = "@"
            targetCell.value = CStr(colsArr(c))
        Next c
    Next r

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If Err.Number = 0 Then
        MsgBox "Clipboard pasted as text successfully.", vbInformation
    Else
        MsgBox "Paste failed: " & Err.Description, vbCritical
    End If
End Sub

Private Function GetClipboardText_PS() As String
    Dim sh As Object
    Dim execObj As Object
    Dim cmd As String
    Dim result As String
    
    Set sh = CreateObject("WScript.Shell")
    
    cmd = "powershell -NoProfile -STA -Command ""Get-Clipboard"""
    
    Set execObj = sh.Exec(cmd)
    result = execObj.StdOut.ReadAll
    
    If Right$(result, 2) = vbCrLf Then
        result = Left$(result, Len(result) - 2)
    ElseIf Right$(result, 1) = vbLf Then
        result = Left$(result, Len(result) - 1)
    End If
    
    GetClipboardText_PS = result
End Function