Option Explicit

Public Sub Ribbon_GetVersionLabel(ByVal control As Object, ByRef returnedVal)
    returnedVal = "GC AutoLoad v" & GC_AUTOLOAD_VERSION
End Sub

'------------------------------------------------------------------------------
' GC_TranslateInput
'
' Entry point for the Translate Input button.
' Uses the currently active workbook as the target workbook.
'------------------------------------------------------------------------------
Public Sub GC_TranslateInput()

    Dim wb As Workbook

    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If

    TranslateRawInput_ToNormalizedRows wb

End Sub


'------------------------------------------------------------------------------
' GC_LoadData
'
' Entry point for the Load Data button.
'------------------------------------------------------------------------------
Public Sub GC_LoadData()

    Dim wb As Workbook

    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If

    LoadData_ToSKUSheet wb

End Sub


'------------------------------------------------------------------------------
' GC_ClearData
'
' Entry point for the Clear Data button.
'------------------------------------------------------------------------------
Public Sub GC_ClearData()

    Dim wb As Workbook

    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If

    ClearData_BySKU wb

End Sub


'------------------------------------------------------------------------------
' GC_ExportXML
'
' Entry point for the Export XML button.
'------------------------------------------------------------------------------
Public Sub GC_ExportXML()

    Dim wb As Workbook

    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If

    ExportXML_UTF8_BySKU wb

End Sub


'------------------------------------------------------------------------------
' GC_ClipboardToText
'
' Entry point for the Clipboard To Text button.
'
' Clipboard paste works with the current selection, so no workbook parameter
' is required.
'------------------------------------------------------------------------------
Public Sub GC_ClipboardToText()

    PasteClipboardAsText

End Sub

Public Sub Ribbon_ClipboardToText(ByVal control As Object)
    GC_ClipboardToText
End Sub

Public Sub Ribbon_TranslateInput(ByVal control As Object)
    GC_TranslateInput
End Sub

Public Sub Ribbon_LoadData(ByVal control As Object)
    GC_LoadData
End Sub

Public Sub Ribbon_ExportXML(ByVal control As Object)
    GC_ExportXML
End Sub

Public Sub Ribbon_ClearData(ByVal control As Object)
    GC_ClearData
End Sub
