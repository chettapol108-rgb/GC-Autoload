Option Explicit

Public Type AutoLoadConfig
    InputSheetName As String
    NormalizedSheetName As String
    
    sku As String          ' Full SKU, e.g. L007-i
    parentSku As String    ' Parent SKU, e.g. L007
    skuCode As String      ' Variant code, e.g. i
    ColSku As String

    SetsPerRow As Long
    CopyNameA As Long
    CopyNameB As Long
    CopyNameC As Long
    CopyNameD As Long
    CopyBarcode As Long
    
    DateYMD As String
    DateMMDD As String
    ShopCode As String
    
    FirstDataRow As Long
    
    ColRawInput As String
    ColBarcode As String
    ColQty As String
End Type

Public Function ReadConfig(ByVal wb As Workbook) As AutoLoadConfig
    Dim cfg As AutoLoadConfig
    Dim ws As Worksheet
    
    cfg.InputSheetName = "AutoLoad-Input"
    cfg.NormalizedSheetName = "NormalizedInput"
    
    Set ws = wb.Sheets(cfg.InputSheetName)
    
    cfg.sku = Trim$(CStr(ws.Range("B1").value))
    SplitSku cfg.sku, cfg.parentSku, cfg.skuCode
    
    cfg.SetsPerRow = CLng(Val(ws.Range("A3").value))
    cfg.CopyNameA = CLng(Val(ws.Range("B3").value))
    cfg.CopyNameB = CLng(Val(ws.Range("C3").value))
    cfg.CopyNameC = CLng(Val(ws.Range("D3").value))
    cfg.CopyNameD = CLng(Val(ws.Range("E3").value))
    cfg.CopyBarcode = CLng(Val(ws.Range("F3").value))
    
    cfg.DateYMD = FormatDateYMD(ws.Range("G3").value)
    cfg.DateMMDD = Mid$(cfg.DateYMD, 5)
    cfg.ShopCode = GetShopCode(ws.Range("H3").value)

    cfg.FirstDataRow = 5
    
    cfg.ColSku = "A"
    cfg.ColBarcode = "B"
    cfg.ColQty = "C"
    cfg.ColRawInput = "D"
    
    ReadConfig = cfg
End Function

Private Sub SplitSku( _
    ByVal fullSku As String, _
    ByRef parentSku As String, _
    ByRef skuCode As String _
)

    Dim p As Long

    fullSku = Trim$(fullSku)
    p = InStrRev(fullSku, "-")

    If p > 0 Then
        parentSku = Left$(fullSku, p - 1)
        skuCode = Mid$(fullSku, p + 1)
    Else
        parentSku = fullSku
        skuCode = fullSku
    End If

End Sub

Private Function FormatDateYMD(ByVal inputValue As Variant) As String
    If IsDate(inputValue) Then
        FormatDateYMD = Format$(CDate(inputValue), "yyyymmdd")
    Else
        FormatDateYMD = CStr(inputValue)
    End If
End Function

Private Function GetShopCode(ByVal platformValue As Variant) As String
    GetShopCode = UCase$(Trim$(CStr(platformValue)))
End Function

'------------------------------------------------------------------------------
' EnsureNormalizedSheet
'
' Returns the existing NormalizedInput worksheet.
' Shows an error and returns Nothing if the worksheet does not exist.
'------------------------------------------------------------------------------
Public Function EnsureNormalizedSheet(ByVal wb As Workbook, ByRef cfg As AutoLoadConfig) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Sheets(cfg.NormalizedSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & cfg.NormalizedSheetName & "' does not exist.", vbCritical
        Set EnsureNormalizedSheet = Nothing
        Exit Function
    End If
    
    Set EnsureNormalizedSheet = ws
End Function
