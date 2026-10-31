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
    CopyOID As Long
    CopyBarcode As Long
    
    DateYMD As String
    DateMMDD As String
    ShopCode As String
    
    FirstDataRow As Long
    
    ColRawInput As String
    ColBarcode As String
    ColQty As String
    ColPlatform As String
End Type

Public Function ReadConfig(ByVal wb As Workbook) As AutoLoadConfig
    Dim cfg As AutoLoadConfig
    Dim ws As Worksheet
    
    cfg.InputSheetName = "AutoLoad-Input"
    cfg.NormalizedSheetName = "NormalizedInput"
    
    Set ws = wb.Sheets(cfg.InputSheetName)
    
    cfg.DateYMD = FormatDateYMD(ws.Range("H3").value)
    cfg.DateMMDD = Mid$(cfg.DateYMD, 5)
    cfg.ShopCode = GetShopCode(ws.Range("I3").value)

    cfg.FirstDataRow = 5
    
    cfg.ColSku = "A"
    cfg.ColBarcode = "B"
    cfg.ColPlatform = "C"
    cfg.ColQty = "D"
    cfg.ColRawInput = "E"
    
    ReadConfig = cfg
End Function

' Batch processing reads B:H directly from Database; B1 and A3:G3 are previews only.
Public Sub ReadSkuConfig(ByVal wb As Workbook, ByRef cfg As AutoLoadConfig, ByVal sku As String)
    Dim ws As Worksheet, r As Long, foundRow As Long, c As Long
    Dim headers As Variant
    Set ws = wb.Worksheets("Database")
    headers = Array("SKU", "Set", "Name A", "Name B", "Name C", "Name D", "OID", "Barcode")
    For c = 0 To UBound(headers)
        If StrComp(Trim$(CStr(ws.Cells(1, c + 1).Value2)), headers(c), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 2100, "ReadSkuConfig", "Unexpected Database header in column " & (c + 1)
        End If
    Next c
    For r = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If StrComp(Trim$(CStr(ws.Cells(r, "A").Value2)), Trim$(sku), vbTextCompare) = 0 Then
            If foundRow > 0 Then Err.Raise vbObjectError + 2101, , "Duplicate Database SKU: " & sku
            foundRow = r
        End If
    Next r
    If foundRow = 0 Then Err.Raise vbObjectError + 2102, , "SKU not found in Database: " & sku
    cfg.sku = Trim$(CStr(ws.Cells(foundRow, "A").Value2))
    SplitSku cfg.sku, cfg.parentSku, cfg.skuCode
    cfg.SetsPerRow = ReadCopyCount(ws.Cells(foundRow, "B"), True)
    cfg.CopyNameA = ReadCopyCount(ws.Cells(foundRow, "C"))
    cfg.CopyNameB = ReadCopyCount(ws.Cells(foundRow, "D"))
    cfg.CopyNameC = ReadCopyCount(ws.Cells(foundRow, "E"))
    cfg.CopyNameD = ReadCopyCount(ws.Cells(foundRow, "F"))
    cfg.CopyOID = ReadCopyCount(ws.Cells(foundRow, "G"))
    cfg.CopyBarcode = ReadCopyCount(ws.Cells(foundRow, "H"))
    If CDbl(cfg.CopyNameA) + cfg.CopyNameB + cfg.CopyNameC + cfg.CopyNameD + cfg.CopyOID + cfg.CopyBarcode = 0 Then
        Err.Raise vbObjectError + 2103, , "No output copies configured for SKU: " & cfg.sku
    End If
End Sub

Private Function ReadCopyCount(ByVal cell As Range, Optional ByVal positive As Boolean = False) As Long
    Dim n As Double
    If IsError(cell.Value2) Then GoTo InvalidCount
    If Not IsNumeric(cell.Value2) Or IsEmpty(cell.Value2) Then GoTo InvalidCount
    n = CDbl(cell.Value2)
    If n < 0 Or n <> Fix(n) Or n > 16384 Then GoTo InvalidCount
    If positive And n = 0 Then GoTo InvalidCount
    ReadCopyCount = CLng(n)
    Exit Function
InvalidCount:
    Err.Raise vbObjectError + 2104, , "Invalid copy count at Database!" & cell.Address(False, False)
End Function

Public Sub ValidateBatchInput(ByVal wb As Workbook, ByRef cfg As AutoLoadConfig)
    Dim ws As Worksheet
    Dim qtyHeader As String
    Set ws = wb.Worksheets(cfg.InputSheetName)
    If cfg.ShopCode <> "SP" And cfg.ShopCode <> "TT" And cfg.ShopCode <> "ALL" Then
        Err.Raise vbObjectError + 2105, , "Select SP, TT or ALL in AutoLoad-Input!I3, then run Translate Input."
    End If
    qtyHeader = UCase$(Trim$(CStr(ws.Range("D4").Value2)))
    If UCase$(Trim$(CStr(ws.Range("C4").Value2))) <> "PLATFORM" Or _
       (qtyHeader <> "QTY" And qtyHeader <> ChrW(&HE08) & ChrW(&HE33) & ChrW(&HE19) & ChrW(&HE27) & ChrW(&HE19)) Or _
       UCase$(Trim$(CStr(ws.Range("E4").Value2))) <> "RAW INPUT" Then
        Err.Raise vbObjectError + 2106, , "Input row 4 must have C=PLATFORM, D=Qty, E=Raw Input."
    End If
End Sub

Public Function MatchesPlatform(ByVal value As Variant, ByVal shopCode As String) As Boolean
    Dim isShopee As Boolean
    If shopCode = "ALL" Then
        MatchesPlatform = True
        Exit Function
    End If
    isShopee = (StrComp(Trim$(CStr(value)), "Shopee", vbTextCompare) = 0)
    Select Case shopCode
        Case "SP": MatchesPlatform = isShopee
        Case "TT": MatchesPlatform = Not isShopee
    End Select
End Function

' Persist batch identity without adding columns to the A:H normalized schema.
Public Sub SetBatchProperty(ByVal ws As Worksheet, ByVal key As String, ByVal value As String)
    Dim i As Long, part As Long
    For i = ws.CustomProperties.Count To 1 Step -1
        If ws.CustomProperties(i).Name = key Or Left$(ws.CustomProperties(i).Name, Len(key) + 1) = key & "_" Then
            ws.CustomProperties(i).Delete
        End If
    Next i
    ' Keep individual values short even when a batch contains many SKUs.
    ws.CustomProperties.Add key, CStr((Len(value) + 199) \ 200)
    For i = 1 To Len(value) Step 200
        part = part + 1
        ws.CustomProperties.Add key & "_" & part, Mid$(value, i, 200)
    Next i
End Sub

Public Function GetBatchProperty(ByVal ws As Worksheet, ByVal key As String) As String
    Dim prop As Object, count As Long, i As Long, part As Long
    For Each prop In ws.CustomProperties
        If prop.Name = key Then
            count = CLng(prop.Value)
            Exit For
        End If
    Next prop
    For i = 1 To count
        For Each prop In ws.CustomProperties
            If prop.Name = key & "_" & i Then
                GetBatchProperty = GetBatchProperty & CStr(prop.Value)
                part = part + 1
                Exit For
            End If
        Next prop
    Next i
    If part <> count Then GetBatchProperty = ""
End Function

Public Function NormalizedSkus(ByVal ws As Worksheet) As Object
    Dim result As Object, r As Long, sku As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For r = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        sku = Trim$(CStr(ws.Cells(r, "A").Value2))
        If sku <> "" Then result(sku) = True
    Next r
    Set NormalizedSkus = result
End Function

Public Function BatchConfigKey(ByVal wb As Workbook, ByRef cfg As AutoLoadConfig, ByVal skus As Object) As String
    Dim sku As Variant, result As String
    result = cfg.ShopCode & "|" & cfg.DateYMD
    For Each sku In skus.Keys
        ReadSkuConfig wb, cfg, CStr(sku)
        result = result & "|" & cfg.sku & ":" & cfg.SetsPerRow & ":" & cfg.CopyNameA & ":" & _
                 cfg.CopyNameB & ":" & cfg.CopyNameC & ":" & cfg.CopyNameD & ":" & cfg.CopyOID & ":" & cfg.CopyBarcode
    Next sku
    BatchConfigKey = result
End Function

Public Sub RequireBatch(ByVal wb As Workbook, ByRef cfg As AutoLoadConfig, ByVal ws As Worksheet, ByVal skus As Object, ByVal loaded As Boolean)
    ValidateBatchInput wb, cfg
    If skus.Count = 0 Then Err.Raise vbObjectError + 2107, , "No normalized data. Run Translate Input."
    If GetBatchProperty(ws, "GC_BatchReady") <> "YES" Then
        Err.Raise vbObjectError + 2108, , "Run Translate Input successfully before continuing."
    End If
    If GetBatchProperty(ws, "GC_BatchConfig") <> BatchConfigKey(wb, cfg, skus) Then
        Err.Raise vbObjectError + 2109, , "Platform, date or Database settings changed. Run Translate Input again."
    End If
    If loaded And GetBatchProperty(ws, "GC_BatchLoaded") <> "YES" Then
        Err.Raise vbObjectError + 2110, , "Run Load Data successfully before exporting XML."
    End If
End Sub

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

    If UCase$(Trim$(CStr(ws.Range("F1").value))) <> "OID" Or _
       UCase$(Trim$(CStr(ws.Range("G1").value))) <> "BARCODE" Or _
       UCase$(Trim$(CStr(ws.Range("H1").value))) <> "QTY" Then
        MsgBox "NormalizedInput must have F1=OID, G1=Barcode and H1=Qty. " & _
               "Update the template and run Translate Input again.", vbCritical
        Set EnsureNormalizedSheet = Nothing
        Exit Function
    End If
    
    Set EnsureNormalizedSheet = ws
End Function
