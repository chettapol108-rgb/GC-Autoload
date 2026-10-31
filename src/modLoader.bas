Option Explicit

Public Sub LoadData_ToSKUSheet(ByVal wb As Workbook)
    Dim cfg As AutoLoadConfig, wsNorm As Worksheet
    Dim skus As Object, sku As Variant
    On Error GoTo Failed
    cfg = ReadConfig(wb)
    Set wsNorm = EnsureNormalizedSheet(wb, cfg)
    If wsNorm Is Nothing Then Exit Sub
    Set skus = NormalizedSkus(wsNorm)
    RequireBatch wb, cfg, wsNorm, skus, False
    SetBatchProperty wsNorm, "GC_BatchLoaded", "NO"
    ' Preflight all destinations before clearing any of them.
    For Each sku In skus.Keys
        ReadSkuConfig wb, cfg, CStr(sku)
        LoadOneSku wb, cfg, True
    Next sku
    For Each sku In skus.Keys
        ReadSkuConfig wb, cfg, CStr(sku)
        LoadOneSku wb, cfg, False
    Next sku
    SetBatchProperty wsNorm, "GC_BatchLoaded", "YES"
    MsgBox "Load Data completed. Platform: " & cfg.ShopCode & "; SKUs: " & skus.Count, vbInformation
    Exit Sub
Failed:
    MsgBox "Load Data stopped: " & Err.Description & vbCrLf & "XML export is blocked until Load Data succeeds.", vbExclamation
End Sub

Private Sub LoadOneSku(ByVal wb As Workbook, ByRef cfg As AutoLoadConfig, ByVal validateOnly As Boolean)
    Dim wsOutput As Worksheet
    Dim wsNorm As Worksheet
    Dim lastRow As Long
    Dim totalCopiesPerSet As Long
    Dim i As Long, outRow As Long, dataSetCount As Long
    Dim startCol As Long, destCol As Long
    Dim tempValue As String
    Dim rawOID As String
    Dim rawBarcode As String
    Dim encodedBarcode As String
    Dim j As Long
    Dim copyCount As Long
    Dim copyIdx As Long
    Dim colDATE As Long, colSHOP As Long, ColSku As Long, colFN As Long
    Dim rowNum As Long
    Dim setIndex As Long, fieldCol As Long
    Dim totalSets As Double

    ' Calculate configuration values before validating them.
    totalCopiesPerSet = cfg.CopyNameA + cfg.CopyNameB + cfg.CopyNameC + cfg.CopyNameD + cfg.CopyOID + cfg.CopyBarcode

    If cfg.SetsPerRow <= 0 Then
        Err.Raise vbObjectError + 2200, , "Invalid sets per row for SKU: " & cfg.sku
    End If

    If cfg.CopyNameA < 0 Or cfg.CopyNameB < 0 Or cfg.CopyNameC < 0 Or _
       cfg.CopyNameD < 0 Or cfg.CopyOID < 0 Or cfg.CopyBarcode < 0 Then
        Err.Raise vbObjectError + 2201, , "Negative copy count for SKU: " & cfg.sku
    End If

    If totalCopiesPerSet <= 0 Then
        Err.Raise vbObjectError + 2202, , "No output copies for SKU: " & cfg.sku
    End If
    Set wsNorm = EnsureNormalizedSheet(wb, cfg)
    If wsNorm Is Nothing Then Err.Raise vbObjectError + 2203, , "Invalid NormalizedInput schema."
    
    If cfg.sku = "" Then
        Err.Raise vbObjectError + 2204, , "Empty SKU."
    End If

    On Error Resume Next
    Set wsOutput = wb.Sheets(cfg.sku)
    On Error GoTo 0
    
    If wsOutput Is Nothing Then
        Err.Raise vbObjectError + 2205, , "Sheet '" & cfg.sku & "' does not exist."
    End If

    colDATE = GetHeaderColumn(wsOutput, "DATE")
    colSHOP = GetHeaderColumn(wsOutput, "SHOP")
    ColSku = GetHeaderColumn(wsOutput, "SKU")
    colFN = GetHeaderColumn(wsOutput, "FN")

    If colDATE = 0 Or colSHOP = 0 Or ColSku = 0 Or colFN = 0 Then
        Err.Raise vbObjectError + 2206, , cfg.sku & ": Missing DATE, SHOP, SKU or FN."
    End If

    ' Check the template before clearing existing output.
    If cfg.SetsPerRow > wsOutput.Columns.Count \ totalCopiesPerSet Then
        Err.Raise vbObjectError + 2207, , cfg.sku & ": Output exceeds the worksheet column limit."
    End If

    fieldCol = cfg.SetsPerRow * totalCopiesPerSet
    If colDATE <= fieldCol Or colSHOP <= fieldCol Or ColSku <= fieldCol Or colFN <= fieldCol Then
        Err.Raise vbObjectError + 2208, , cfg.sku & ": DATE, SHOP, SKU and FN must follow all output sets."
    End If

    For setIndex = 0 To cfg.SetsPerRow - 1
        fieldCol = setIndex * totalCopiesPerSet + _
                   cfg.CopyNameA + cfg.CopyNameB + cfg.CopyNameC + cfg.CopyNameD + 1
        For j = 1 To cfg.CopyOID
            If Right$(Replace$(UCase$(Trim$(CStr(wsOutput.Cells(1, fieldCol).value))), "_", "-"), 4) <> "-OID" Then
                Err.Raise vbObjectError + 2209, , cfg.sku & ": Expected Nx_OID or Nx-OID in column " & fieldCol
            End If
            fieldCol = fieldCol + 1
        Next j
        For j = 1 To cfg.CopyBarcode
            If Right$(Replace$(UCase$(Trim$(CStr(wsOutput.Cells(1, fieldCol).value))), "_", "-"), 3) <> "-BC" Then
                Err.Raise vbObjectError + 2210, , cfg.sku & ": Expected Nx_BC or Nx-BC in column " & fieldCol
            End If
            fieldCol = fieldCol + 1
        Next j
    Next setIndex

    lastRow = wsNorm.Cells(wsNorm.rows.Count, "G").End(xlUp).Row
    If lastRow < 2 Then
        Err.Raise vbObjectError + 2211, , "No normalized data. Run Translate Input first."
    End If

    If wsOutput.ProtectContents Then Err.Raise vbObjectError + 2212, , "Protected output sheet: " & cfg.sku
    For i = 2 To lastRow
        If StrComp(Trim$(CStr(wsNorm.Cells(i, "A").Value2)), cfg.sku, vbTextCompare) = 0 Then
            copyCount = CLng(Val(wsNorm.Cells(i, "H").Value2))
            If copyCount <= 0 Then Err.Raise vbObjectError + 2213, , "Invalid normalized Qty on row " & i
            totalSets = totalSets + copyCount
            ' Validate barcode encoding before clearing destinations.
            If cfg.CopyBarcode > 0 Then encodedBarcode = CODE128_LIBRE(CStr(wsNorm.Cells(i, "G").Value2))
        End If
    Next i
    If totalSets = 0 Or totalSets > CDbl(wsOutput.Rows.Count - 1) * cfg.SetsPerRow Then
        Err.Raise vbObjectError + 2214, , "Invalid output row count for SKU: " & cfg.sku
    End If
    If validateOnly Then Exit Sub

    wsOutput.rows("2:" & wsOutput.rows.Count).ClearContents

    outRow = 2
    dataSetCount = 0
        
    For i = 2 To lastRow
        If UCase$(Trim$(CStr(wsNorm.Cells(i, "A").value))) = UCase$(cfg.sku) Then
        
            copyCount = CLng(Val(wsNorm.Cells(i, "H").value))
            If copyCount <= 0 Then copyCount = 1
            
            rawOID = CStr(wsNorm.Cells(i, "F").value)
            rawBarcode = CStr(wsNorm.Cells(i, "G").value)
            
            If rawBarcode <> "" And cfg.CopyBarcode > 0 Then
                encodedBarcode = CODE128_LIBRE(rawBarcode)
            Else
                encodedBarcode = ""
            End If
            
            For copyIdx = 1 To copyCount
                startCol = (dataSetCount Mod cfg.SetsPerRow) * totalCopiesPerSet + 1

                If dataSetCount > 0 And (dataSetCount Mod cfg.SetsPerRow) = 0 Then
                    outRow = outRow + 1
                End If

                destCol = startCol

                tempValue = wsNorm.Cells(i, "B").value
                For j = 1 To cfg.CopyNameA
                    wsOutput.Cells(outRow, destCol).value = tempValue
                    destCol = destCol + 1
                Next j

                tempValue = wsNorm.Cells(i, "C").value
                For j = 1 To cfg.CopyNameB
                    wsOutput.Cells(outRow, destCol).value = tempValue
                    destCol = destCol + 1
                Next j

                tempValue = wsNorm.Cells(i, "D").value
                For j = 1 To cfg.CopyNameC
                    wsOutput.Cells(outRow, destCol).value = tempValue
                    destCol = destCol + 1
                Next j

                tempValue = wsNorm.Cells(i, "E").value
                For j = 1 To cfg.CopyNameD
                    wsOutput.Cells(outRow, destCol).value = tempValue
                    destCol = destCol + 1
                Next j

                For j = 1 To cfg.CopyOID
                    wsOutput.Cells(outRow, destCol).NumberFormat = "@"
                    wsOutput.Cells(outRow, destCol).value = rawOID
                    destCol = destCol + 1
                Next j

                For j = 1 To cfg.CopyBarcode
                    wsOutput.Cells(outRow, destCol).value = encodedBarcode
                    destCol = destCol + 1
                Next j

                dataSetCount = dataSetCount + 1
            Next copyIdx
        End If
    Next i

    For rowNum = 2 To outRow
        wsOutput.Cells(rowNum, colDATE).value = cfg.DateYMD
        wsOutput.Cells(rowNum, colSHOP).value = cfg.ShopCode
        wsOutput.Cells(rowNum, ColSku).value = cfg.skuCode
        wsOutput.Cells(rowNum, colFN).value = cfg.DateMMDD & "-" & cfg.ShopCode & "-" & cfg.sku & "-" & Format$(rowNum - 1, "00")
    Next rowNum

End Sub


Public Sub ClearData_BySKU(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim skuName As String
    Dim cfg As AutoLoadConfig

    cfg = ReadConfig(wb)
    skuName = Trim$(CStr(wb.Worksheets(cfg.InputSheetName).Range("B1").Value2))

    If skuName = "" Then
        MsgBox "SKU name in cell B1 is empty.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set ws = wb.Sheets(skuName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & skuName & "' does not exist.", vbCritical
        Exit Sub
    End If

    ws.rows("2:" & ws.rows.Count).ClearContents
    SetBatchProperty wb.Worksheets(cfg.NormalizedSheetName), "GC_BatchLoaded", "NO"

    MsgBox "Data from row 2 down in sheet '" & skuName & _
           "' has been cleared.", vbInformation
End Sub

Private Function GetHeaderColumn(ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim colIdx As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For colIdx = 1 To lastCol
        If UCase$(Trim$(CStr(ws.Cells(1, colIdx).value))) = UCase$(headerName) Then
            GetHeaderColumn = colIdx
            Exit Function
        End If
    Next colIdx
    
    GetHeaderColumn = 0
End Function