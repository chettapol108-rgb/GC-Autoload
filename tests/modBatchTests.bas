Attribute VB_Name = "modBatchTests"
Option Explicit

' Import with the source modules into an isolated development VBA project.
' Uses only synthetic data. Production MsgBox dialogs remain interactive.
' Leaves the unsaved synthetic workbook open for manual XML export checks.
Public Sub GC_TestBatchProcessing()
    Dim wb As Workbook, inputSheet As Worksheet, norm As Worksheet, db As Worksheet
    Dim cfg As AutoLoadConfig, skus As Object, text As String
    Dim blocked As Boolean, failure As String
    On Error GoTo Failed
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set inputSheet = wb.Worksheets(1)
    inputSheet.Name = "AutoLoad-Input"
    Set db = wb.Worksheets.Add(After:=inputSheet)
    db.Name = "Database"
    Set norm = wb.Worksheets.Add(After:=db)
    norm.Name = "NormalizedInput"
    WriteTestRow db, 1, Array("SKU", "Set", "Name A", "Name B", "Name C", "Name D", "OID", "Barcode")
    WriteTestRow db, 2, Array("TEST-A", 2, 1, 0, 0, 0, 1, 1)
    WriteTestRow db, 3, Array("TEST-B", 1, 2, 1, 0, 0, 0, 0)
    WriteTestRow norm, 1, Array("SKU", "NameA", "NameB", "NameC", "NameD", "OID", "Barcode", "Qty")
    WriteTestRow inputSheet, 4, Array("SKU", "Barcode (Order ID)", "PLATFORM", "Qty", "Raw Input")
    inputSheet.Range("H3").NumberFormat = "@"
    inputSheet.Range("H3").Value2 = "20261031"
    inputSheet.Range("I3").Value2 = "SP"
    ' Deliberately invalid preview: the batch must not depend on B1 or A3:G3.
    inputSheet.Range("B1").Value = CVErr(xlErrNA)
    MakeTestOutput wb, "TEST-A", Array("N1-A1", "N1_OID", "N1_BC", "N2-A1", "N2-OID", "N2-BC", "DATE", "SHOP", "SKU", "FN")
    MakeTestOutput wb, "TEST-B", Array("N1-A1", "N1-A2", "N1-B1", "DATE", "SHOP", "SKU", "FN")
    ' Thai, Chinese, Korean, emoji surrogate pair, and all XML escape characters.
    text = ChrW(&HE01) & ChrW(&H4E2D) & ChrW(&HD55C) & ChrW(-10179) & ChrW(-8704) & " & < > "" '"
    WriteTestRow inputSheet, 5, Array("TEST-A", "00012345678901234567", " Shopee ", 1, "TEST-A *3pcs || " & text)
    WriteTestRow inputSheet, 6, Array("TEST-A", "00012345678901234567", "shopee", 2, "TEST-A *3pcs || " & text)
    WriteTestRow inputSheet, 7, Array("TEST-B", "00012345678901234567", "Shopee", 1, "#1|| BETA #2|| GAMMA")
    WriteTestRow inputSheet, 8, Array("TEST-A", "TT-001", "TikTok", 2, "ALPHA")
    WriteTestRow inputSheet, 9, Array("TEST-B", "FB-001", "Facebook", 1, "#1|| DELTA #2|| EPSILON")
    RequireTest MatchesPlatform(" SHOPEE ", "SP"), "Shopee comparison"
    RequireTest Not MatchesPlatform("Shopee", "TT"), "TT excludes Shopee"
    RequireTest MatchesPlatform("TikTok", "TT") And MatchesPlatform("Facebook", "TT"), "TT includes other platforms"
    RequireTest MatchesPlatform("Shopee", "ALL") And MatchesPlatform("TikTok", "ALL") And _
                MatchesPlatform("Facebook", "ALL") And MatchesPlatform("", "ALL"), "ALL includes every platform"
    RequireTest MatchesPlatform(CVErr(xlErrNA), "ALL"), "ALL does not inspect platform values"

    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "YES", "SP translation"
    RequireTest norm.Cells(norm.Rows.Count, "A").End(xlUp).Row = 3, "Duplicate order+SKU grouped; different SKU kept"
    RequireTest norm.Range("H2").Value2 = 3, "Explicit quantity equals grouped quantity"
    RequireTest norm.Range("F2").Value2 = "00012345678901234567", "Long OID and leading zeros"
    RequireTest norm.Range("B2").Value2 = text, "Unicode unchanged"
    LoadData_ToSKUSheet wb
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "YES", "SP batch load"
    RequireTest wb.Worksheets("TEST-A").Range("A3").Value2 = text, "Sets wrap to next row"
    RequireTest wb.Worksheets("TEST-A").Range("D3").Value2 = "", "Partial final row"
    RequireTest wb.Worksheets("TEST-A").Range("B2").Value2 = "00012345678901234567", "Output plain-text OID"
    RequireTest wb.Worksheets("TEST-A").Range("H2").Value2 = "SP", "Output platform"
    RequireTest wb.Worksheets("TEST-B").Range("A2").Value2 = "BETA" And _
                wb.Worksheets("TEST-B").Range("B2").Value2 = "BETA" And _
                wb.Worksheets("TEST-B").Range("C2").Value2 = "GAMMA", "Per-SKU copy settings"

    ' A different I3 cannot consume the previous batch.
    inputSheet.Range("I3").Value2 = "TT"
    cfg = ReadConfig(wb)
    Set skus = NormalizedSkus(norm)
    On Error Resume Next
    RequireBatch wb, cfg, norm, skus, True
    blocked = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    RequireTest blocked, "Changed platform blocks old batch"
    TranslateRawInput_ToNormalizedRows wb
    RequireTest norm.Cells(norm.Rows.Count, "A").End(xlUp).Row = 3, "TT contains TikTok and Facebook only"
    RequireTest norm.Range("F2").Value2 = "TT-001" And norm.Range("H2").Value2 = 2, "No explicit quantity uses Qty"
    LoadData_ToSKUSheet wb
    RequireTest wb.Worksheets("TEST-B").Range("A2").Value2 = "DELTA", "Facebook routed to its SKU"

    inputSheet.Range("I3").Value2 = " all "
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "YES", "ALL translation"
    RequireTest norm.Cells(norm.Rows.Count, "A").End(xlUp).Row = 5, "ALL contains SP and TT groups"
    LoadData_ToSKUSheet wb
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "YES", "ALL batch load"
    RequireTest wb.Worksheets("TEST-A").Range("A4").Value2 = "ALPHA", "ALL retains TikTok after Shopee"
    RequireTest wb.Worksheets("TEST-B").Range("A2").Value2 = "BETA" And _
                wb.Worksheets("TEST-B").Range("A3").Value2 = "DELTA", "ALL loads Shopee and Facebook"
    RequireTest wb.Worksheets("TEST-A").Range("H2").Value2 = "ALL", "ALL output platform"

    ' One invalid destination must leave every SKU's old output untouched.
    wb.Worksheets("TEST-A").Range("A2").Value2 = "KEEP-A"
    wb.Worksheets("TEST-B").Range("D1").Value2 = "BROKEN-DATE"
    LoadData_ToSKUSheet wb
    RequireTest wb.Worksheets("TEST-A").Range("A2").Value2 = "KEEP-A", "Batch destination preflight"
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "NO", "Failed load blocks XML"
    wb.Worksheets("TEST-B").Range("D1").Value2 = "DATE"

    ' Mismatching explicit quantity and conflicting duplicate raw text.
    inputSheet.Range("I3").Value2 = "SP"
    inputSheet.Range("D6").Value2 = 1
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "NO", "Explicit quantity mismatch blocks batch"
    RequireTest inputSheet.Range("A5").Interior.Color = RGB(255, 199, 206), "Mismatch highlights source group"
    inputSheet.Range("D6").Value2 = 2
    inputSheet.Range("E6").Value2 = "TEST-A *3pcs || DIFFERENT"
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "NO", "Conflicting raw input blocks batch"
    inputSheet.Range("E6").Value2 = inputSheet.Range("E5").Value2

    ' Missing and duplicate Database SKUs fail before replacing old normalized rows.
    inputSheet.Range("A5:A6").Value2 = "MISSING-X"
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "NO", "Missing Database SKU blocks batch"
    inputSheet.Range("A5:A6").Value2 = "TEST-A"
    WriteTestRow db, 4, Array("TEST-A", 2, 1, 0, 0, 0, 1, 1)
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "NO", "Duplicate Database SKU blocks batch"
    db.Rows(4).ClearContents

    ' Final valid SP batch remains available for manual XML and Ribbon tests.
    TranslateRawInput_ToNormalizedRows wb
    LoadData_ToSKUSheet wb
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "YES", "Final valid batch"
    wb.Activate
    MsgBox "Batch assertions passed. Synthetic workbook left open for XML export checks. No file was saved.", vbInformation
    Exit Sub
Failed:
    failure = Err.Description
    MsgBox "BATCH TEST FAILED: " & failure, vbCritical
    Err.Raise vbObjectError + 2400, "GC_TestBatchProcessing", failure
End Sub

Private Sub WriteTestRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal values As Variant)
    Dim i As Long
    For i = 0 To UBound(values)
        If rowIndex >= 5 And i = 1 Then ws.Cells(rowIndex, i + 1).NumberFormat = "@"
        ws.Cells(rowIndex, i + 1).Value2 = values(i)
    Next i
End Sub

Private Sub MakeTestOutput(ByVal wb As Workbook, ByVal sku As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sku
    WriteTestRow ws, 1, headers
End Sub

Private Sub RequireTest(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2401, "Batch assertion", message
End Sub
