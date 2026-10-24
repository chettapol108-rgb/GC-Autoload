Public Function CODE128_LIBRE(ByVal inputValue As Variant) As String
    Dim strInput As String
    Dim codes() As Long
    Dim codeCount As Long
    Dim segments() As String
    Dim segmentTypes() As String
    Dim segmentCount As Long
    Dim currentMode As String
    Dim currentSegment As String
    Dim i As Long, j As Long
    Dim ch As String
    Dim isDigit As Boolean
    Dim useCodeB As Boolean
    Dim currentCodeSet As String
    Dim digits As String
    Dim evenLength As Long
    Dim pair As String
    Dim checksum As Long
    Dim result As String
    Dim c As Long
    
    Const START_B As Long = 104
    Const START_C As Long = 105
    Const CODE_B As Long = 100
    Const CODE_C As Long = 99
    Const STOP_CODE As Long = 106
    
    If IsError(inputValue) Then
        CODE128_LIBRE = ""
        Exit Function
    End If
        
    If IsNull(inputValue) Or IsEmpty(inputValue) Then
        CODE128_LIBRE = ""
        Exit Function
    End If
    
    strInput = CStr(inputValue)
    
    If Len(strInput) = 0 Then
        CODE128_LIBRE = ""
        Exit Function
    End If
    
    ' Split input into digit and non-digit segments
    currentMode = ""
    currentSegment = ""
    segmentCount = 0
    
    For i = 1 To Len(strInput)
        ch = Mid$(strInput, i, 1)
        isDigit = (ch Like "#")
        
        If currentMode = "" Then
            currentMode = IIf(isDigit, "digit", "char")
            currentSegment = ch
        ElseIf (currentMode = "digit" And isDigit) Or (currentMode = "char" And Not isDigit) Then
            currentSegment = currentSegment & ch
        Else
            segmentCount = segmentCount + 1
            ReDim Preserve segments(1 To segmentCount)
            ReDim Preserve segmentTypes(1 To segmentCount)
            segments(segmentCount) = currentSegment
            segmentTypes(segmentCount) = currentMode
            
            currentMode = IIf(isDigit, "digit", "char")
            currentSegment = ch
        End If
    Next i
    
    If Len(currentSegment) > 0 Then
        segmentCount = segmentCount + 1
        ReDim Preserve segments(1 To segmentCount)
        ReDim Preserve segmentTypes(1 To segmentCount)
        segments(segmentCount) = currentSegment
        segmentTypes(segmentCount) = currentMode
    End If
    
    ' Choose starting code set based on the first segment
    useCodeB = True
    If segmentCount > 0 Then
        If segmentTypes(1) = "digit" And Len(segments(1)) >= 4 Then
            useCodeB = False
        End If
    End If
    
    codeCount = 0
    AddCode codes, codeCount, IIf(useCodeB, START_B, START_C)
    currentCodeSet = IIf(useCodeB, "B", "C")
    
    ' Process each segment
    For i = 1 To segmentCount
        If segmentTypes(i) = "digit" And Len(segments(i)) >= 4 Then
            If currentCodeSet <> "C" Then
                AddCode codes, codeCount, CODE_C
                currentCodeSet = "C"
            End If
            
            digits = segments(i)
            evenLength = Len(digits) - (Len(digits) Mod 2)
            
            For j = 1 To evenLength Step 2
                pair = Mid$(digits, j, 2)
                AddCode codes, codeCount, CLng(pair)
            Next j
            
            ' If one digit remains, switch to Code B for the last digit
            If Len(digits) Mod 2 = 1 Then
                AddCode codes, codeCount, CODE_B
                currentCodeSet = "B"
                ch = Right$(digits, 1)
                AddCode codes, codeCount, CodeBValue(ch)
            End If
        Else
            If currentCodeSet <> "B" Then
                AddCode codes, codeCount, CODE_B
                currentCodeSet = "B"
            End If
            
            For j = 1 To Len(segments(i))
                ch = Mid$(segments(i), j, 1)
                AddCode codes, codeCount, CodeBValue(ch)
            Next j
        End If
    Next i
    
    ' Calculate checksum
    checksum = codes(1)
    For i = 2 To codeCount
        checksum = checksum + (codes(i) * (i - 1))
    Next i
    checksum = checksum Mod 103
    
    AddCode codes, codeCount, checksum
    AddCode codes, codeCount, STOP_CODE
    
    ' Convert code values to Libre Barcode 128 glyph string
    result = ""
    For i = 1 To codeCount
        c = codes(i)
        If c < 95 Then
            result = result & ChrW$(c + 32)
        Else
            result = result & ChrW$(c + 100)
        End If
    Next i
    
    CODE128_LIBRE = result
End Function

Private Sub AddCode(ByRef codes() As Long, ByRef codeCount As Long, ByVal value As Long)
    codeCount = codeCount + 1
    ReDim Preserve codes(1 To codeCount)
    codes(codeCount) = value
End Sub

Private Function CodeBValue(ByVal ch As String) As Long
    Dim asciiValue As Long
    
    asciiValue = AscW(ch)
    
    If asciiValue < 32 Or asciiValue > 126 Then
        Err.Raise vbObjectError + 513, "CODE128_LIBRE", _
                  "Unsupported character for Code Set B: " & ch
    End If
    
    CodeBValue = asciiValue - 32
End Function