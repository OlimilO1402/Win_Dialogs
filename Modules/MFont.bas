Attribute VB_Name = "MFont"
Option Explicit

'Extensions to the class VBA.StdFont
Function New_StdFont(s As String) As StdFont
    Set New_StdFont = New StdFont
    Dim sLines() As String: sLines = Split(s, "{")
    If LCase(Strings.Trim(sLines(0))) <> "stdfont" Then Exit Function
    sLines = Split(sLines(1), vbCrLf)
    Dim i As Long, sElems() As String, sKey As String, sVal As String
    For i = LBound(sLines) + 1 To UBound(sLines) - 1
        sElems() = Split(sLines(i), "=")
        sKey = LCase(Trim(sElems(0)))
        If UBound(sElems) > 0 Then
            sVal = Trim(sElems(1))
            With New_StdFont
                Select Case sKey
                Case "name":                   .Name = sVal
                Case "bold":                   .Bold = Boolean_Parse(sVal)
                Case "charset":             .Charset = CInt(sVal)
                Case "italic":               .Italic = Boolean_Parse(sVal)
                Case "size":                   .Size = CCur(sVal)
                Case "strikethrough": .Strikethrough = Boolean_Parse(sVal)
                Case "underline":         .Underline = Boolean_Parse(sVal)
                Case "weight":               .Weight = CInt(sVal)
                End Select
            End With
        End If
    Next
End Function

Public Function StdFont_Clone(this As StdFont) As StdFont
    Dim DstF As New StdFont: StdFont_Copy DstF, this
    Set StdFont_Clone = DstF 'StdFont_Copy(New StdFont, this)
End Function

Public Sub StdFont_Copy(DstFont As StdFont, SrcFont As StdFont)
    With DstFont
        .Name = SrcFont.Name
        .Size = SrcFont.Size
        .Bold = SrcFont.Bold
        .Italic = SrcFont.Italic
        .Weight = SrcFont.Weight
        .Charset = SrcFont.Charset
        .Underline = SrcFont.Underline
        .Strikethrough = SrcFont.Strikethrough
    End With
End Sub

Public Function StdFont_Equals(this As StdFont, other As StdFont) As Boolean
    Dim b As Boolean
    With this
        b = .Name = other.Name:                   If Not b Then Exit Function
        b = .Size = other.Size:                   If Not b Then Exit Function
        b = .Bold = other.Bold:                   If Not b Then Exit Function
        b = .Italic = other.Italic:               If Not b Then Exit Function
        b = .Weight = other.Weight:               If Not b Then Exit Function
        b = .Charset = other.Charset:             If Not b Then Exit Function
        b = .Underline = other.Underline:         If Not b Then Exit Function
        b = .Strikethrough = other.Strikethrough: If Not b Then Exit Function
    End With
    StdFont_Equals = True
End Function

Public Function StdFont_ToStr(this As StdFont) As String
    Dim s As String: s = "StdFont{" & vbCrLf
    With this
        s = s & "Name:          " & .Name & vbCrLf
        s = s & "Size:          " & .Size & vbCrLf
        s = s & "Bold:          " & .Bold & vbCrLf
        s = s & "Italic:        " & .Italic & vbCrLf
        s = s & "Weight:        " & .Weight & vbCrLf
        s = s & "Charset:       " & .Charset & vbCrLf
        s = s & "Underline:     " & .Underline & vbCrLf
        s = s & "Strikethrough: " & .Strikethrough & vbCrLf
    End With
    StdFont_ToStr = s & "}"
End Function

Function Boolean_Parse(ByVal sVal As String) As Boolean
    Dim b As Boolean
    sVal = LCase(sVal)
    Select Case True
    Case sVal = "falsch": b = False
    Case sVal = "false":  b = False
    Case sVal = "nein":   b = False
    Case sVal = "no":     b = False
    Case sVal = "wahr":   b = True
    Case sVal = "true":   b = True
    Case sVal = "yes":    b = True
    Case sVal = "ja":     b = True
    End Select
    Boolean_Parse = b
End Function

