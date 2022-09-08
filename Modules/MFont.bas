Attribute VB_Name = "MFont"
Option Explicit

'Extensions to the class VBA.StdFont
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
