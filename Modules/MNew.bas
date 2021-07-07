Attribute VB_Name = "MNew"
Option Explicit

Public Function CMatOp(ByVal p As Long, s As String, Optional ByVal nk As Byte) As CMatOp
    Set CMatOp = New CMatOp: CMatOp.New_ p, s, nk
End Function

Public Function CMatOpByVal(ByVal pDbl As Long, Optional ByVal mRows As Byte = 1, Optional ByVal nCols As Byte = 1) As CMatOp
    Set CMatOpByVal = New CMatOp: CMatOpByVal.NewByVal pDbl, mRows, nCols
End Function

Public Function CMatOpByValQ(ByVal pDbl As Long, Optional ByVal maxRowsCols As Byte = 1) As CMatOp
    Set CMatOpByValQ = New CMatOp: CMatOpByValQ.NewByVal pDbl, maxRowsCols, maxRowsCols
End Function

Function StdFont(ByVal Name As String, ByVal Size As Single, _
                 Optional ByVal bBold As Boolean = False, _
                 Optional ByVal bItalic As Boolean = False, _
                 Optional ByVal aWeight As Long = 1, _
                 Optional ByVal aCharSet As Long = 1, _
                 Optional ByVal bUnderline As Boolean = False, _
                 Optional ByVal bStrikeThru As Boolean = False) As StdFont
    Set StdFont = New StdFont
    With StdFont
        .Name = Name
        .Size = Size
        .Bold = bBold
        .Italic = bItalic
        .Weight = aWeight
        .Charset = aCharSet
        .Underline = bUnderline
        .Strikethrough = bStrikeThru
    End With
End Function

Function StdFontC(ByVal other As StdFont) As StdFont
    With other
        Set StdFontC = MNew.StdFont(.Name, .Size, .Bold, .Italic, .Weight, .Charset, .Underline, .Strikethrough)
    End With
End Function

Public Sub StdFont_Assign(fntDst As StdFont, fntSrc As StdFont)
    With fntDst
        .Name = fntSrc.Name
        .Size = fntSrc.Size
        .Bold = fntSrc.Bold
        .Italic = fntSrc.Italic
        .Weight = fntSrc.Weight
        .Charset = fntSrc.Charset
        .Underline = fntSrc.Underline
        .Strikethrough = fntSrc.Strikethrough
    End With
End Sub

Public Function Splitter(BolMDI As Boolean, MyOwner As Object, MyContainer As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ BolMDI, MyOwner, MyContainer, Name, LeftTop, RghtBot
End Function


