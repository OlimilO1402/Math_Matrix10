Attribute VB_Name = "MMain"
Option Explicit
Public Type FrmCreateSetting
    mRows   As String 'number of rows    (Zeilen)
    nCols   As String 'number of columns (Spalten)
    Format  As String 'number of digits after period
    Values  As String '
    NumFrom As String 'min-value
    NumTo   As String 'max-value
    MatVec  As String 'contains the matrix as string
End Type

Private m_MyFont As StdFont
'zur Erzeugung von Zufallszahlen:
Dim m_dblFrom As Double
Dim m_dblDist As Double

Sub Main()
    Set MyFont = MNew.StdFont("Consolas", 10#)
    FrmMain.Show
End Sub
Public Function New_FrmCreateSetting(ByVal aRows As Long, ByVal aCols As Long) As FrmCreateSetting
    With New_FrmCreateSetting
        .mRows = aRows
        .nCols = aCols
        '.Format
        '.Values
        '.NumFrom
        '.NumTo
        '.MatVec
    End With
End Function
Public Function FrmCreateSetting_IsEmpty(this As FrmCreateSetting) As Boolean
    Dim b As Boolean
    With this
        b = Len(.mRows) = 0:   If Not b Then Exit Function
        b = Len(.nCols) = 0:   If Not b Then Exit Function
        b = Len(.Format) = 0:  If Not b Then Exit Function
        b = Len(.Values) = 0:  If Not b Then Exit Function
        b = Len(.NumFrom) = 0: If Not b Then Exit Function
        b = Len(.NumTo) = 0:   If Not b Then Exit Function
        b = Len(.MatVec) = 0:  If Not b Then Exit Function
    End With
    FrmCreateSetting_IsEmpty = True
End Function
Private Sub FrmCreateSetting_Assign(this As FrmCreateSetting, other As FrmCreateSetting)
    With other
        If Len(.mRows) Then this.mRows = .mRows
        If Len(.nCols) Then this.nCols = .nCols
        If Len(.Format) Then this.Format = .Format
        If Len(.Values) Then this.Values = .Values
        If Len(.NumFrom) Then this.NumFrom = .NumFrom
        If Len(.NumTo) Then this.NumTo = .NumTo
        If Len(.MatVec) Then this.MatVec = .MatVec
    End With
End Sub



Public Property Set MyFont(Value As StdFont)
    Set m_MyFont = Value
End Property
Public Property Get MyFont() As StdFont
    Set MyFont = m_MyFont
End Property

'Hilfsfunktionen
Public Sub Rnd_SetFromTo(ByVal dblFrom As Double, ByVal dblTo As Double)
    Randomize
    m_dblFrom = Min(dblFrom, dblTo)
    dblTo = Max(dblFrom, dblTo)
    m_dblDist = Abs(dblTo - m_dblFrom)
End Sub

Public Function rv() As Double
    rv = m_dblFrom + Rnd * m_dblDist
End Function

Public Function GetNextRandom() As Double
    GetNextRandom = rv
End Function

Public Function DeleteMultiWS(s As String) As String
    DeleteMultiWS = Trim$(s)
    If InStr(1, s, "  ") = 0 Then Exit Function
    DeleteMultiWS = Replace(s, "  ", " ")
    DeleteMultiWS = DeleteMultiWS(DeleteMultiWS)
End Function
Public Function DeleteCRLF(s As String) As String
    DeleteCRLF = Trim$(s)
    If InStr(1, s, vbLf) = 0 Then Exit Function
    If InStr(1, s, vbCr) = 0 Then Exit Function
    DeleteCRLF = Replace(Replace(Replace(s, vbCrLf, " "), vbLf, " "), vbCr, " ")
    DeleteCRLF = DeleteCRLF(DeleteCRLF)
End Function

Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function
Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
Public Function TStr(d As Double) As String
    TStr = Trim$(Str$(d))
End Function
Public Function DblParse(ByVal s As String) As Double
Try: On Error GoTo Catch
    s = Replace(Trim$(s), ",", ".")
    DblParse = Val(s)
Catch: 'out
End Function

Public Function Double_TryParse(ByVal s As String, ByRef d_out As Double) As Boolean
Try: On Error GoTo Catch
    s = Replace(Trim$(s), ",", ".")
    d_out = Val(s)
    Double_TryParse = True
Catch: 'out
End Function

Public Function PadLeft(StrVal As String, _
                        ByVal totalWidth As Long, _
                        Optional ByVal paddingChar As String) As String
    ' der String wird mit der angegebenen Länge zurückgegeben, der
    ' String wird nach rechts gerückt, und links mit PadChar aufgefüllt
    ' ist PadChar nicht angegeben, so wird mit RSet der String in
    ' Spaces eingefügt.
    If Len(paddingChar) Then
        If Len(StrVal) <= totalWidth Then _
            PadLeft = StrVal & String$(totalWidth - Len(StrVal), paddingChar)
    Else
        PadLeft = Space$(totalWidth)
        RSet PadLeft = StrVal
    End If
End Function
Public Function PadRight(StrVal As String, _
                         ByVal totalWidth As Long, _
                         Optional ByVal paddingChar As String) As String
    ' der String wird mit der angegebenen Länge zurückgegeben, der
    ' String wird nach links gerückt, und rechts mit PadChar aufgefüllt
    ' ist PadChar nicht angegeben, so wird mit LSet der String in
    ' Spaces eingefügt.
    If Len(paddingChar) Then
        If Len(StrVal) <= totalWidth Then _
            PadRight = StrVal & String$(totalWidth - Len(StrVal), paddingChar)
    Else
        PadRight = Space$(totalWidth)
        LSet PadRight = StrVal
    End If
End Function



'Sicherheitskopie
'Ausgabefunktionen
'Public Function Matrix2_ToStr(m As Matrix2) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb)
''    End With
''    Matrix2_ToStr = s
'    Matrix2_ToStr = MatrixA_ToStr(Matrix2_ToArr(m), 2, 2)
'End Function
'Public Function Matrix3_ToStr(m As Matrix3) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & vbCrLf
''        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc)
''    End With
''    Matrix3_ToStr = s
'    Matrix3_ToStr = MatrixA_ToStr(Matrix3_ToArr(m), 3, 3)
'End Function
'Public Function Matrix4_ToStr(m As Matrix4) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & " " & TStr(.ad) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & " " & TStr(.bd) & vbCrLf
''        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc) & " " & TStr(.cd) & vbCrLf
''        s = s & TStr(.da) & " " & TStr(.db) & " " & TStr(.dc) & " " & TStr(.dd)
''    End With
''    Matrix4_ToStr = s
'    Matrix4_ToStr = MatrixA_ToStr(Matrix4_ToArr(m), 4, 4)
'End Function
'Public Function Matrix5_ToStr(m As Matrix5) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & " " & TStr(.ad) & " " & TStr(.ae) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & " " & TStr(.bd) & " " & TStr(.be) & vbCrLf
''        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc) & " " & TStr(.cd) & " " & TStr(.ce) & vbCrLf
''        s = s & TStr(.da) & " " & TStr(.db) & " " & TStr(.dc) & " " & TStr(.dd) & " " & TStr(.de) & vbCrLf
''        s = s & TStr(.ea) & " " & TStr(.eb) & " " & TStr(.ec) & " " & TStr(.ed) & " " & TStr(.ee)
''    End With
''    Matrix5_ToStr = s
'    Matrix5_ToStr = MatrixA_ToStr(Matrix5_ToArr(m), 5, 5)
'End Function
'Public Function Matrix6_ToStr(m As Matrix6) As String
''    Dim s As String: s = ""
''    With m
''        s = s & TStr(.aa) & " " & TStr(.ab) & " " & TStr(.ac) & " " & TStr(.ad) & " " & TStr(.ae) & " " & TStr(.af) & vbCrLf
''        s = s & TStr(.ba) & " " & TStr(.bb) & " " & TStr(.bc) & " " & TStr(.bd) & " " & TStr(.be) & " " & TStr(.bf) & vbCrLf
''        s = s & TStr(.ca) & " " & TStr(.cb) & " " & TStr(.cc) & " " & TStr(.cd) & " " & TStr(.ce) & " " & TStr(.cf) & vbCrLf
''        s = s & TStr(.da) & " " & TStr(.db) & " " & TStr(.dc) & " " & TStr(.dd) & " " & TStr(.de) & " " & TStr(.df) & vbCrLf
''        s = s & TStr(.ea) & " " & TStr(.eb) & " " & TStr(.ec) & " " & TStr(.ed) & " " & TStr(.ee) & " " & TStr(.ef) & vbCrLf
''        s = s & TStr(.fa) & " " & TStr(.fb) & " " & TStr(.fc) & " " & TStr(.fd) & " " & TStr(.fe) & " " & TStr(.ff)
''    End With
''    Matrix6_ToStr = s
'    Matrix6_ToStr = MatrixA_ToStr(Matrix6_ToArr(m), 6, 6)
'End Function
'
'
''allgemein
''Public Function MatrixT_ToStr(m As MatrixT, ByVal mRows As Long, ByVal nCols As Long) As String
''    MatrixT_ToStr = MatrixA_ToStr(m.a, mRows, nCols)
''End Function
'
'
