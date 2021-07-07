VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMatOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dim scal As Double 'Scalar
'
'Dim V2  As Vector2
'Dim V3  As Vector3
'Dim V4  As Vector4
'Dim V5  As Vector5
'Dim V6  As Vector6
'Dim V7  As Vector7
'Dim V8  As Vector8
'Dim V9  As Vector9
'Dim V10 As Vector10
'
'Dim M2  As Matrix2
'Dim M3  As Matrix3
'Dim M4  As Matrix4
'Dim M5  As Matrix5
'Dim M6  As Matrix6
'Dim M7  As Matrix7
'Dim M8  As Matrix8
'Dim M9  As Matrix9
'Dim M10 As Matrix10
'
Dim m_mRows As Byte 'number of rows in the matrix from top to down
Dim m_nCols As Byte 'number of columns in the matrix from left to right
Dim m_nk    As Byte 'number of digits after "."
Dim vm  As TVecMatUnion 'virtual vector matrix
Dim mem As Matrix10 'could either be used or not

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef Src As Any, ByVal bytLen As Long)

Private Sub Class_Initialize()
    'New_TVecMatUnion mv
End Sub

Private Sub Class_Terminate()
    Del_TVecMatUnion vm
End Sub
Friend Sub New_(ByVal p As Long, s As String, Optional ByVal nk As Byte) 'ByVal mRows As Byte, ByVal nCols As Byte,
    
    'wenn p = 0 dann hier selber Array anlegen
    'oder den p mit geben, dann wird das vm.VMArr verwendet
    'sollte man vm.VMArr wirklich verwenden
    'oder sollen wir hier drin ein Array anlegen?
    'OK, aber dann das vm.VMarr wieder entfernen!
    'OK das vm.VMArr word zum Skalar der bis jetzt noch fehlt!
    'dan ist der Scalar ein .scal(0) as Double
    'vm.pudt.pvData = p
    'hja wohin eigentlich?
    'eigentlich sollten wir die Daten schlicht nur in ein Array schieben
    m_nk = nk
    If p = 0 Then p = VarPtr(mem)
    New_TVecMatUnion vm, p
    
    Parse s
End Sub
'Friend Sub NewC(other As CMatOp)
'    m_mRows = other.mRows
'    m_nCols = other.nCols
'    m_nk = other.nk
'    'vm = other.
'End Sub
Friend Sub NewByVal(ByVal pDbl As Long, Optional ByVal mRows As Byte = 1, Optional ByVal nCols As Byte = 1)
    New_TVecMatUnion vm, VarPtr(mem)
    'If mRows = 0 Then mRows = 1
    'If nCols = 0 Then nCols = 1
    m_mRows = mRows: m_nCols = nCols
    RtlMoveMemory mem, ByVal pDbl, m_mRows * m_nCols * 8
End Sub

'Public Function Clone() As CMatOp
'    set clone = mnew.CMatOp(
'End Function
Public Sub Parse(s As String)
    Dim maxrc As Long: maxrc = GetRowsCols(s, m_mRows, m_nCols)
    If m_mRows > 10 Or m_nCols > 10 Then MsgBox "Max 10 dimensions allowed!": Exit Sub

    Select Case True
    Case IsScalar
        Dim d As Double
        If Double_TryParse(s, d) Then vm.Scalar(0) = d
    Case IsVector
        Select Case maxrc
        Case 2:  vm.AVec2(0) = Vec2_Parse(s)
        Case 3:  vm.AVec3(0) = Vec3_Parse(s)
        Case 4:  vm.AVec4(0) = Vec4_Parse(s)
        Case 5:  vm.AVec5(0) = Vec5_Parse(s)
        Case 6:  vm.AVec6(0) = Vec6_Parse(s)
        Case 7:  vm.AVec7(0) = Vec7_Parse(s)
        Case 8:  vm.AVec8(0) = Vec8_Parse(s)
        Case 9:  vm.AVec9(0) = Vec9_Parse(s)
        Case 10: vm.AVec10(0) = Vec10_Parse(s)
        End Select
    Case IsMatrix
        Select Case maxrc
        Case 2:  vm.AMat2(0) = Mat2_Parse(s)
        Case 3:  vm.AMat3(0) = Mat3_Parse(s)
        Case 4:  vm.AMat4(0) = Mat4_Parse(s)
        Case 5:  vm.AMat5(0) = Mat5_Parse(s)
        Case 6:  vm.AMat6(0) = Mat6_Parse(s)
        Case 7:  vm.AMat7(0) = Mat7_Parse(s)
        Case 8:  vm.AMat8(0) = Mat8_Parse(s)
        Case 9:  vm.AMat9(0) = Mat9_Parse(s)
        Case 10: vm.AMat10(0) = Mat10_Parse(s)
        End Select
    End Select
End Sub

Public Property Let nk(ByVal Value As Byte)
    m_nk = Value
End Property
Public Property Get nk() As Byte
    nk = m_nk
End Property

Public Property Get IsScalar() As Boolean
    IsScalar = (m_mRows = 1) And (m_nCols = m_mRows)
End Property
Public Property Get IsRowVec() As Boolean
    IsRowVec = (m_mRows = 1) And (m_nCols > 0)
End Property
Public Property Get IsVector() As Boolean
    IsVector = (m_nCols <> m_mRows) And ((m_mRows = 1) Or (m_nCols = 1))
End Property
Public Property Get IsMatrix() As Boolean
    IsMatrix = Not IsVector
End Property
Public Property Get mRows() As Byte
    mRows = m_mRows
End Property
Public Property Get nCols() As Byte
    nCols = m_nCols
End Property
Public Property Get MaxRows(other As CMatOp) As Byte
    MaxRows = Max(m_mRows, other.mRows)
End Property
Public Property Get MaxCols(other As CMatOp) As Byte
    MaxCols = Max(m_nCols, other.nCols)
End Property
Public Property Get maxRowsCols(Optional other As CMatOp) As Byte
    If Not other Is Nothing Then maxRowsCols = other.maxRowsCols
    maxRowsCols = Max(maxRowsCols, Max(m_mRows, m_nCols))
End Property

Function GetRowsCols(T As String, ByRef rows_out As Byte, ByRef cols_out As Byte) As Byte
    rows_out = GetRows(T):    cols_out = GetCols(T)
    GetRowsCols = Max(rows_out, cols_out)
End Function
Function GetRows(T As String) As Long
    Dim s As String: s = DeleteMultiWS(T)
    Dim sa() As String: sa = Split(s, vbCrLf)
    Dim i As Long
    GetRows = UBound(sa) + 1
    For i = UBound(sa) To 0 Step -1
        If Len(sa(i)) = 0 Then GetRows = GetRows - 1 Else Exit For
    Next
End Function
Function GetCols(T As String) As Long
    Dim s As String: s = DeleteMultiWS(T)
    Dim sa() As String: sa = Split(s, vbCrLf)
    Dim i As Long
    For i = 0 To UBound(sa)
        GetCols = Max(GetCols, UBound(Split(Trim(sa(i)), " ")) + 1)
    Next
End Function

Friend Function ToStr(Optional ByVal dFormat As Integer = -1) As String
    Dim s As String
    Dim maxrc As Long: maxrc = Max(m_mRows, m_nCols)
    If dFormat < 0 Then dFormat = m_nk
    Select Case True
    Case IsScalar
        s = CStr(vm.Scalar(0))
    Case IsVector
        Select Case maxrc
        Case 2:  s = Vec2_ToStr(vm.AVec2(0), Me.IsRowVec, dFormat)
        Case 3:  s = Vec3_ToStr(vm.AVec3(0), Me.IsRowVec, dFormat)
        Case 4:  s = Vec4_ToStr(vm.AVec4(0), Me.IsRowVec, dFormat)
        Case 5:  s = Vec5_ToStr(vm.AVec5(0), Me.IsRowVec, dFormat)
        Case 6:  s = Vec6_ToStr(vm.AVec6(0), Me.IsRowVec, dFormat)
        Case 7:  s = Vec7_ToStr(vm.AVec7(0), Me.IsRowVec, dFormat)
        Case 8:  s = Vec8_ToStr(vm.AVec8(0), Me.IsRowVec, dFormat)
        Case 9:  s = Vec9_ToStr(vm.AVec9(0), Me.IsRowVec, dFormat)
        Case 10: s = Vec10_ToStr(vm.AVec10(0), Me.IsRowVec, dFormat)
        End Select
    Case IsMatrix
        Select Case maxrc
        Case 2:  s = Mat2_ToStr(vm.AMat2(0), m_mRows, m_nCols, dFormat)
        Case 3:  s = Mat3_ToStr(vm.AMat3(0), m_mRows, m_nCols, dFormat)
        Case 4:  s = Mat4_ToStr(vm.AMat4(0), m_mRows, m_nCols, dFormat)
        Case 5:  s = Mat5_ToStr(vm.AMat5(0), m_mRows, m_nCols, dFormat)
        Case 6:  s = Mat6_ToStr(vm.AMat6(0), m_mRows, m_nCols, dFormat)
        Case 7:  s = Mat7_ToStr(vm.AMat7(0), m_mRows, m_nCols, dFormat)
        Case 8:  s = Mat8_ToStr(vm.AMat8(0), m_mRows, m_nCols, dFormat)
        Case 9:  s = Mat9_ToStr(vm.AMat9(0), m_mRows, m_nCols, dFormat)
        Case 10: s = Mat10_ToStr(vm.AMat10(0), m_mRows, m_nCols, dFormat)
        End Select
    End Select
    ToStr = s
End Function
Public Property Get Row(ByVal Index As Byte) As CMatOp
    'Index = 0-basiert
    If m_mRows < (Index + 1) Then Exit Property
    Dim maxrc As Byte: maxrc = Max(m_mRows, m_nCols)
    If Me.IsVector Then
        If Me.IsRowVec Then
            Select Case maxrc
            Case 2:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec2(0)), 1, m_nCols)
            Case 3:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec3(0)), 1, m_nCols)
            Case 4:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec4(0)), 1, m_nCols)
            Case 5:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec5(0)), 1, m_nCols)
            Case 6:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec6(0)), 1, m_nCols)
            Case 7:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec7(0)), 1, m_nCols)
            Case 8:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec8(0)), 1, m_nCols)
            Case 9:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec9(0)), 1, m_nCols)
            Case 10: Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0)), 1, m_nCols)
            End Select
        Else
            Select Case Index
            Case 0:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).a), 1, 1)
            Case 1:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).b), 1, 1)
            Case 2:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).c), 1, 1)
            Case 3:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).d), 1, 1)
            Case 4:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).e), 1, 1)
            Case 5:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).f), 1, 1)
            Case 6:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).g), 1, 1)
            Case 7:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).H), 1, 1)
            Case 8:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).i), 1, 1)
            Case 9:  Set Row = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).j), 1, 1)
            End Select
        End If
    Else
        Select Case maxrc
        Case 2:  Set Row = MNew.CMatOpByVal(VarPtr(Mat2_Row(vm.AMat2(0), Index)), 1, m_nCols)
        Case 3:  Set Row = MNew.CMatOpByVal(VarPtr(Mat3_Row(vm.AMat3(0), Index)), 1, m_nCols)
        Case 4:  Set Row = MNew.CMatOpByVal(VarPtr(Mat4_Row(vm.AMat4(0), Index)), 1, m_nCols)
        Case 5:  Set Row = MNew.CMatOpByVal(VarPtr(Mat5_Row(vm.AMat5(0), Index)), 1, m_nCols)
        Case 6:  Set Row = MNew.CMatOpByVal(VarPtr(Mat6_Row(vm.AMat6(0), Index)), 1, m_nCols)
        Case 7:  Set Row = MNew.CMatOpByVal(VarPtr(Mat7_Row(vm.AMat7(0), Index)), 1, m_nCols)
        Case 8:  Set Row = MNew.CMatOpByVal(VarPtr(Mat8_Row(vm.AMat8(0), Index)), 1, m_nCols)
        Case 9:  Set Row = MNew.CMatOpByVal(VarPtr(Mat9_Row(vm.AMat9(0), Index)), 1, m_nCols)
        Case 10: Set Row = MNew.CMatOpByVal(VarPtr(Mat10_Row(vm.AMat10(0), Index)), 1, m_nCols)
        End Select
    End If
End Property
Public Property Get Col(ByVal Index As Byte) As CMatOp
    'Index = 0-basiert
    If m_nCols < (Index + 1) Then Exit Property
    Dim maxrc As Byte: maxrc = Max(m_mRows, m_nCols)
    If Me.IsVector Then
        If Me.IsRowVec Then
            Select Case Index
            Case 0:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).a), 1, 1)
            Case 1:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).b), 1, 1)
            Case 2:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).c), 1, 1)
            Case 3:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).d), 1, 1)
            Case 4:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).e), 1, 1)
            Case 5:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).f), 1, 1)
            Case 6:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).g), 1, 1)
            Case 7:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).H), 1, 1)
            Case 8:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).i), 1, 1)
            Case 9:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0).j), 1, 1)
            End Select
        Else
            Select Case maxrc
            Case 2:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec2(0)), m_mRows, 1)
            Case 3:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec3(0)), m_mRows, 1)
            Case 4:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec4(0)), m_mRows, 1)
            Case 5:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec5(0)), m_mRows, 1)
            Case 6:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec6(0)), m_mRows, 1)
            Case 7:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec7(0)), m_mRows, 1)
            Case 8:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec8(0)), m_mRows, 1)
            Case 9:  Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec9(0)), m_mRows, 1)
            Case 10: Set Col = MNew.CMatOpByVal(VarPtr(vm.AVec10(0)), m_mRows, 1)
            End Select
        End If
    Else
        Select Case maxrc
        Case 2:  Set Col = MNew.CMatOpByVal(VarPtr(Mat2_Col(vm.AMat2(0), Index)), m_mRows, 1)
        Case 3:  Set Col = MNew.CMatOpByVal(VarPtr(Mat3_Col(vm.AMat3(0), Index)), m_mRows, 1)
        Case 4:  Set Col = MNew.CMatOpByVal(VarPtr(Mat4_Col(vm.AMat4(0), Index)), m_mRows, 1)
        Case 5:  Set Col = MNew.CMatOpByVal(VarPtr(Mat5_Col(vm.AMat5(0), Index)), m_mRows, 1)
        Case 6:  Set Col = MNew.CMatOpByVal(VarPtr(Mat6_Col(vm.AMat6(0), Index)), m_mRows, 1)
        Case 7:  Set Col = MNew.CMatOpByVal(VarPtr(Mat7_Col(vm.AMat7(0), Index)), m_mRows, 1)
        Case 8:  Set Col = MNew.CMatOpByVal(VarPtr(Mat8_Col(vm.AMat8(0), Index)), m_mRows, 1)
        Case 9:  Set Col = MNew.CMatOpByVal(VarPtr(Mat9_Col(vm.AMat9(0), Index)), m_mRows, 1)
        Case 10: Set Col = MNew.CMatOpByVal(VarPtr(Mat10_Col(vm.AMat10(0), Index)), m_mRows, 1)
        End Select
    End If
End Property


'Op unär oder binär
'
Public Function op(ByVal aOp As String, Optional other1 As CMatOp) As CMatOp
'Hja, das ist hier die Frage, sollte man die Operationen nur in der Ergebnismatrix machen
'oder in einem der beiden anderen
'so dass man die Ergebismatrix anlegt un die Operation findet dann auf die privaten Elemente der Matrix statt
'Nö, es sollte eigentlich egal sein wierum
'OK wir gehen den Weg dass die Ergebnismatrix erst von der Operation erzeugt wird
'so könnte die Operation auch eine externe Klasse sein.

    Select Case LCase(Trim(aOp))
    Case "+" 'Addiere
        Set op = Me.Add(other1)
    Case "-" 'Subtrahiere
        Set op = Me.Subt(other1)
    Case "*" 'Multipliziere
        Set op = Me.Mult(other1)
    Case "row"
        Set op = Me.Row(CByte(other1.ToScalar))
    Case "col"
        Set op = Me.Col(CByte(other1.ToScalar))
    Case "tra" 'Transponiere
        Set op = Me.Transp
    Case "inv" 'Invertiere
        
    'Case "rnd"
    '    Set op = MMatrices.Mat2_Rnd
    'Case "unit"
    'Case "ones"
    End Select
End Function
Public Function Rnd(ByVal mRows As Byte, ByVal nCols As Byte, ByVal dbl_From As Double, ByVal dbl_To As Double) As CMatOp
    Dim maxrc As Byte: maxrc = Max(mRows, nCols)
    Select Case maxrc
    Case 1:  Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat2_Rnd(dbl_From, dbl_To)), 1, 1)
    Case 2:  Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat2_Rnd(dbl_From, dbl_To)), mRows, nCols)
    Case 3:  Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat3_Rnd(dbl_From, dbl_To)), mRows, nCols)
    Case 4:  Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat4_Rnd(dbl_From, dbl_To)), mRows, nCols)
    Case 5:  Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat5_Rnd(dbl_From, dbl_To)), mRows, nCols)
    Case 6:  Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat6_Rnd(dbl_From, dbl_To)), mRows, nCols)
    Case 7:  Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat7_Rnd(dbl_From, dbl_To)), mRows, nCols)
    Case 8:  Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat8_Rnd(dbl_From, dbl_To)), mRows, nCols)
    Case 9:  Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat9_Rnd(dbl_From, dbl_To)), mRows, nCols)
    Case 10: Set Rnd = MNew.CMatOpByVal(VarPtr(MMatrices.Mat10_Rnd(dbl_From, dbl_To)), mRows, nCols)
    End Select
End Function
Public Function Add(other As CMatOp) As CMatOp
    'Matrizenaddition oder Vektoraddition
    Dim maxrc As Byte: maxrc = Me.maxRowsCols(other)
    Dim MaxMRows As Byte: MaxMRows = Me.MaxRows(other)
    Dim MaxNCols As Byte: MaxNCols = Me.MaxCols(other)

    Select Case maxrc
    Case 1
        'nur 2 Skalare
        Dim result As Double: result = vm.Scalar(0) + other.ToScalar
        Set Add = CMatOpByVal(VarPtr(result))
    Case Else
        Select Case True
        Case Me.IsScalar
            If other.IsVector Then
                'aus dem skalar einen Vektor machen?
                'oder einfach auf das erste Element draufaddieren
            Else
                
            End If
        Case Me.IsVector
' 1  2  3     1  2  3     2  4  6
'          +  4  5  6  =  4  5  6
'             7  8  9     7  8  9
'
' 1           1  2  3     2  2  3
' 2        +  4  5  6  =  6  5  6
' 3           7  8  9     10 8  9
'
' 1  2  3     1  2     2  4  6
'          +  3  4  =  4  5  6
'             5  6     7  8  9
'
' 1           1  2     2  2
' 2        +  3  4  =  5  4
' 3           5  6     8  6

            Select Case maxrc
            
            End Select
        Case Me.IsMatrix
            Select Case maxrc
            Case 2:  Set Add = CMatOpByVal(VarPtr(Mat2_add(vm.AMat2(0), other.ToMat2)), MaxMRows, MaxNCols)
            Case 3:  Set Add = CMatOpByVal(VarPtr(Mat3_add(vm.AMat3(0), other.ToMat3)), MaxMRows, MaxNCols)
            Case 4:  Set Add = CMatOpByVal(VarPtr(Mat4_add(vm.AMat4(0), other.ToMat4)), MaxMRows, MaxNCols)
            Case 5:  Set Add = CMatOpByVal(VarPtr(Mat5_add(vm.AMat5(0), other.ToMat5)), MaxMRows, MaxNCols)
            Case 6:  Set Add = CMatOpByVal(VarPtr(Mat6_add(vm.AMat6(0), other.ToMat6)), MaxMRows, MaxNCols)
            Case 7:  Set Add = CMatOpByVal(VarPtr(Mat7_add(vm.AMat7(0), other.ToMat7)), MaxMRows, MaxNCols)
            Case 8:  Set Add = CMatOpByVal(VarPtr(Mat8_add(vm.AMat8(0), other.ToMat8)), MaxMRows, MaxNCols)
            Case 9:  Set Add = CMatOpByVal(VarPtr(Mat9_add(vm.AMat9(0), other.ToMat9)), MaxMRows, MaxNCols)
            Case 10: Set Add = CMatOpByVal(VarPtr(Mat10_add(vm.AMat10(0), other.ToMat10)), MaxMRows, MaxNCols)
            End Select
        End Select
    End Select
End Function
Public Function Subt(other As CMatOp) As CMatOp
    'Matrizensubtraktion oder Vektorsubtraktion
    
End Function
Public Function Mult(other As CMatOp) As CMatOp
    'entw Skalarmultiplikation
    'oder Kreuzprodukt
    'oder Matrizenmultiplikation
    'oder Vektormultiplikation
End Function

Public Function Transp() As CMatOp
    
End Function
    
Friend Function ToScalar() As Double
    ToScalar = vm.Scalar(0)
End Function

Friend Function ToVec2() As Vector2
    ToVec2 = vm.AVec2(0)
End Function
Friend Function ToVec3() As Vector3
    ToVec3 = vm.AVec3(0)
End Function
Friend Function ToVec4() As Vector4
    ToVec4 = vm.AVec4(0)
End Function
Friend Function ToVec5() As Vector5
    ToVec5 = vm.AVec5(0)
End Function
Friend Function ToVec6() As Vector6
    ToVec6 = vm.AVec6(0)
End Function
Friend Function ToVec7() As Vector7
    ToVec7 = vm.AVec7(0)
End Function
Friend Function ToVec8() As Vector8
    ToVec8 = vm.AVec8(0)
End Function
Friend Function ToVec9() As Vector9
    ToVec9 = vm.AVec9(0)
End Function
Friend Function ToVec10() As Vector10
    ToVec10 = vm.AVec10(0)
End Function

Friend Function ToMat2() As Matrix2
    ToMat2 = vm.AMat2(0)
End Function
Friend Function ToMat3() As Matrix3
    ToMat3 = vm.AMat3(0)
End Function
Friend Function ToMat4() As Matrix4
    ToMat4 = vm.AMat4(0)
End Function
Friend Function ToMat5() As Matrix5
    ToMat5 = vm.AMat5(0)
End Function
Friend Function ToMat6() As Matrix6
    ToMat6 = vm.AMat6(0)
End Function
Friend Function ToMat7() As Matrix7
    ToMat7 = vm.AMat7(0)
End Function
Friend Function ToMat8() As Matrix8
    ToMat8 = vm.AMat8(0)
End Function
Friend Function ToMat9() As Matrix9
    ToMat9 = vm.AMat9(0)
End Function
Friend Function ToMat10() As Matrix10
    ToMat10 = vm.AMat10(0)
End Function



