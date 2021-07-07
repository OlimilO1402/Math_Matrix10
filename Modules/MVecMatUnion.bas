Attribute VB_Name = "MVecMatUnion"
Option Explicit
Public Const FADF_AUTO       As Integer = &H1     'Das Array liegt im Stack
Public Const FADF_STATIC     As Integer = &H2     'Ein statisches Array
Public Const FADF_EMBEDDED   As Integer = &H4     'Das Feld ist in einer Struktur eingebettet
Public Const FADF_FIXEDSIZE  As Integer = &H10    'Die Grenzen des Arrays sind nicht änderbar
Public Const FADF_RECORD     As Integer = &H20    'Das Array enthält Records

'Public Type TUDTPtr
'    pSA        As Long
'    Reserved   As Long ' z.B. für vbVarType oder IRecordInfo
'    cDims      As Integer
'    fFeatures  As Integer
'    cbElements As Long
'    cLocks     As Long
'    pvData     As Long
'    cElements  As Long
'    lLBound    As Long
'End Type

Public Type TUdtPtr
    pSA        As Long    ' 4
    Reserved   As Long    ' 4 ' z.B. für vbVarType oder IRecordInfo
    cDims      As Integer ' 2
    fFeatures  As Integer ' 2
    cbElements As Long    ' 4
    cLocks     As Long    ' 4
    pvData     As Long    ' 4
    cElemts0   As Long    ' 4
    lLBound0   As Long    ' 4
    cElemts1   As Long    ' 4
    lLBound1   As Long    ' 4
End Type            ' Sum: 40

'OK wir brauchen hier auch einen pudt für ein 1D-Vektor-Array und für ein 2D-Matrix-Array
'soll man Skalar, Vektor und Matrix trennen? ok hier alles in einer großen Klasse
'nicht trennen sonst bräuchte man extra für den Skalar eine eigene Klasse
Public Type TVecMatUnion
    pudt0    As TUdtPtr  '40 'nur 1 Element (Array(0))
    pudt1    As TUdtPtr  '40 '1D-Vektor (Array(0 to u)
    pudt2    As TUdtPtr  '40 '2D-Matrix (Array(0 to u1, 0 to u2)
    Scalar() As Double   ' 4
    AVec2()  As Vector2  ' 4
    AVec3()  As Vector3  ' 4
    AVec4()  As Vector4  ' 4
    AVec5()  As Vector5  ' 4
    AVec6()  As Vector6  ' 4
    AVec7()  As Vector7  ' 4
    AVec8()  As Vector8  ' 4
    AVec9()  As Vector9  ' 4
    AVec10() As Vector10 ' 4
    AMat2()  As Matrix2  ' 4
    AMat3()  As Matrix3  ' 4
    AMat4()  As Matrix4  ' 4
    AMat5()  As Matrix5  ' 4
    AMat6()  As Matrix6  ' 4
    AMat7()  As Matrix7  ' 4
    AMat8()  As Matrix8  ' 4
    AMat9()  As Matrix9  ' 4
    AMat10() As Matrix10 ' 4
    AVec()   As Double   ' 4
    AMat()   As Double   ' 4
End Type          ' Sum: 204

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef Src As Any, ByVal BytLength As Long)

Public Declare Sub RtlZeroMemory Lib "kernel32" (ByRef Dst As Any, ByVal BytLength As Long)
    
'die Funktion ArrPtr geht bei allen Arrays außer bei String-Arrays
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef Arr() As Any) As Long

Public Sub New_TVecMatUnion(tmvu As TVecMatUnion, ByVal ptr As Long) ', ByVal mRows As Byte, ByVal nCols As Byte)
    'Dim dbl As Double
    'Debug.Print LenB(dbl)
    Dim pSA As Long
    With tmvu
        With .pudt0
            .pSA = VarPtr(.cDims)
            .cDims = 1
            .fFeatures = CInt(FADF_AUTO Or FADF_FIXEDSIZE) 'FADF_STATIC Or FADF_EMBEDDED Or FADF_FIXEDSIZE Or FADF_RECORD ' 54 ' = &H36
            .cbElements = 8 'sizeof Double
            '.cLocks = 1
            '.mRows = 10
            '.nCols = 10
            .cElemts0 = 1 '00
            .pvData = ptr
        End With
        .pudt1 = .pudt0
        With .pudt1
            .pSA = VarPtr(.cDims)
            .cElemts0 = 10
        End With
        .pudt2 = .pudt1
        With .pudt1
            .pSA = VarPtr(.cDims)
            .cElemts1 = 10
        End With
        
        pSA = VarPtr(.pudt0)
        
'Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
'    ByRef pDst As Any, ByRef pSrc As Any, ByVal bLength As Long)
'        Call RtlMoveMemory(ByVal ArrPtr(.Chars), ByVal VarPtr(.pudt), 4)

        'pSA = VarPtr(.pudt)
        SAPtr(ArrPtr(.Scalar)) = pSA
        SAPtr(ArrPtr(.AVec2)) = pSA
        SAPtr(ArrPtr(.AVec3)) = pSA
        SAPtr(ArrPtr(.AVec4)) = pSA
        SAPtr(ArrPtr(.AVec5)) = pSA
        SAPtr(ArrPtr(.AVec6)) = pSA
        SAPtr(ArrPtr(.AVec7)) = pSA
        SAPtr(ArrPtr(.AVec8)) = pSA
        SAPtr(ArrPtr(.AVec9)) = pSA
        SAPtr(ArrPtr(.AVec10)) = pSA
        
        SAPtr(ArrPtr(.AMat2)) = pSA
        SAPtr(ArrPtr(.AMat3)) = pSA
        SAPtr(ArrPtr(.AMat4)) = pSA
        SAPtr(ArrPtr(.AMat5)) = pSA
        SAPtr(ArrPtr(.AMat6)) = pSA
        SAPtr(ArrPtr(.AMat7)) = pSA
        SAPtr(ArrPtr(.AMat8)) = pSA
        SAPtr(ArrPtr(.AMat9)) = pSA
        SAPtr(ArrPtr(.AMat10)) = pSA
        
        pSA = VarPtr(.pudt1)
        SAPtr(ArrPtr(.AVec)) = pSA
        
        pSA = VarPtr(.pudt2)
        SAPtr(ArrPtr(.AMat)) = pSA
    End With
End Sub

Public Sub Del_TVecMatUnion(tmvu As TVecMatUnion)
    Dim l As Long: l = LenB(tmvu)
    If l <> 204 Then Debug.Assert True
    RtlZeroMemory ByVal VarPtr(tmvu), l
End Sub

'deswegen hier eine Hilfsfunktion für StringArrays
'Public Function StrArrPtr(ByRef strArr As Variant) As Long
'    Call RtlMoveMemory(StrArrPtr, ByVal VarPtr(strArr) + 8, 4)
'End Function

'jetzt kann das Property SAPtr für Alle Arrays verwendet werden,
'um den Zeiger auf den Safe-Array-Descriptor eines Arrays einem
'anderen Array zuzuweisen.
Public Property Get SAPtr(ByVal pArr As Long) As Long
    Call RtlMoveMemory(SAPtr, ByVal pArr, 4)
End Property

Public Property Let SAPtr(ByVal pArr As Long, ByVal RHS As Long)
    Call RtlMoveMemory(ByVal pArr, ByVal RHS, 4)
End Property

Public Sub ZeroSAPtr(ByVal pArr As Long)
    Call RtlZeroMemory(ByVal pArr, 4)
End Sub

