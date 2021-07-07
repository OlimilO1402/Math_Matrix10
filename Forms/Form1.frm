VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17055
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   17055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnFont 
      Caption         =   "Font"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton BtnCreate 
      Caption         =   "Create Vector / Matrix"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton BtnCalcOld 
      Caption         =   " = (old)"
      Height          =   375
      Left            =   13440
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Panel 
      BorderStyle     =   0  'Kein
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   16935
      TabIndex        =   4
      Top             =   600
      Width           =   16935
      Begin VB.PictureBox PanelInner 
         BorderStyle     =   0  'Kein
         Height          =   4215
         Left            =   0
         ScaleHeight     =   4215
         ScaleWidth      =   11175
         TabIndex        =   6
         Top             =   0
         Width           =   11175
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   5640
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Beides
            TabIndex        =   8
            Top             =   0
            Width           =   5535
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3975
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Beides
            TabIndex        =   7
            Top             =   0
            Width           =   5535
         End
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   11280
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   5
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.ComboBox CmbRC 
      Height          =   315
      Left            =   12600
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton BtnRndMat 
      Caption         =   "Rnd Mat (old)"
      Height          =   375
      Left            =   14640
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton BtnCalc 
      Caption         =   "="
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox CmbOps 
      Height          =   315
      Left            =   5160
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ok; Create
'ok; Column
'ok; Column-Swap
'ok; Unitymatrix (creational)
'ok; IsSymmetric
'ok; All Ones(creational)
'ok; Row
'ok; Row -Swap
'ok; Trace
'ok; Add matrix
'ok; Adjunkte
'ok; Determinante
'ok; Inverse
'ok; IsEqual
'ok; Minore
'ok; Multiply matrix
'ok; Multiply scalar
'ok; Parse (from a string)
'ok; Random matrix(creational)
'ok; Solve Linear Equation A*x=b(with Determinante)
'ok; Untermatrix
'ok; To Array
'ok; Multiply vector

'Matrix-Tipp für kleine Matrizen bis 10x10
'* einfache Matrizenoperationen
'  - Erzeugen, (matx)
'  - Einheitsmatrix (emat)
'  - Addieren (add),
'  - Subtrahieren (sub),
'  - Multiplizieren Matrix mit Skalar (smul),
'  - Multiplikation Matrix mit Vektor (vmul),
'  - Multiplizieren Matrix mit Matrix (mul),
'  - nach Array konvertieren
'  - Einlesen von  String (Parse)
'  - Ausgeben nach String (ToStr)
'
'* fortgeschrittene Matrizenoperationen
'  - Zeile/Spalte lesen und schreiben PropGet/Let (Row, Col)
'  - Transponierte (tra)
'  - Determinante (det)
'  - Adjunkte (adj)
'  - Inverse (inv)
'  - Lösen von LGS (solve)
'  - Untermatrix (umat)
'  - Minoren-Determinante
'
'zum Überprüfen der Berechnung siehe z.B.:
'http://matrizen-rechner.de/
'https://matrixcalc.org/de/
'https://rechneronline.de/lineare-algebra/matrizen.php
'https://rechneronline.de/lineare-algebra/gleichungssysteme.php

'Matrizen-Rechner online:
'https://matrixcalc.org/de/
'oder auch
'https://matrizen-rechner.de/
'(hat sehr viele Funktionen)
'Generelle Operationen:
'----------------------
'ok; Adjunkte
'==-->> Äquilibrierung
'ok; ax=b nach x lösen 'LGS
'==-->> Cholesky -Zerlegung
'ok; Determinante nach Laplace
'==-->> Determinante mittels LR-Zerlegung
'==-->> Dimension Kern
'==-->> Dimensionsregel
'==-->> Gauss Elimination
'==-->> Householder Matrix
'ok; Inverse
'ok; Kreuzprodukt
'==-->> LR -Zerlegung
'ok; Matrixaddition
'ok; Matrix -Vektor - Multiplikation
'ok; Matrixmultiplikation
'==-->> QR-Zerlegung mittels Householdermatrix
'==-->> QR-Zerlegung nach Givens
'==-->> Rang der Matrix
'ok; Skalar * Matrix
'==-->> Skalarprodukt
'==-->> Spatprodukt
'==-->> Spur der Matrix
'ok; Matrix transponieren
'ok;Vektor1_transponiert * Vektor2
'
'Matrixnormen:
'-------------
'||A|| 1
'||A|| oo
'||A||
'kappa -rel
'Vektornormen:
'-------------
'||A||
'||A|| 2
'||A|| oo
'Unfertig:
'---------
'Eigenwerte berechnen (nur reelle)
'Pivotisierung
'Pivotisierung und Äuquilibrierung

'bsp für eine 5x5 Matrix:
'3.0 5.3 5.6 3.5 6.8
'0.4 8.2 6.7 1.9 2.2
'7.8 8.3 7.7 3.3 1.9
'5.5 8.8 3.0 1.0 5.1
'5.1 5.1 3.6 5.8 5.7
'Determinante: 5902.33429
'stimmt juhu : 5902.33429

'bsp: für eine 6x6 Matrix:
'3.0 5.3 5.6 3.5 6.8 5.7
'0.4 8.2 6.7 1.9 2.2 5.3
'7.8 8.3 7.7 3.3 1.9 4.8
'5.5 8.8 3.0 1.0 5.1 6.4
'5.1 5.1 3.6 5.8 5.7 4.9
'3.5 2.7 5.7 8.2 9.6 2.9
'Determinante: -16178.400057
'stimmt juhu : -16178.400057

'bsp für eine 7x7 Matrix
'3.0 5.3 5.6 3.5 6.8 5.7 1.0
'0.4 8.2 6.7 1.9 2.2 5.3 2.0
'7.8 8.3 7.7 3.3 1.9 4.8 3.0
'5.5 8.8 3.0 1.0 5.1 6.4 4.0
'5.1 5.1 3.6 5.8 5.7 4.9 5.0
'3.5 2.7 5.7 8.2 9.6 2.9 6.0
'1.0 2.0 3.0 4.0 5.0 6.0 7.0
'Determinante: -144607.476869
'stimmt juhu : -144607.476869

'bsp für eine 8x8 Matrix
'3.0 5.3 5.6 3.5 6.8 5.7 1.0 7.0
'0.4 8.2 6.7 1.9 2.2 5.3 2.0 6.0
'7.8 8.3 7.7 3.3 1.9 4.8 3.0 5.0
'5.5 8.8 3.0 1.0 5.1 6.4 4.0 4.0
'5.1 5.1 3.6 5.8 5.7 4.9 5.0 3.0
'3.5 2.7 5.7 8.2 9.6 2.9 6.0 2.0
'1.0 2.0 3.0 4.0 5.0 6.0 7.0 1.0
'7.0 6.0 5.0 4.0 3.0 2.0 1.0 1.5
'Determinante: 357238.1026035
'stimmt juhu : 357238.1026035

'bsp für eine 9x9 Matrix
' 3.0 5.3 5.6 3.5 6.8 5.7 1.0 7.0 10.0
' 0.4 8.2 6.7 1.9 2.2 5.3 2.0 6.0  8.0
' 7.8 8.3 7.7 3.3 1.9 4.8 3.0 5.0  4.0
' 5.5 8.8 3.0 1.0 5.1 6.4 4.0 4.0  2.0
' 5.1 5.1 3.6 5.8 5.7 4.9 5.0 3.0  1.0
' 3.5 2.7 5.7 8.2 9.6 2.9 6.0 2.0  3.0
' 1.0 2.0 3.0 4.0 5.0 6.0 7.0 1.0  5.0
' 7.0 6.0 5.0 4.0 3.0 2.0 1.0 1.5  7.0
'10.0 8.0 4.0 2.0 1.0 3.0 5.0 7.0  9.0
'Determinante: 11490774.085104993
'stimmt juhu : 11490774.085104993

'bsp für eine 10x10 Matrix
' 3.0 5.3 5.6 3.5 6.8 5.7 1.0 7.0 10.0 1.5
' 0.4 8.2 6.7 1.9 2.2 5.3 2.0 6.0  8.0 2.4
' 7.8 8.3 7.7 3.3 1.9 4.8 3.0 5.0  4.0 3.3
' 5.5 8.8 3.0 1.0 5.1 6.4 4.0 4.0  2.0 4.2
' 5.1 5.1 3.6 5.8 5.7 4.9 5.0 3.0  1.0 5.1
' 3.5 2.7 5.7 8.2 9.6 2.9 6.0 2.0  3.0 4.2
' 1.0 2.0 3.0 4.0 5.0 6.0 7.0 1.0  5.0 3.3
' 7.0 6.0 5.0 4.0 3.0 2.0 1.0 1.5  7.0 2.4
'10.0 8.0 4.0 2.0 1.0 3.0 5.0 7.0  9.0 1.5
'0.1  0.2 0.3 0.4 0.5 0.6 0.7 0.8  0.9 0.123
'Determinante: -300781.274894236
'stimmt juhu : -300781.274894235
'Dim sc  As Double
'Dim A2  As Matrix2, b2   As Matrix2, c2   As Matrix2, V2   As Vector2
'Dim A3  As Matrix3, b3   As Matrix3, C3   As Matrix3, V3   As Vector3
'Dim A4  As Matrix4, b4   As Matrix4, C4   As Matrix4, V4   As Vector4
'Dim A5  As Matrix5, b5   As Matrix5, C5   As Matrix5, V5   As Vector5
'Dim A6  As Matrix6, b6   As Matrix6, C6   As Matrix6, V6   As Vector6
'Dim A7  As Matrix7, b7   As Matrix7, C7   As Matrix7, V7   As Vector7
'Dim A8  As Matrix8, b8   As Matrix8, C8   As Matrix8, V8   As Vector8
'Dim A9  As Matrix9, b9   As Matrix9, C9   As Matrix9, V9   As Vector9
'Dim A10 As Matrix10, b10 As Matrix10, C10 As Matrix10, V10 As Vector10
'Dim vb2 As Vector2, vb3 As Vector3, vb4 As Vector4, vb5 As Vector5, vb6 As Vector6, vb7 As Vector7, vb8 As Vector8, vb9 As Vector9, vb10 As Vector10
'Dim vx2 As Vector2, vx3 As Vector3, vx4 As Vector4, vx5 As Vector5, vx6 As Vector6, vx7 As Vector7, vx8 As Vector8, vx9 As Vector9, vx10 As Vector10
'
'Dim MyMat10_1 As Matrix10
'Dim MyMat10_2 As Matrix10

Dim WithEvents Split1 As Splitter
Attribute Split1.VB_VarHelpID = -1
Dim WithEvents Split2 As Splitter
Attribute Split2.VB_VarHelpID = -1
Dim ActTB As TextBox

Dim FCSetMat1 As FrmCreateSetting
Dim FCSetMat2 As FrmCreateSetting
Dim FCSetMat3 As FrmCreateSetting

Private Sub Form_Load()
        
    Set Split1 = Splitter(False, Me, PanelInner, "Splitter1", Text1, Text2)
    Split1.LeftTopPos = Text1.Width
    Split1.BorderStyle = bsXPStyl
    
    Set Split2 = Splitter(False, Me, Panel, "Splitter2", PanelInner, Text3)
    Split2.LeftTopPos = PanelInner.Width
    Split2.BorderStyle = bsXPStyl
    
    Set SynchronizedFont = MMain.MyFont
    
    'MMatrices.InitMatXOnes
    'Debug.Print MMatrices.Mat10_ToStr(MMatrices.Mat10_Ones)
    
    'ok; Create 'ok; Column 'ok; Column-Swap 'ok; Unitymatrix (creational) 'ok; IsSymmetric 'ok; All Ones(creational) 'ok; Row 'ok; Row -Swap 'ok; Trace
    'ok; Add matrix 'ok; Adjunkte 'ok; Determinante 'ok; Inverse 'ok; IsEqual 'ok; Minore 'ok; Multiply matrix 'ok; Multiply scalar 'ok; Parse (from a string)
    'ok; Random matrix(creational) 'ok; Solve Linear Equation A*x=b(with Determinante) 'ok; Untermatrix 'ok; To Array 'ok; Multiply vector
    
    Me.Caption = "Matrizen - bis 10x10"
    With CmbOps
        '.Font = "Courier New": .FontSize = 10
        '.Font = "Consolas": .FontSize = 10
        Set .Font = MMain.MyFont
        .AddItem "  +  ": .AddItem "  -  ": .AddItem "  *  ": .AddItem "  x  "
        .AddItem " tra ": .AddItem " det ": .AddItem " adj ": .AddItem " inv "
        .AddItem " uni ": .AddItem " one ": .AddItem " col ": .AddItem " row "
        .AddItem "  =  ": .AddItem "isSym": .AddItem "eigval": .AddItem "z+=z*f"
        .AddItem "A*x=b": .AddItem "Len-V": .AddItem "NormV"
        .ListIndex = 0
    End With
    'New_TVecMatUnion VMat, 2, 1
    
    Randomize
    Split1_OnMove Split1
    Split2_OnMove Split2
    
End Sub


Private Sub BtnCreate_Click()
    If ActTB Is Nothing Then Set ActTB = Text1
    Dim s As String: s = ActTB.Text
    Dim mA As Matrix10
    Dim VMatA As CMatOp
    If Len(s) Then
        Set VMatA = CMatOp(VarPtr(mA), s)
    End If
    If FrmCreateMatVec.ShowDialog(Me, FCSetMat1, VMatA) = vbCancel Then Exit Sub
    If Not VMatA Is Nothing Then
        ActTB.Text = VMatA.ToStr
    End If
    
    'ohje, ja halt nicht einfach nur den string übergeben, und den rest löschen
    
    
    
'    Dim mRows As Byte: mRows = Cmb_GetValue(CmbmRows)
'    Select Case LCase(Trim(CmbmRows.Text))
'    Case "rnd"
'    Case Else
'    End Select
'
'    Dim nCols As Byte
End Sub


Private Sub BtnCalc_Click()
    'Dim A_mRows As Long, A_nCols As Long
    'Dim B_mRows As Long, B_nCols As Long
    'GetRowsCols Text1.Text, A_mRows, A_nCols
    'GetRowsCols Text2.Text, B_mRows, B_nCols
    'Dim maxrc As Long: maxrc = Max(A_mRows, Max(B_mRows, Max(A_nCols, B_nCols)))
    
'yeah the following ideas:
'* on the one hand we have textbox and matrix will be read newly everytime the user presses the calc-button
'  so the user can edit the matrix at any time in the textbox
'* on the other hand we could store the matrix in the form and store properties like number of digits and all
'  the other settings one can do in the create-matrix-dialog
'
'OK we have to make a decision: we do the following,
'we do not store the matrices in the form, the matrices live solely in the textboxes,
'but we store the last settings of the dialog, OK we could do the following
'maybe we have 3 different instances of the FrmCreateMatVec-Dialog for every textbox?
'OK wie soll man überhaupt Rnd raus und rein bringen?
'
    Dim s As String
    
    'WORK IN PROGRESS!!!
    
    s = Text1.Text
    If Len(s) Then
        Dim mA As Matrix10
        Dim VMatA As CMatOp: Set VMatA = CMatOp(VarPtr(mA), s)
        If Not VMatA Is Nothing Then Text1.Text = VMatA.ToStr
    End If
    
    s = Text2.Text
    If Len(s) Then
        Dim mB As Matrix10
        Dim VMatB As CMatOp: Set VMatB = CMatOp(VarPtr(mB), s)
        If Not VMatB Is Nothing Then Text2.Text = VMatB.ToStr
    End If
    Dim op As String: op = CmbOps.List(CmbOps.ListIndex)
    
    Text3.Text = ""
    If Not VMatA Is Nothing Then
        Dim VMatC As CMatOp: Set VMatC = VMatA.op(op, VMatB)
        If Not VMatC Is Nothing Then
            'Dim nk As Byte: nk = CByte(CmbNk.Text)
            'Dim nk As Integer: nk = CInt(CmbRC.Text)
            Text3.Text = VMatC.ToStr
        End If
    End If
    'Dim mB As Matrix10
    'Dim VMatB As CMatOp: Set VMatB = CMatOp(VarPtr(mB), Text2.Text) 'B_mRows, B_nCols)
    
    'Debug.Print VMatA.ToStr
    
End Sub

Private Sub BtnFont_Click()
    Dim fnt As StdFont: Set fnt = MNew.StdFontC(MMain.MyFont)
    If FrmFont.ShowDialog(Me, fnt) = vbCancel Then Exit Sub
    StdFont_Assign MMain.MyFont, fnt
End Sub

Private Sub Text1_Click()
    Set ActTB = Text1
End Sub
Private Sub Text2_Click()
    Set ActTB = Text2
End Sub
Private Sub Text3_Click()
    Set ActTB = Text3
End Sub

'Private Sub Command1_Click()
''    Dim m1 As Matrix4: m1 = Mat4_Rnd
''    Text1.Text = Mat4_ToStr(m1, 3)
''    Dim m2 As Matrix4
''
''    m2 = Mat4_ColSwap(m1, 1, 2)
''    Text2.Text = Mat4_ToStr(m2, 3)
''    m2 = Mat4_ColSwap(m1, 2, 1)
''    Text3.Text = Mat4_ToStr(m2, 3)
'
'    Dim V1 As Vector3: V1 = Vec3(1, 2, 3)
'    Dim V2 As Vector3: V2 = Vec3(-2, 1, -1)
'    'Dim v3 As Vector3: v3 = Vec3(15, 12, 11)
'
'    Dim vErg As Vector3: vErg = Vec3_cross(V1, V2)
'
'    Text3.Text = Vec3_ToStr(vErg) '  5, -10, 5
'
'    Dim m As Matrix3
'    Mat3_Col(m, 0) = Vec3(1, 1, 1)
'    Mat3_Col(m, 1) = V1
'    Mat3_Col(m, 2) = V2
'
'    Text2.Text = Mat3_det(m)
'
'End Sub


Private Property Set SynchronizedFont(Value As StdFont)
    Set Text1.Font = Value
    Static IsFontSynchronized As Boolean
    If IsFontSynchronized Then Exit Property
    Set Text2.Font = Text1.Font
    Set Text3.Font = Text2.Font
    IsFontSynchronized = True
End Property

Sub GetRowsCols(T As String, ByRef rows_out As Long, ByRef cols_out As Long)
    rows_out = GetRows(T):    cols_out = GetCols(T)
End Sub
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

'Private Sub BtnRndMat_Click()
'    Text1.Text = GetRandomMat
'    Text2.Text = GetRandomMat 'Text1.Text
'    Text3.Text = ""
'    BtnCalc_Click
'End Sub

Function Cmb_GetValue(aCmb As ComboBox)
    'if acmb.Text = "rnd"
End Function
Function GetRndMat() As String
    Dim Size As Long: Size = Rnd * 8 + 2
    Dim s As String
    'ReDim a(0 To size - 1, 0 To size - 1)
    Dim i As Long, j As Long
    For i = 1 To Size 'UBound(a, 1)
        For j = 1 To Size 'UBound(a, 2)
            'a(i, j) =  Rnd * 200 * Sgn(Rnd - 0.5)
            s = s & Replace(Format(Rnd * 200 * Sgn(Rnd - 0.5), "0.00"), ",", ".") & " "
        Next
        s = s & vbCrLf
    Next
    GetRndMat = s
End Function

Private Sub Form_Resize()
    Dim brdr As Single ': brdr = 8 * Screen.TwipsPerPixelX
    Dim l As Single ': l = brdr
    Dim T As Single: T = Panel.Top
    
    'W = (Me.ScaleWidth - 4 * brdr) / 3
    Dim W As Single: W = Me.ScaleWidth - l - brdr / 4
    Dim H As Single: H = Me.ScaleHeight - T - brdr / 4
    If W > 0 And H > 0 Then
        Panel.Move l, T, W, H
        l = l + Text1.Width + brdr
        'Dim L2 As Single: L2 = l - (brdr + CmbOps.Width) / 2 - Split1.Width / 2
        'CmbOps.Move L2, T - CmbOps.Height
        'Text2.Move l, T, W, H
        'l = l + Text2.Width + brdr
        'L2 = l - (brdr + BtnCalc.Width) / 2
        'BtnCalc.Move L2, T - BtnCalc.Height
        'Text3.Move l, T, W, H
    End If

'    Dim brdr: brdr = 8 * Screen.TwipsPerPixelX
'    Dim l, T, W, H
'    Dim L2
'    l = brdr: T = Text1.Top
'    W = (Me.ScaleWidth - 4 * brdr) / 3
'    H = Me.ScaleHeight - T - brdr
'    If W > 0 And H > 0 Then
'        Text1.Move l, T, W, H
'        l = l + Text1.Width + brdr
'        L2 = l - (brdr + CmbOps.Width) / 2
'        CmbOps.Move L2, T - CmbOps.Height
'        Text2.Move l, T, W, H
'        l = l + Text2.Width + brdr
'        L2 = l - (brdr + BtnCalc.Width) / 2
'        BtnCalc.Move L2, T - BtnCalc.Height
'        Text3.Move l, T, W, H
'    End If
End Sub

Private Sub Split1_OnMove(Sender As Splitter)
    'Dim brdr As Single ': brdr = 8 * Screen.TwipsPerPixelX
    Dim l As Single ': l = brdr
    'Dim T As Single: T = Panel.Top
    'l = l + Text1.Width + brdr
    l = Text1.Width - CmbOps.Width / 2 + Split1.Width / 2
    'Dim L2 As Single: L2 = l - (brdr + CmbOps.Width) / 2 '- Split1.Width
    'CmbOps.Move L2 ', T - CmbOps.Height
    CmbOps.Move l
    'W = (Me.ScaleWidth - 4 * brdr) / 3
    'Dim W As Single: W = (Me.ScaleWidth - 2 * brdr)
    'Dim H As Single: H = Me.ScaleHeight - T - brdr
    

End Sub

Private Sub Split2_OnMove(Sender As Splitter)
    'Dim brdr As Single ': brdr = 8 * Screen.TwipsPerPixelX
    Dim l As Single ': l = brdr
    'Dim T As Single: T = Panel.Top - BtnCalc.Height
    'l = Text1.Width + split1.Width + text2.Width
    'l = l + Text2.Width + brdr
    'Dim L2 As Single: L2 = l - (brdr + CmbOps.Width) / 2 '- Split2.Width
    'L2 = l - (brdr + BtnCalc.Width) / 2
    l = PanelInner.Width - BtnCalc.Width / 2 + Split2.Width / 2
    BtnCalc.Move l '2 ', T
End Sub

'Function GetRandomMat() As String
''    Dim r As Long: r = Rnd * 10 + 2
''    Select Case r
''    Case 2: GetRandomMat = GetMat2
''    Case 3: GetRandomMat = GetMat3
''    Case 4: GetRandomMat = GetMat4
''    Case 5: GetRandomMat = GetMat5
''    Case 6: GetRandomMat = GetMat6
''    Case 7: GetRandomMat = GetMat7
''    Case 8: GetRandomMat = GetMat8
''    Case 9: GetRandomMat = GetMat9
''    Case 10: GetRandomMat = GetMat10
''    Case Else
''            GetRandomMat = GetRndMat
''    End Select
'End Function

'Function GetMat2() As String
'    GetMat2 = "11 12" & vbCrLf & _
'              "21 22"
'End Function
'Function GetMat3() As String
'    GetMat3 = "11 12 13" & vbCrLf & _
'              "21 22 23" & vbCrLf & _
'              "31 32 33"
'End Function
'Function GetMat4() As String
'    GetMat4 = "11 12 13 14" & vbCrLf & _
'              "21 22 23 24" & vbCrLf & _
'              "31 32 33 34" & vbCrLf & _
'              "41 42 43 44"
'End Function
'Function GetMat5() As String
'    GetMat5 = "11 12 13 14 15" & vbCrLf & _
'              "21 22 23 24 25" & vbCrLf & _
'              "31 32 33 34 35" & vbCrLf & _
'              "41 42 43 44 45" & vbCrLf & _
'              "51 52 53 54 55"
'End Function
'Function GetMat6() As String
'    GetMat6 = "11 12 13 14 15 16" & vbCrLf & _
'              "21 22 23 24 25 26" & vbCrLf & _
'              "31 32 33 34 35 36" & vbCrLf & _
'              "41 42 43 44 45 46" & vbCrLf & _
'              "51 52 53 54 55 56" & vbCrLf & _
'              "61 62 63 64 65 66"
'End Function
'Function GetMat7() As String
'    GetMat7 = "11 12 13 14 15 16 17" & vbCrLf & _
'              "21 22 23 24 25 26 27" & vbCrLf & _
'              "31 32 33 34 35 36 37" & vbCrLf & _
'              "41 42 43 44 45 46 47" & vbCrLf & _
'              "51 52 53 54 55 56 57" & vbCrLf & _
'              "61 62 63 64 65 66 67" & vbCrLf & _
'              "71 72 73 74 75 76 77"
'End Function
'Function GetMat8() As String
'    GetMat8 = "11 12 13 14 15 16 17 18" & vbCrLf & _
'              "21 22 23 24 25 26 27 28" & vbCrLf & _
'              "31 32 33 34 35 36 37 38" & vbCrLf & _
'              "41 42 43 44 45 46 47 48" & vbCrLf & _
'              "51 52 53 54 55 56 57 58" & vbCrLf & _
'              "61 62 63 64 65 66 67 68" & vbCrLf & _
'              "71 72 73 74 75 76 77 78" & vbCrLf & _
'              "81 82 83 84 85 86 87 88"
'End Function
'Function GetMat9() As String
'    GetMat9 = "11 12 13 14 15 16 17 18 19" & vbCrLf & _
'              "21 22 23 24 25 26 27 28 29" & vbCrLf & _
'              "31 32 33 34 35 36 37 38 39" & vbCrLf & _
'              "41 42 43 44 45 46 47 48 49" & vbCrLf & _
'              "51 52 53 54 55 56 57 58 59" & vbCrLf & _
'              "61 62 63 64 65 66 67 68 69" & vbCrLf & _
'              "71 72 73 74 75 76 77 78 79" & vbCrLf & _
'              "81 82 83 84 85 86 87 88 89" & vbCrLf & _
'              "91 92 93 94 95 96 97 98 99"
'End Function
'Function GetMat10() As String
'    GetMat10 = "01 02 03 04 05 06 07 08 09 10" & vbCrLf & _
'               "11 12 13 14 15 16 17 18 19 20" & vbCrLf & _
'               "21 22 23 24 25 26 27 28 29 30" & vbCrLf & _
'               "31 32 33 34 35 36 37 38 39 40" & vbCrLf & _
'               "41 42 43 44 45 46 47 48 49 50" & vbCrLf & _
'               "51 52 53 54 55 56 57 58 59 60" & vbCrLf & _
'               "61 62 63 64 65 66 67 68 69 70" & vbCrLf & _
'               "71 72 73 74 75 76 77 78 79 80" & vbCrLf & _
'               "81 82 83 84 85 86 87 88 89 90" & vbCrLf & _
'               "91 92 93 94 95 96 97 98 99 100"
'End Function



'Private Sub BtnCalcOld_Click()
'    '
'    'So ist das Mist, besser für jede Operation eine eigene Function die dann Strings t1, t2, t3 zurückgibt!!!
'    '
'    'zuviele lokale nichtstatische Variablen
'
'    'die maximalen Größen der Matrizen herausfinden
'    Dim A_mRows As Long, A_nCols As Long
'    Dim B_mRows As Long, B_nCols As Long
'    GetRowsCols Text1.Text, A_mRows, A_nCols
'    GetRowsCols Text2.Text, B_mRows, B_nCols
'    Dim maxrc As Long: maxrc = Max(A_mRows, Max(B_mRows, Max(A_nCols, B_nCols)))
'    'hja Blöd Vektor kommt nicht vor, so ist das Mist
'    Select Case maxrc
'    Case 2:   A2 = Mat2_Parse(Text1.Text):  b2 = Mat2_Parse(Text2.Text)
'    Case 3:   A3 = Mat3_Parse(Text1.Text):  b3 = Mat3_Parse(Text2.Text)
'    Case 4:   A4 = Mat4_Parse(Text1.Text):  b4 = Mat4_Parse(Text2.Text)
'    Case 5:   A5 = Mat5_Parse(Text1.Text):  b5 = Mat5_Parse(Text2.Text)
'    Case 6:   A6 = Mat6_Parse(Text1.Text):  b6 = Mat6_Parse(Text2.Text)
'    Case 7:   A7 = Mat7_Parse(Text1.Text):  b7 = Mat7_Parse(Text2.Text)
'    Case 8:   A8 = Mat8_Parse(Text1.Text):  b8 = Mat8_Parse(Text2.Text)
'    Case 9:   A9 = Mat9_Parse(Text1.Text):  b9 = Mat9_Parse(Text2.Text)
'    Case 10: A10 = Mat10_Parse(Text1.Text): b10 = Mat10_Parse(Text2.Text)
'    End Select
'    'Dim C6 As Matrix6
'    Dim Scalar As Double
'    Dim op As String: op = CmbOps.List(CmbOps.ListIndex)
'    If op = "  *  " And (A_mRows = 1 And A_nCols = 1) Or (B_mRows = 1 And B_nCols = 1) Then
'        op = "smul"
'        Scalar = Val(Text2.Text)
'    End If
'    If op = "  *  " And (B_nCols = 1) Then
'       'op = " vmul"
'    End If
'    Select Case op
'    Case "  +  "
'        Select Case maxrc
'        Case 2:  c2 = Mat2_add(A2, b2):  LSet C10 = c2
'        Case 3:  C3 = Mat3_add(A3, b3):  LSet C10 = C3
'        Case 4:  C4 = Mat4_add(A4, b4):  LSet C10 = C4
'        Case 5:  C5 = Mat5_add(A5, b5):  LSet C10 = C5
'        Case 6:  C6 = Mat6_add(A6, b6):  LSet C10 = C6
'        Case 7:  C7 = Mat7_add(A7, b7):  LSet C10 = C7
'        Case 8:  C8 = Mat8_add(A8, b8):  LSet C10 = C8
'        Case 9:  C9 = Mat9_add(A9, b9):  LSet C10 = C9
'        Case 10: C10 = Mat10_add(A10, b10)
'        End Select
'    Case "  -  "
'        Select Case maxrc
'        Case 2: c2 = Mat2_sub(A2, b2): LSet C10 = c2
'        Case 3: C3 = Mat3_sub(A3, b3): LSet C10 = C3
'        Case 4: C4 = Mat4_sub(A4, b4): LSet C10 = C4
'        Case 5: C5 = Mat5_sub(A5, b5): LSet C10 = C5
'        Case 6: C6 = Mat6_sub(A6, b6): LSet C10 = C6
'        Case 7: C7 = Mat7_sub(A7, b7): LSet C10 = C7
'        Case 8: C8 = Mat8_sub(A8, b8): LSet C10 = C8
'        Case 9: C9 = Mat9_sub(A9, b9): LSet C10 = C9
'        Case 10: C10 = Mat10_sub(A10, b10)
'        End Select
'    Case "vmul"
'        Select Case maxrc 'wieso hier LSet = ??
'        Case 2: LSet c2 = Mat2_vmul(A2, V2): LSet C10 = c2
'        Case 3: LSet C3 = Mat3_vmul(A3, V3): LSet C10 = C3
'        Case 4: LSet C4 = Mat4_vmul(A4, V4): LSet C10 = C4
'        Case 5: LSet C5 = Mat5_vmul(A5, V5): LSet C10 = C5
'        Case 6: LSet C6 = Mat6_vmul(A6, V6): LSet C10 = C6
'        Case 7: LSet C7 = Mat7_vmul(A7, V7): LSet C10 = C7
'        Case 8: LSet C8 = Mat8_vmul(A8, V8): LSet C10 = C8
'        Case 9: LSet C9 = Mat9_vmul(A9, V9): LSet C10 = C9
'        Case 10: LSet C10 = Mat10_vmul(A10, V10)
'        End Select
'    Case "smul"
'        Select Case maxrc
'        Case 2: c2 = Mat2_smul(A2, Scalar): LSet C10 = c2
'        Case 3: C3 = Mat3_smul(A3, Scalar): LSet C10 = C3
'        Case 4: C4 = Mat4_smul(A4, Scalar): LSet C10 = C4
'        Case 5: C5 = Mat5_smul(A5, Scalar): LSet C10 = C5
'        Case 6: C6 = Mat6_smul(A6, Scalar): LSet C10 = C6
'        Case 7: C7 = Mat7_smul(A7, Scalar): LSet C10 = C7
'        Case 8: C8 = Mat8_smul(A8, Scalar): LSet C10 = C8
'        Case 9: C9 = Mat9_smul(A9, Scalar): LSet C10 = C9
'        Case 10: C10 = Mat10_smul(A10, Scalar)
'        End Select
'    Case "  *  "
'        Select Case maxrc
'        Case 2: c2 = Mat2_mul(A2, b2): LSet C10 = c2
'        Case 3: C3 = Mat3_mul(A3, b3): LSet C10 = C3
'        Case 4: C4 = Mat4_mul(A4, b4): LSet C10 = C4
'        Case 5: C5 = Mat5_mul(A5, b5): LSet C10 = C5
'        Case 6: C6 = Mat6_mul(A6, b6): LSet C10 = C6
'        Case 7: C7 = Mat7_mul(A7, b7): LSet C10 = C7
'        Case 8: C8 = Mat8_mul(A8, b8): LSet C10 = C8
'        Case 9: C9 = Mat9_mul(A9, b9): LSet C10 = C9
'        Case 10: C10 = Mat10_mul(A10, b10)
'        End Select
'    Case "  x  "
''        'Dim sc As Double
'        Select Case maxrc
''        'Case 2: sc = Vec2_cross(Mat2_Col(A2, 0), Mat2_Col(b2, 0)): C10.aa = sc
'        Case 3: vb3 = Vec3_cross(Mat3_Col(A3, 0), Mat3_Col(A3, 1)): LSet C10 = vb3
'        'Case 4: vb4 = Vec4_cross(Mat4_Col(A4, 0), Mat4_Col(A4, 1), Mat4_Col(A4, 2)): LSet C10 = vb4
'        'Case 5: vb5 = Vec5_cross(Mat5_Col(A5, 0), Mat5_Col(A5, 1), Mat5_Col(A5, 2), Mat5_Col(A5, 3)): LSet C10 = vb5
'        'Case 6: vb6 = Vec6_cross(Mat6_Col(A6, 0), Mat6_Col(A6, 1), Mat6_Col(A6, 2), Mat6_Col(A6, 3), Mat6_Col(A6, 4)): LSet C10 = vb6
'        'Case 7: vb7 = Vec7_cross(Mat7_Col(A7, 0), Mat7_Col(A7, 1), Mat7_Col(A7, 2), Mat7_Col(A7, 3), Mat7_Col(A7, 4), Mat7_Col(A7, 5)): LSet C10 = vb7
'        'Case 8: vb8 = Vec8_cross(Mat8_Col(A8, 0), Mat8_Col(A8, 1), Mat8_Col(A8, 2), Mat8_Col(A8, 3), Mat8_Col(A8, 4), Mat8_Col(A8, 5), Mat8_Col(A8, 6)): LSet C10 = vb8
'        'Case 9: vb9 = Vec9_cross(Mat9_Col(A9, 0), Mat9_Col(A9, 1), Mat9_Col(A9, 2), Mat9_Col(A9, 3), Mat9_Col(A9, 4), Mat9_Col(A9, 5), Mat9_Col(A9, 6), Mat9_Col(A9, 7)): LSet C10 = vb9
'        'Case 10: vb10 = Vec10_cross(Mat10_Col(A10, 0), Mat10_Col(A10, 1), Mat10_Col(A10, 2), Mat10_Col(A10, 3), Mat10_Col(A10, 4), Mat10_Col(A10, 5), Mat10_Col(A10, 6), Mat10_Col(A10, 7), Mat10_Col(A10, 8)): LSet C10 = vb10
'        End Select
'    Case " tra "
'        Select Case maxrc
'        Case 2: c2 = Mat2_tra(A2): LSet C10 = c2
'        Case 3: C3 = Mat3_tra(A3): LSet C10 = C3
'        Case 4: C4 = Mat4_tra(A4): LSet C10 = C4
'        Case 5: C5 = Mat5_tra(A5): LSet C10 = C5
'        Case 6: C6 = Mat6_tra(A6): LSet C10 = C6
'        Case 7: C7 = Mat7_tra(A7): LSet C10 = C7
'        Case 8: C8 = Mat8_tra(A8): LSet C10 = C8
'        Case 9: C9 = Mat9_tra(A9): LSet C10 = C9
'        Case 10: C10 = Mat10_tra(A10)
'        End Select
'    Case "  =  "
'        Dim b As Boolean
'        Select Case maxrc
'        Case 2: b = Mat2_IsEqual(A2, b2)
'        Case 3: b = Mat3_IsEqual(A3, b3)
'        Case 4: b = Mat4_IsEqual(A4, b4)
'        Case 5: b = Mat5_IsEqual(A5, b5)
'        Case 6: b = Mat6_IsEqual(A6, b6)
'        Case 7: b = Mat7_IsEqual(A7, b7)
'        Case 8: b = Mat8_IsEqual(A8, b8)
'        Case 9: b = Mat9_IsEqual(A9, b9)
'        Case 10: b = Mat10_IsEqual(A10, b10)
'        End Select
'    Case " det "
'        Dim d As Double, d2 As Double
'        Select Case maxrc
'        Case 2: d = Mat2_det(A2)
'        Case 3: d = Mat3_det(A3) ': d2 = Mat3_det2(A3)
'        Case 4: d = Mat4_det(A4)
'        Case 5: d = Mat5_det(A5)
'        Case 6: d = Mat6_det(A6)
'        Case 7: d = Mat7_det(A7)
'        Case 8: d = Mat8_det(A8)
'        Case 9: d = Mat9_det(A9)
'        Case 10: d = Mat10_det(A10)
'        End Select
'    Case " adj "
'        Select Case maxrc
'        Case 2: c2 = Mat2_Adj(A2):  LSet C10 = c2
'        Case 3: C3 = Mat3_Adj(A3):  LSet C10 = C3
'        Case 4: C4 = Mat4_Adj(A4):  LSet C10 = C4
'        Case 5: C5 = Mat5_Adj(A5):  LSet C10 = C5
'        Case 6: C6 = Mat6_Adj(A6):  LSet C10 = C6
'        Case 7: C7 = Mat7_Adj(A7):  LSet C10 = C7
'        Case 8: C8 = Mat8_Adj(A8):  LSet C10 = C8
'        Case 9: C9 = Mat9_Adj(A9):  LSet C10 = C9
'        Case 10: C10 = Mat10_Adj(A10)
'        End Select
'    Case " inv "
'        Select Case maxrc
'        Case 2: c2 = Mat2_inv(A2):  LSet C10 = c2
'        Case 3: C3 = Mat3_inv(A3):  LSet C10 = C3
'        Case 4: C4 = Mat4_inv(A4):  LSet C10 = C4
'        Case 5: C5 = Mat5_inv(A5):  LSet C10 = C5
'        Case 6: C6 = Mat6_inv(A6):  LSet C10 = C6
'        Case 7: C7 = Mat7_inv(A7):  LSet C10 = C6
'        Case 8: C8 = Mat8_inv(A8):  LSet C10 = C6
'        Case 9: C9 = Mat9_inv(A9):  LSet C10 = C6
'        Case 10: C10 = Mat10_inv(A10)
'        End Select
'    Case "A*x=b"
'        'z.B.: A[-2 -2 | -1  1] * x[-2 | 0] = b [ 4 | 2 ];
'        'z.B.: A[ 5  6 |  3  4] * x[-13|12] = b [ 7 | 8 ];
'        Select Case maxrc
'        Case 2: vb2 = Mat2_Col(Mat2_Parse(Text3.Text), 0)
'                vx2 = Mat2_solve(A2, vb2)
'        Case 3: vb3 = Mat3_Col(Mat3_Parse(Text3.Text), 0)
'                vx3 = Mat3_solve(A3, vb3)
'        Case 4: vb4 = Mat4_Col(Mat4_Parse(Text3.Text), 0)
'                vx4 = Mat4_solve(A4, vb4)
'        Case 5: vb5 = Mat5_Col(Mat5_Parse(Text3.Text), 0)
'                vx5 = Mat5_solve(A5, vb5)
'        Case 6: vb6 = Mat6_Col(Mat6_Parse(Text3.Text), 0)
'                vx6 = Mat6_solve(A6, vb6)
'        Case 7: vb7 = Mat7_Col(Mat7_Parse(Text3.Text), 0)
'                vx7 = Mat7_solve(A7, vb7)
'        Case 8: vb8 = Mat8_Col(Mat8_Parse(Text3.Text), 0)
'                vx8 = Mat8_solve(A8, vb8)
'        Case 9: vb9 = Mat9_Col(Mat9_Parse(Text3.Text), 0)
'                vx9 = Mat9_solve(A9, vb9)
'        Case 10: vb10 = Mat10_Col(Mat10_Parse(Text3.Text), 0)
'                 vx10 = Mat10_solve(A10, vb10)
'        End Select
'    Case "eigval"
'        Select Case maxrc
'        Case 2: vb2 = Mat2_Eigenvalues(Mat2_Parse(Text1.Text))
'        Case 3: vb3 = Mat3_Eigenvalues(Mat3_Parse(Text1.Text))
'        Case Else
'        End Select
''    Case "cross"
''        Select Case maxrc
''        'Case 2: vb2 =
''        Case 3:
''        Case Else
''        End Select
'    'Case "Len-V"
'
'    'Case "NormV"
'
'    End Select
'    'Dim A_mRows As Long, A_nCols As Long
'    'Dim B_mRows As Long, B_nCols As Long
'
'    Select Case maxrc
'    Case 2:  Text1.Text = MMatrices.Mat2_ToStr(A2, A_mRows, A_nCols): Text2.Text = MMatrices.Mat2_ToStr(b2, B_mRows, B_nCols)
'    Case 3:  Text1.Text = MMatrices.Mat3_ToStr(A3, A_mRows, A_nCols): Text2.Text = MMatrices.Mat3_ToStr(b3, B_mRows, B_nCols)
'    Case 4:  Text1.Text = MMatrices.Mat4_ToStr(A4, A_mRows, A_nCols): Text2.Text = MMatrices.Mat4_ToStr(b4, B_mRows, B_nCols)
'    Case 5:  Text1.Text = MMatrices.Mat5_ToStr(A5, A_mRows, A_nCols): Text2.Text = MMatrices.Mat5_ToStr(b5, B_mRows, B_nCols)
'    Case 6:  Text1.Text = MMatrices.Mat6_ToStr(A6, A_mRows, A_nCols): Text2.Text = MMatrices.Mat6_ToStr(b6, B_mRows, B_nCols)
'    Case 7:  Text1.Text = MMatrices.Mat7_ToStr(A7, A_mRows, A_nCols): Text2.Text = MMatrices.Mat7_ToStr(b7, B_mRows, B_nCols)
'    Case 8:  Text1.Text = MMatrices.Mat8_ToStr(A8, A_mRows, A_nCols): Text2.Text = MMatrices.Mat8_ToStr(b8, B_mRows, B_nCols)
'    Case 9:  Text1.Text = MMatrices.Mat9_ToStr(A9, A_mRows, A_nCols): Text2.Text = MMatrices.Mat9_ToStr(b9, B_mRows, B_nCols)
'    Case 10: Text1.Text = MMatrices.Mat10_ToStr(A10, A_mRows, A_nCols): Text2.Text = MMatrices.Mat10_ToStr(b10, B_mRows, B_nCols)
'    End Select
'    'If Not op = "smul" Then
'    'End If
'    Dim T As String
'    If op = "  =  " Then
'        T = b
'        Text3.Text = T
'    ElseIf op = " det " Then
'        T = Str(d) '& vbCrLf & Str(d2)
'        Text3.Text = T
'    ElseIf op = "A*x=b" Then
'        Select Case maxrc
'        Case 2: T = Vec2_ToStr(vx2, False)
'        Case 3: T = Vec3_ToStr(vx3, False)
'        Case 4: T = Vec4_ToStr(vx4, False)
'        Case 5: T = Vec5_ToStr(vx5, False)
'        Case 6: T = Vec6_ToStr(vx6, False)
'        Case 7: T = Vec7_ToStr(vx7, False)
'        Case 8: T = Vec8_ToStr(vx8, False)
'        Case 9: T = Vec9_ToStr(vx9, False)
'        Case 10: T = Vec10_ToStr(vx10, False)
'        End Select
'        Text2.Text = T
'    ElseIf op = "eigval" Then
'        Select Case maxrc
'        Case 2: Text2.Text = MVector.Vec2_ToStr(vb2)
'        Case 3: Text2.Text = MVector.Vec3_ToStr(vb3)
'        End Select
'    Else
'        'hmmm wieso so??????
'        'weil entwede Matrix oder Vektor
'        'die Ausgabe der Matrix sollte nach SpaltenÄ/Zeilen erfolgen
'        'und nicht einfach nur mit 0en auffüllen
'        'OK Anzahl spalten richtig aber spalten nicht untereinander
'        T = Matrix_ToStr(VarPtr(C10), maxrc, maxrc)
'        Text3.Text = T
'        'hier spalten zwar untereinander aber Anzahl spalten nicht richtig
'        'Text3.Text = MMatrices.Mat10_ToStr(C10)
'
'        'hmm wie soll mans machen?
'        '
'        'also macn braucht eine Funktion die eine beliebige große matrix anzeigt mit den spalten richtig untereinander
'        'aber soviele Spalten und Zeilen wie erforderlich und nciht mehr!
'    End If
'
'End Sub

