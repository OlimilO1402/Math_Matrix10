VERSION 5.00
Begin VB.Form FrmCreateMatVec 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Create Mat/Vec"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PnlMatVec 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   3360
      ScaleHeight     =   2115
      ScaleWidth      =   4155
      TabIndex        =   16
      Top             =   120
      Width           =   4215
      Begin VB.TextBox LblMatVec 
         BorderStyle     =   0  'Kein
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   17
         Text            =   "FrmCreateMatVec.frx":0000
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.TextBox TxtRndTo 
      Alignment       =   2  'Zentriert
      Height          =   315
      Left            =   2160
      TabIndex        =   15
      Tag             =   "Rnd-bis"
      Text            =   "200"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TxtRndFrom 
      Alignment       =   2  'Zentriert
      Height          =   315
      Left            =   2160
      TabIndex        =   14
      Tag             =   "Rnd-von"
      Text            =   "-200"
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton BtnPreview 
      Caption         =   "&Update Preview"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox ResizeGrip 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   10
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox CmbmRows 
      Height          =   315
      ItemData        =   "FrmCreateMatVec.frx":0006
      Left            =   720
      List            =   "FrmCreateMatVec.frx":0008
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox CmbnCols 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   540
      Width           =   855
   End
   Begin VB.ComboBox CmbCreate 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox CmbNk 
      Height          =   315
      ItemData        =   "FrmCreateMatVec.frx":000A
      Left            =   720
      List            =   "FrmCreateMatVec.frx":000C
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "to:"
      Height          =   195
      Left            =   1800
      TabIndex        =   13
      Top             =   1020
      Width           =   180
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "from:"
      Height          =   195
      Left            =   1800
      TabIndex        =   12
      Top             =   600
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "mRows:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   180
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "nCols:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Values:"
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   180
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Format:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   525
   End
End
Attribute VB_Name = "FrmCreateMatVec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_MatVec As CMatOp
Private mbr As VbMsgBoxResult
Private mbt As Single
Private bInit As Boolean
Private m_FCSetMat As FrmCreateSetting

'OK, also wie wollen wir das machen, was wollen wir denn überhaupt?
'wir wollen
' * dass beim ersten Aufruf des Dialogs ohne dass bereits eine Matrix besteht
'   sinnvolle mögliche Voreinstellungen gewählt sind
'   was sind diese Voreinstellungen?
'   nCol und mRow sollen zufällig sein
'   Wertebereich von -200 bis +200
' * dass die Einstellungen die der User zuletzt gemacht hat,
'   mit dem nächsten Dialog wieder gestartet werden können
' * dass Die Einstellungen aus der aktuellen Matrix gelesen werden
'   falls zwischendrin eine ganz andere Matrix eingegeben wurde
Friend Function ShowDialog(owner As Form, FCSetMat As FrmCreateSetting, aMatVec_out As CMatOp) As VbMsgBoxResult
    FormInit FCSetMat
    Set m_MatVec = aMatVec_out
    Me.Show vbModal, owner
    ShowDialog = mbr
    If Not ShowDialog = vbOK Then Set m_MatVec = Nothing
    Set aMatVec_out = m_MatVec
End Function

'Private Sub Form_Initialize()
'    FormInit
'End Sub

Private Sub FormInit(fcs As FrmCreateSetting)
    m_FCSetMat = fcs
    Dim i As Long
    CmbmRows.AddItem "Rnd"
    CmbnCols.AddItem "Rnd"
    For i = 1 To 10
        CmbmRows.AddItem i
        CmbnCols.AddItem i
    Next
    CmbmRows.Text = fcs.mRows
    CmbnCols.Text = fcs.nCols
    'bInit = True
'    With m_FCSetMat
'        .mRows = "Rnd"
'        .nCols = "Rnd"
'        .NumFrom = "-200"
'        .NumTo = "200"
'        .Format = "2"
'    End With
    For i = 0 To 10: CmbNk.AddItem i: Next
    CmbNk.AddItem "E" 'für wissenschaftliche Anzeige
    CmbNk.Text = fcs.Format
    'CmbNk.ListIndex = 2
    'For i = 2 To 10: CmbRC.AddItem i: Next
    'CmbRC.ListIndex = 0
    
    
    'CmbmRows.ListIndex = 0
    'CmbnCols.ListIndex = 0
    With CmbCreate
        .AddItem " Rnd "
        .AddItem " ij  "
        .AddItem "unity"
        .AddItem "all 1"
        .Text = fcs.Values
    End With
    'CmbCreate.ListIndex = 0
    TxtRndFrom.Text = "-200"
    TxtRndTo.Text = "200"
    'TxtMatVec.Visible = False
    mbr = vbCancel
    mbt = BtnOK.Top
    Set LblMatVec.Font = MMain.MyFont
    Randomize
    bInit = False
    UpdateView
End Sub
Private Sub BtnCancel_Click()
    If bInit Then Exit Sub
    mbr = vbCancel
    Set m_MatVec = Nothing
    Unload Me
End Sub

Private Sub BtnOK_Click()
    If bInit Then Exit Sub
    mbr = vbOK
    'Set m_MatVec = GetMatVec
    Unload Me
End Sub

Private Sub BtnPreview_Click()
    If bInit Then Exit Sub
    Set m_MatVec = GetMatVec
    UpdateView
End Sub

Private Sub CmbNk_Click()
    If bInit Then Exit Sub
    UpdateView
End Sub

Function CmbGetValue(aCmb As ComboBox) As Long
    If bInit Then Exit Function
    If LCase(aCmb.Text) = "rnd" Then
        CmbGetValue = Max(1, Rnd * 10)
    Else
        CmbGetValue = CLng(aCmb.ListIndex)
    End If
End Function

Private Sub CmbmRows_Click()
    If bInit Then Exit Sub
    BtnPreview_Click
End Sub
Private Sub CmbnCols_Click()
    If bInit Then Exit Sub
    BtnPreview_Click
End Sub
Private Sub CmbCreate_Click()
    If bInit Then Exit Sub
    Dim bEnb As Boolean: bEnb = CmbCreate.Text = " Rnd "
    TxtRndFrom.Enabled = bEnb
    TxtRndTo.Enabled = bEnb
    UpdateView
End Sub

Sub UpdateView()
    If bInit Then Exit Sub
    If m_MatVec Is Nothing Then Set m_MatVec = GetMatVec
    Dim fmt As Integer: fmt = Get_Format
    If m_MatVec Is Nothing Then Exit Sub
    LblMatVec.Text = m_MatVec.ToStr(fmt)
End Sub

Function GetMatVec() As CMatOp
    If bInit Then Exit Function
    Dim r As Integer:    r = Get_mRows
    Dim c As Integer:    c = Get_nCols
    Dim dfr As Double: dfr = GetRndFrom
    Dim dto As Double: dto = GetRndTo
    Dim nk  As Byte:    nk = Get_Format
    Set GetMatVec = MVMFactory.CMatRnd(r, c, dfr, dto)
    If GetMatVec Is Nothing Then Exit Function
    GetMatVec.nk = nk
End Function

Function Get_mRows() As Long
    Get_mRows = CmbGetValue(CmbmRows)
End Function

Function Get_nCols() As Long
    Get_nCols = CmbGetValue(CmbnCols)
End Function
Function Get_Format() As Integer
    Dim s As String: s = CmbNk.Text
    If Trim(UCase(s)) = "E" Then Get_Format = 255 Else If IsNumeric(s) Then Get_Format = CByte(s)
End Function

Function GetRndFrom() As Double
    GetRndFrom = TB_ParseDouble(TxtRndFrom)
End Function
Function GetRndTo() As Double
    GetRndTo = TB_ParseDouble(TxtRndTo)
End Function
Function TB_ParseDouble(TB As TextBox) As Double
    Dim d As Double
    If Double_TryParse(TB.Text, d) Then
        TB_ParseDouble = d
    Else
        MsgBox "Geben Sie eine gültige zahl ein für " & TB.Tag & "."
    End If
End Function

Private Sub Form_Resize()
    
    ResizeGrip.Move Me.ScaleWidth - ResizeGrip.Width, Me.ScaleHeight - ResizeGrip.Height
    
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim l As Single: l = PnlMatVec.Left
    Dim T As Single: T = PnlMatVec.Top
    Dim W As Single: W = Me.ScaleWidth - l - brdr '/ 4
    Dim H As Single: H = Me.ScaleHeight - T - brdr '/ 4
    If W > 0 And H > 0 Then
        PnlMatVec.Move l, T, W, H
        LblMatVec.Move brdr, brdr, W - 2 * brdr, H - 2 * brdr
    End If
    T = Max(mbt, Me.ScaleHeight - brdr - BtnOK.Height)
    BtnOK.Top = T
    BtnCancel.Top = T
    Dim cbrdr As Single: cbrdr = Me.Height - Me.ScaleHeight
    If Abs(Me.ScaleHeight - (mbt + brdr + BtnOK.Height)) < 180 Then
        Me.Height = (mbt + brdr + BtnOK.Height) + cbrdr
        Exit Sub
    End If
    
End Sub

