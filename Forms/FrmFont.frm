VERSION 5.00
Begin VB.Form FrmFont 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Font"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbFonts 
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1920
      Width           =   2415
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Lucida Console"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Courier New"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Consolas"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Test:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label LblTestFont2 
      AutoSize        =   -1  'True
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   8
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label LblTestFont1 
      AutoSize        =   -1  'True
      Caption         =   "----------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "FrmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbr As VbMsgBoxResult
Private m_Font As StdFont
Private bInit As Boolean

Public Function ShowDialog(owner As Form, font_inout As StdFont) As VbMsgBoxResult
    Set m_Font = font_inout
    Me.Show vbModal, owner
    ShowDialog = mbr
End Function

Private Sub Form_Load()
    bInit = True
    mbr = vbCancel
    CmbFonts_FillDicts
    Set LblTestFont1.Font = m_Font
    Set LblTestFont2.Font = LblTestFont1.Font
    MyFont = m_Font.Name
    bInit = False
End Sub

Private Sub CmbFonts_Click()
    If bInit Then Exit Sub
    Option4.Value = True
    LblTestFont1.Font.Name = CmbFonts.Text
    LblTestFont1.Font.Italic = False
End Sub

Private Sub CmbFonts_GotFocus()
    CmbFonts_Click
End Sub

Private Sub CmbFonts_FillDicts()
    Dim AllFontNames As Collection: Set AllFontNames = MFont.GetAllFonts(Me.hDC)
    Debug.Print AllFontNames.Count
    Dim v, sfnm As String
    CmbFonts.Visible = False
    For Each v In AllFontNames
        sfnm = v
        If Len(sfnm) Then
            If Font_HasEqualDickts(sfnm) Then
                CmbFonts.AddItem sfnm
            End If
        End If
    Next
    CmbFonts.Visible = True
    CmbFonts.Text = m_Font.Name
    Set CmbFonts.Font = m_Font
    'CmbFonts.ListIndex = 0
End Sub

Function Font_HasEqualDickts(ByVal fntnam As String) As Boolean
    'the fonts are synchronized, which means both labels share the same font-object
    LblTestFont1.Font.Name = fntnam
    LblTestFont1.Font.Italic = False
    LblTestFont2.Font.Name = fntnam
    LblTestFont2.Font.Italic = False
    Font_HasEqualDickts = (LblTestFont1.Width = LblTestFont2.Width)
End Function

Private Property Let MyFont(Value As String)
    'Set m_Font = New StdFont
    'm_FontName = Value
    Select Case True
    Case Value = Option1.Caption: Option1.Value = True
    Case Value = Option2.Caption: Option2.Value = True
    Case Value = Option3.Caption: Option3.Value = True
    Case Else
        CmbFonts.Text = Value
        Option4.Value = True
        Font_HasEqualDickts Value
    End Select
End Property

Private Property Get MyFont() As String
    Select Case True
    Case Option1.Value: MyFont = Option1.Font.Name
    Case Option2.Value: MyFont = Option2.Font.Name
    Case Option3.Value: MyFont = Option3.Font.Name
    Case Option4.Value: MyFont = CmbFonts.Text 'LblTestFont1.Font.Name
    End Select
End Property

Private Sub BtnOK_Click()
    mbr = vbOK
    Unload Me
End Sub
Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub Option1_Click()
    StdFont_Assign LblTestFont1.Font, Option1.Font
End Sub
Private Sub Option2_Click()
    StdFont_Assign LblTestFont1.Font, Option2.Font
End Sub
Private Sub Option3_Click()
    StdFont_Assign LblTestFont1.Font, Option3.Font
End Sub

Private Sub Option4_Click()
    LblTestFont1.Font.Name = Option4.Caption
End Sub

Private Sub Option1_DblClick()
    BtnOK_Click
End Sub

Private Sub Option2_DblClick()
    BtnOK_Click
End Sub

Private Sub Option3_DblClick()
    BtnOK_Click
End Sub

Private Sub Option4_DblClick()
    BtnOK_Click
End Sub

