VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   0  'None
   Caption         =   "Help"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Controle_de_Estoque.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   5100
      TabIndex        =   6
      Top             =   1110
      Width           =   1125
      _extentx        =   1984
      _extenty        =   661
      caption         =   "&Sair"
      font            =   "frmhelp.frx":0000
   End
   Begin VB.TextBox itemtxt 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   330
      TabIndex        =   1
      Top             =   1080
      Width           =   3210
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      Height          =   2790
      Left            =   345
      TabIndex        =   0
      Top             =   2040
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Digite o que deseja encontrar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   210
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   255
      TabIndex        =   3
      Top             =   1725
      Width           =   2460
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   2730
      TabIndex        =   4
      Top             =   1725
      Width           =   4185
      Begin VB.TextBox descricaotxt 
         BackColor       =   &H00E0E0E0&
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Image Image13 
      Height          =   4245
      Left            =   7160
      Picture         =   "frmhelp.frx":002C
      Top             =   960
      Width           =   105
   End
   Begin VB.Image Image12 
      Height          =   4245
      Left            =   0
      Picture         =   "frmhelp.frx":09BB
      Top             =   900
      Width           =   105
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   6630
      Picture         =   "frmhelp.frx":1325
      Top             =   150
      Width           =   315
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HELP - SIE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   330
      TabIndex        =   7
      Top             =   210
      Width           =   1005
   End
   Begin VB.Image Image9 
      Height          =   390
      Left            =   6975
      Picture         =   "frmhelp.frx":1D7A
      Top             =   5130
      Width           =   285
   End
   Begin VB.Image Image8 
      Height          =   570
      Left            =   6840
      Picture         =   "frmhelp.frx":268A
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   4245
      Left            =   7155
      Picture         =   "frmhelp.frx":3134
      Top             =   420
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmhelp.frx":3AC3
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmhelp.frx":456D
      Top             =   420
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmhelp.frx":4ED7
      Top             =   5145
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   7575
      Picture         =   "frmhelp.frx":577D
      Top             =   4650
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   7440
      Picture         =   "frmhelp.frx":608D
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   7760
      Picture         =   "frmhelp.frx":6B37
      Top             =   420
      Width           =   105
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   120
      Picture         =   "frmhelp.frx":74C6
      Top             =   5070
      Width           =   8505
   End
   Begin VB.Image Image11 
      Height          =   585
      Left            =   -840
      Picture         =   "frmhelp.frx":8378
      Top             =   0
      Width           =   8505
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim cd As Integer

Private Sub Form_Load()
  Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18
  
  SQLString = "select * from help"
  fecharRS
  rs.Open SQLString, Con
  
  If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
      List1.AddItem (rs!chave)
      rs.MoveNext
    Wend
  End If
End Sub

Private Sub List1_Click()
  SQLString = "select * from help where chave='" & List1.Text & "'"
  fecharRS
  rs.Open SQLString, Con
  
  If rs.RecordCount > 0 Then
    descricaotxt.Text = rs!descrição
  End If
End Sub

Private Sub itemtxt_Change()
  cd = Len(itemtxt)
  SQLString = "SELECT * FROM help WHERE chave LIKE '" & Trim(itemtxt.Text) & "%' "
  fecharRS
  rs.Open SQLString, Con
  List1.Clear
  
  If rs.RecordCount > 0 Then
    rs.MoveFirst
    While Not rs.EOF
      List1.AddItem (rs!chave)
      rs.MoveNext
    Wend
  End If
End Sub

Private Sub picX_Click()
  Unload Me
End Sub

Private Sub xpcmdbutton1_Click()
  Unload Me
End Sub
