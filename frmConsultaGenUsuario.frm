VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCGUsuario 
   BorderStyle     =   0  'None
   Caption         =   "Consulta Genérica Usuários"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Controle_de_Estoque.xpcmdbutton cmdVoltar 
      Height          =   345
      Left            =   270
      TabIndex        =   0
      Top             =   4020
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      Caption         =   "&Voltar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3195
      Left            =   270
      TabIndex        =   1
      Top             =   690
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   5636
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   6390
      Picture         =   "frmConsultaGenUsuario.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmConsultaGenUsuario.frx":0A55
      Top             =   4665
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   6735
      Picture         =   "frmConsultaGenUsuario.frx":12FB
      Top             =   4650
      Width           =   285
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta Genérica de Usuários"
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
      TabIndex        =   2
      Top             =   210
      Width           =   2925
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   6600
      Picture         =   "frmConsultaGenUsuario.frx":1C0B
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   6915
      Picture         =   "frmConsultaGenUsuario.frx":26B5
      Top             =   420
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmConsultaGenUsuario.frx":3044
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmConsultaGenUsuario.frx":3AEE
      Top             =   420
      Width           =   105
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   120
      Picture         =   "frmConsultaGenUsuario.frx":4458
      Top             =   4590
      Width           =   8505
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmConsultaGenUsuario.frx":530A
      Top             =   0
      Width           =   8505
   End
End
Attribute VB_Name = "frmCGusuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVoltar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18

SQLString = "select cod_usuario, nome, cod_perfil, ativo from tab_usuario"
fecharRS
rs.Open SQLString, Con
Set MSHFlexGrid1.DataSource = rs
MSHFlexGrid1.TextMatrix(0, 0) = "Código"
MSHFlexGrid1.TextMatrix(0, 1) = "Usuário"
MSHFlexGrid1.TextMatrix(0, 2) = "Perfil"
MSHFlexGrid1.TextMatrix(0, 3) = "Ativo"
End Sub

Private Sub MSHFlexGrid1_DblClick()
  If MsgBox("Deseja editar o usuário seleciondo?", vbYesNo) = vbYes Then
    frmusuariogenerico.codtxt = Str(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 0))
    frmusuariogenerico.Show 1
  End If
End Sub

Private Sub picX_Click()
Unload Me
End Sub
