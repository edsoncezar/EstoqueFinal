VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSaldoAtual 
   BorderStyle     =   0  'None
   Caption         =   "Saldo Atual por Produto"
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3285
      Left            =   270
      TabIndex        =   0
      Top             =   720
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Controle_de_Estoque.xpcmdbutton v_voltar 
      Height          =   345
      Left            =   300
      TabIndex        =   2
      Top             =   4140
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
   Begin VB.Image picX 
      Height          =   315
      Left            =   6390
      Picture         =   "frmSaldoAtual.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta Genérica de Itens"
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
      TabIndex        =   1
      Top             =   210
      Width           =   2565
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   6735
      Picture         =   "frmSaldoAtual.frx":0A55
      Top             =   4650
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   6600
      Picture         =   "frmSaldoAtual.frx":1365
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   6915
      Picture         =   "frmSaldoAtual.frx":1E0F
      Top             =   420
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmSaldoAtual.frx":279E
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmSaldoAtual.frx":3248
      Top             =   420
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmSaldoAtual.frx":3BB2
      Top             =   4665
      Width           =   255
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmSaldoAtual.frx":4458
      Top             =   0
      Width           =   8505
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   120
      Picture         =   "frmSaldoAtual.frx":5BCD
      Top             =   4590
      Width           =   8505
   End
End
Attribute VB_Name = "frmSaldoAtual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18
SQLString = "SELECT "
SQLString = SQLString & "  i.item, "
SQLString = SQLString & "SUM(IIF(e.tipo_movto = 'S', qtd_Movto * (-1) , qtd_Movto)) AS saldo "
SQLString = SQLString & "FROM tab_item i, "
SQLString = SQLString & "  Estoque e "
SQLString = SQLString & "Where e.cod_item = i.cod_item "
SQLString = SQLString & "GROUP BY item "
SQLString = SQLString & "ORDER BY item "

fecharRS
rs.Open SQLString, Con

Set MSHFlexGrid1.DataSource = rs

MSHFlexGrid1.TextMatrix(0, 0) = "Item"
MSHFlexGrid1.TextMatrix(0, 1) = "Saldo Atual"

MSHFlexGrid1.ColWidth(0) = 4000
MSHFlexGrid1.ColWidth(1) = 2000

End Sub

Private Sub picX_Click()
Unload Me
End Sub

Private Sub v_voltar_Click()
  Unload Me
End Sub
