VERSION 5.00
Begin VB.Form frmExclusaoItem 
   BorderStyle     =   0  'None
   Caption         =   "Exclusão de Item"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Controle_de_Estoque.xpcmdbutton b_Excluir 
      Height          =   375
      Left            =   390
      TabIndex        =   6
      Top             =   2190
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      Caption         =   "&Excluir"
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
   Begin Controle_de_Estoque.xpcmdbutton b_limpar 
      Height          =   375
      Left            =   1965
      TabIndex        =   5
      Top             =   2190
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      Caption         =   "&Limpar"
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
   Begin Controle_de_Estoque.xpcmdbutton b_Fechar 
      Height          =   375
      Left            =   3540
      TabIndex        =   4
      Top             =   2190
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      Caption         =   "&Fechar"
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
   Begin VB.TextBox itemtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1140
      TabIndex        =   1
      Top             =   1170
      Width           =   2145
   End
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1140
      TabIndex        =   0
      Top             =   750
      Width           =   2175
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exclusão de Itens"
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
      Width           =   1665
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmExclusaoItem.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   4740
      Picture         =   "frmExclusaoItem.frx":0AAA
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmExclusaoItem.frx":14FF
      Top             =   2745
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   5115
      Picture         =   "frmExclusaoItem.frx":1DA5
      Top             =   2730
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   300
      TabIndex        =   3
      Top             =   780
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   300
      TabIndex        =   2
      Top             =   1200
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   5295
      Picture         =   "frmExclusaoItem.frx":26B5
      Top             =   480
      Width           =   105
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmExclusaoItem.frx":3044
      Top             =   180
      Width           =   105
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   4980
      Picture         =   "frmExclusaoItem.frx":39AE
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmExclusaoItem.frx":4458
      Top             =   0
      Width           =   8505
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   150
      Picture         =   "frmExclusaoItem.frx":5BCD
      Top             =   2670
      Width           =   8505
   End
End
Attribute VB_Name = "frmexclusaoitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub b_excluir_Click()

If codtxt.Text <> Empty Then
  SQLString = "Delete FROM tab_item WHERE cod_item = " & Val(codtxt.Text)
  fecharRS
  rs.Open SQLString, Con
  MsgBox "Registro Excluido!", vbExclamation
  b_limpar_Click
End If

End Sub

Private Sub b_fechar_Click()
  codtxt.Text = ""
  Unload Me
End Sub

Private Sub b_limpar_Click()
  codtxt.Text = ""
  itemtxt.Text = ""
  codtxt.SetFocus
  b_excluir.Enabled = False
End Sub

Private Sub codtxt_LostFocus()

If codtxt.Text <> Empty Then
  SQLString = "SELECT * FROM tab_item WHERE cod_item = " & Val(codtxt.Text)
  fecharRS
  rs.Open SQLString, Con

  If rs.EOF Or rs.BOF Then
    MsgBox "Registro não Encontrado!"
    b_excluir.Enabled = False
    b_limpar_Click
  Else
    itemtxt.Text = rs!Item
    b_excluir.Enabled = True
  End If
End If

End Sub

Private Sub Form_Activate()
  b_limpar_Click
End Sub

Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18
End Sub

Private Sub picX_Click()
Unload Me
End Sub
