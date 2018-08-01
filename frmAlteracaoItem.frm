VERSION 5.00
Begin VB.Form frmAlteracaoItem 
   BorderStyle     =   0  'None
   Caption         =   "Alteração de Item"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Controle_de_Estoque.xpcmdbutton b_Fechar 
      Height          =   375
      Left            =   3540
      TabIndex        =   6
      Top             =   2190
      Width           =   1425
      _extentx        =   2514
      _extenty        =   661
      caption         =   "&Fechar"
      font            =   "frmAlteracaoItem.frx":0000
   End
   Begin Controle_de_Estoque.xpcmdbutton b_Limpar 
      Height          =   375
      Left            =   1965
      TabIndex        =   5
      Top             =   2190
      Width           =   1425
      _extentx        =   2514
      _extenty        =   661
      caption         =   "&Limpar"
      font            =   "frmAlteracaoItem.frx":002C
   End
   Begin Controle_de_Estoque.xpcmdbutton b_Alterar 
      Height          =   375
      Left            =   390
      TabIndex        =   4
      Top             =   2190
      Width           =   1425
      _extentx        =   2514
      _extenty        =   661
      caption         =   "&Alterar"
      font            =   "frmAlteracaoItem.frx":0058
   End
   Begin VB.TextBox itemtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1140
      TabIndex        =   1
      Top             =   1170
      Width           =   2745
   End
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1140
      TabIndex        =   0
      Top             =   750
      Width           =   2745
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alteração de Itens"
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
      Width           =   1725
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   4740
      Picture         =   "frmAlteracaoItem.frx":0084
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmAlteracaoItem.frx":0AD9
      Top             =   2745
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   5115
      Picture         =   "frmAlteracaoItem.frx":137F
      Top             =   2730
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   4980
      Picture         =   "frmAlteracaoItem.frx":1C8F
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   5295
      Picture         =   "frmAlteracaoItem.frx":2739
      Top             =   450
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmAlteracaoItem.frx":30C8
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmAlteracaoItem.frx":3B72
      Top             =   150
      Width           =   105
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmAlteracaoItem.frx":44DC
      Top             =   0
      Width           =   8505
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
   Begin VB.Image Image5 
      Height          =   450
      Left            =   120
      Picture         =   "frmAlteracaoItem.frx":5C51
      Top             =   2670
      Width           =   8505
   End
End
Attribute VB_Name = "frmAlteracaoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub b_alterar_Click()
If codtxt <> Empty Then
  SQLString = "UPDATE tab_item SET "
  SQLString = SQLString & " item = '" & itemtxt.Text & "' "
  SQLString = SQLString & " WHERE cod_item = " & codtxt.Text & " "
  fecharRS
  rs.Open SQLString, Con
  MsgBox "Registro Salvo!", vbInformation
End If
End Sub

Private Sub b_limpar_Click()
  codtxt.Text = ""
  itemtxt.Text = ""
  b_Alterar.Enabled = False
  codtxt.SetFocus
End Sub

Private Sub b_fechar_Click()
Unload Me
End Sub

Private Sub codtxt_LostFocus()

If codtxt.Text <> Empty Then
  SQLString = "SELECT * FROM tab_item WHERE cod_item = " & Str(Val(codtxt.Text))
  fecharRS
  rs.Open SQLString, Con
  
  If rs.EOF Or rs.BOF Then
    MsgBox "Registro não Encontrado!"
    b_Alterar.Enabled = False
    b_limpar_Click
  Else
    b_Alterar.Enabled = True
    itemtxt.Text = rs!Item
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
