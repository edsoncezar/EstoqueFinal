VERSION 5.00
Begin VB.Form frmExclusaoMovto 
   BorderStyle     =   0  'None
   Caption         =   "Exclusão de Movimentos de Estoque"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1230
      TabIndex        =   9
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox qtdMovtotxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1230
      TabIndex        =   8
      Top             =   2370
      Width           =   1215
   End
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1230
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox datatxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1230
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.OptionButton radEntrada 
      BackColor       =   &H00808080&
      Caption         =   "Entrada"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1230
      TabIndex        =   5
      Top             =   1530
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton radSaida 
      BackColor       =   &H00808080&
      Caption         =   "Saída"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2430
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox qtdDisp 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      Top             =   2370
      Width           =   1215
   End
   Begin Controle_de_Estoque.xpcmdbutton f_fechar 
      Height          =   375
      Left            =   4500
      TabIndex        =   0
      Top             =   2880
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
   Begin Controle_de_Estoque.xpcmdbutton b_limpar 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2880
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
   Begin Controle_de_Estoque.xpcmdbutton b_excluir 
      Height          =   375
      Left            =   1020
      TabIndex        =   2
      Top             =   2880
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
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alteração de Movimentos de Estoque"
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
      Left            =   360
      TabIndex        =   10
      Top             =   180
      Width           =   3525
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   6915
      Picture         =   "frmExclusaoMovto.frx":0000
      Top             =   3360
      Width           =   285
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmExclusaoMovto.frx":0910
      Top             =   3360
      Width           =   255
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   6540
      Picture         =   "frmExclusaoMovto.frx":11B6
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   6780
      Picture         =   "frmExclusaoMovto.frx":1C0B
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmExclusaoMovto.frx":26B5
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   60
      Picture         =   "frmExclusaoMovto.frx":315F
      Top             =   3300
      Width           =   8505
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   240
      Picture         =   "frmExclusaoMovto.frx":4011
      Top             =   0
      Width           =   8505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
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
      Left            =   405
      TabIndex        =   16
      Top             =   1965
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Qtde:"
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
      Left            =   405
      TabIndex        =   15
      Top             =   2430
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
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
      Left            =   405
      TabIndex        =   14
      Top             =   765
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Data:"
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
      Left            =   405
      TabIndex        =   13
      Top             =   1125
      Width           =   435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Tipo:"
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
      Left            =   405
      TabIndex        =   12
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Qtde Disponível em Estoque:"
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
      Left            =   2550
      TabIndex        =   11
      Top             =   2415
      Width           =   2415
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmExclusaoMovto.frx":5786
      Top             =   60
      Width           =   105
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   7095
      Picture         =   "frmExclusaoMovto.frx":60F0
      Top             =   540
      Width           =   105
   End
End
Attribute VB_Name = "frmExclusaoMovto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub b_excluir_Click()
If codtxt.Text <> Empty Then
  SQLString = "DELETE FROM estoque WHERE cod_movto = " & codtxt.Text & " "
  fecharRS
  rs.Open SQLString, Con
  MsgBox "Registro Excluído!", vbInformation
  b_limpar_Click
End If

End Sub

Private Sub b_limpar_Click()
  cboItem.SelText = ""
  qtdMovtotxt.Text = ""
  datatxt.Text = ""
  codtxt.SetFocus
End Sub


Private Sub codtxt_LostFocus()
If codtxt.Text <> Empty Then
SQLString = "SELECT e.cod_movto,  e.dat_movto,  e.tipo_movto,  i.item,  e.qtd_movto "
SQLString = SQLString + " FROM estoque e, tab_item i "
SQLString = SQLString + " WHERE e.cod_item = i.cod_item "
SQLString = SQLString + " AND e.cod_movto = " & Val(codtxt.Text) & " "

fecharRS
rs.Open SQLString, Con

If Not rs.EOF And Not rs.BOF Then
  datatxt.Text = (rs!dat_movto)
  
  If rs!tipo_movto = "E" Then
    radEntrada.Value = 1
  Else
    radSaida.Value = 1
  End If

  cboItem.Clear
  cboItem.AddItem (rs!Item)
  cboItem.SelText = rs!Item

  qtdMovtotxt.Text = rs!qtd_movto
Else
  MsgBox "Registro não encontrado!", vbInformation
End If
End If
End Sub

Private Sub f_fechar_Click()
  Unload Me
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
