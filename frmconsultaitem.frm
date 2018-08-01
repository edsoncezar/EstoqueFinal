VERSION 5.00
Begin VB.Form frmconsultaitem 
   BorderStyle     =   0  'None
   Caption         =   "Consulta Item"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Controle_de_Estoque.xpcmdbutton Command3 
      Height          =   375
      Left            =   1980
      TabIndex        =   9
      Top             =   2220
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
   Begin Controle_de_Estoque.xpcmdbutton Command2 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   2220
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
   Begin Controle_de_Estoque.xpcmdbutton Command1 
      Height          =   375
      Left            =   330
      TabIndex        =   7
      Top             =   2220
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      Caption         =   "&Salvar"
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
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      TabIndex        =   2
      Top             =   1590
      Width           =   945
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   1170
      Width           =   2595
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   750
      Width           =   945
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmconsultaitem.frx":0000
      Top             =   2745
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   5115
      Picture         =   "frmconsultaitem.frx":08A6
      Top             =   2730
      Width           =   285
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   4740
      Picture         =   "frmconsultaitem.frx":11B6
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   4980
      Picture         =   "frmconsultaitem.frx":1C0B
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmconsultaitem.frx":26B5
      Top             =   180
      Width           =   105
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   5295
      Picture         =   "frmconsultaitem.frx":301F
      Top             =   480
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmconsultaitem.frx":39AE
      Top             =   0
      Width           =   345
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
      TabIndex        =   6
      Top             =   210
      Width           =   1665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Código do Item:"
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
      TabIndex        =   5
      Top             =   780
      Width           =   1305
   End
   Begin VB.Label Label2 
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
      Left            =   300
      TabIndex        =   4
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Quantidade no Estoque :"
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
      Top             =   1620
      Width           =   2055
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmconsultaitem.frx":4458
      Top             =   0
      Width           =   8505
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   150
      Picture         =   "frmconsultaitem.frx":5BCD
      Top             =   2670
      Width           =   8505
   End
End
Attribute VB_Name = "frmconsultaitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
If Text1.Text <> Empty And Text2 <> Empty Then
SQLString = "select * from tab_item where iten='" & Text1.Text & "'"
fecharRS

rs.Open SQLString, Con

If rs.EOF Or rs.BOF Then
MsgBox "Registro não encontrado", vbExclamation, "Mensagem"
Else
Text1.Text = rs!iten
Text2.Text = rs!cod_item
Text3.Text = rs!qtd_estoque
End If
End If
End Sub
Private Sub Command2_Click()
 Unload Me
End Sub
Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub
Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18

Conexao
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub picX_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
KeyAscii = 0
End If

End Sub

Private Sub Text1_LostFocus()
SQLString = "SELECT * FROM tab_item WHERE cod_item = " & Val(Text1.Text)
fecharRS
rs.Open SQLString, Con

If rs.EOF Or rs.BOF Then
  MsgBox "Registro não Encontrado!"
Else
  Text1.Text = rs!cod_item
  Text2.Text = rs!Item
  Text3.Text = rs!qtd_estoque
   
  
  
  
End If

End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  Text3.SetFocus
  KeyAscii = 0
End If

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57, 8, 44
Case Else
KeyAscii = 0
End Select
End Sub
