VERSION 5.00
Begin VB.Form frmItemGenerico 
   BorderStyle     =   0  'None
   Caption         =   "Acesso Genérico a Item"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Controle_de_Estoque.xpcmdbutton b_fechar 
      Height          =   375
      Left            =   3990
      TabIndex        =   14
      Top             =   2220
      Width           =   1215
      _ExtentX        =   2143
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
   Begin Controle_de_Estoque.xpcmdbutton b_excluir 
      Height          =   375
      Left            =   2730
      TabIndex        =   13
      Top             =   2220
      Width           =   1215
      _ExtentX        =   2143
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
   Begin Controle_de_Estoque.xpcmdbutton b_incluir 
      Height          =   375
      Left            =   1470
      TabIndex        =   12
      Top             =   2220
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Incluir"
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
   Begin Controle_de_Estoque.xpcmdbutton b_alterar 
      Height          =   375
      Left            =   210
      TabIndex        =   11
      Top             =   2220
      Width           =   1215
      _ExtentX        =   2143
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
   Begin Controle_de_Estoque.xpcmdbutton b_last 
      Height          =   315
      Left            =   4800
      TabIndex        =   10
      Top             =   600
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      Caption         =   ">>"
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
   Begin Controle_de_Estoque.xpcmdbutton b_next 
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      Top             =   600
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      Caption         =   ">"
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
   Begin Controle_de_Estoque.xpcmdbutton b_prior 
      Height          =   315
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      Caption         =   "<"
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
   Begin Controle_de_Estoque.xpcmdbutton b_first 
      Height          =   315
      Left            =   3360
      TabIndex        =   7
      Top             =   600
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      Caption         =   "<<"
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
   Begin VB.TextBox qtdtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1710
      TabIndex        =   5
      Top             =   1710
      Width           =   2295
   End
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1710
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox itemtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1710
      TabIndex        =   2
      Top             =   1335
      Width           =   2295
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   4740
      Picture         =   "frmItemGenerico.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   5115
      Picture         =   "frmItemGenerico.frx":0A55
      Top             =   2730
      Width           =   285
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmItemGenerico.frx":1365
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmItemGenerico.frx":1E0F
      Top             =   2745
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Qtd Estoque:"
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
      Left            =   450
      TabIndex        =   4
      Top             =   1740
      Width           =   1080
   End
   Begin VB.Label Label3 
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
      Left            =   450
      TabIndex        =   1
      Top             =   990
      Width           =   630
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
      Left            =   450
      TabIndex        =   0
      Top             =   1335
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmItemGenerico.frx":26B5
      Top             =   150
      Width           =   105
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   5295
      Picture         =   "frmItemGenerico.frx":301F
      Top             =   450
      Width           =   105
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   4980
      Picture         =   "frmItemGenerico.frx":39AE
      Top             =   0
      Width           =   420
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acesso Genérico a Item"
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
      Width           =   2250
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   120
      Picture         =   "frmItemGenerico.frx":4458
      Top             =   2670
      Width           =   8505
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmItemGenerico.frx":530A
      Top             =   0
      Width           =   8505
   End
End
Attribute VB_Name = "frmItemGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsItem As ADODB.Recordset

Private Sub b_alterar_Click()
  If codtxt <> Empty Then
    SQLString = "UPDATE tab_item SET "
    SQLString = SQLString & " item = '" & itemtxt.Text & "' "
    SQLString = SQLString & " WHERE cod_item = " & codtxt.Text & " "
    fecharRS
    rs.Open SQLString, Con
    MsgBox "Registro Salvo!", vbInformation
    atualiza
  End If
End Sub

Private Sub b_fechar_Click()
  Unload Me
End Sub


Private Sub CarregaCod()
  fecharRS
  SQLString = "SELECT MAX(cod_item+1) as COD from tab_item "
  rs.Open SQLString, Con
  codtxt.Text = rs!COD
End Sub


Private Sub b_incluir_Click()
  CarregaCod
  If codtxt.Text <> Empty Then
    SQLString = "INSERT INTO tab_item (cod_item) VALUES (" & codtxt.Text & ")"
    fecharRS
    rs.Open SQLString, Con
    atualiza
  End If
End Sub

Private Sub b_excluir_Click()
  If codtxt.Text <> Empty Then
    SQLString = "DELETE FROM tab_item WHERE cod_item = " & Val(codtxt.Text)
    fecharRS
    rs.Open SQLString, Con
    MsgBox "Item Excluído", vbExclamation, "Mensagem"
    b_first_Click
    atualiza
  End If
End Sub

Private Sub b_first_Click()
  rsItem.MoveFirst
  PopulaCampos
End Sub

Private Sub b_last_Click()
  rsItem.MoveLast
  PopulaCampos
End Sub

Private Sub b_next_Click()
  rsItem.MoveNext
  PopulaCampos
End Sub

Private Sub b_prior_Click()
  rsItem.MovePrevious
  PopulaCampos
End Sub

Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18
  
  SQLString = "SELECT * FROM tab_item"
  Set rsItem = New ADODB.Recordset
  rsItem.Open SQLString, Con

  If rsItem.EOF Or rsItem.BOF Then
    MsgBox "Registro não Encontrado!"
  Else
    rsItem.MoveFirst
    PopulaCampos
  End If

End Sub

Private Sub codtxt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    itemtxt.SetFocus
    KeyAscii = 0
  End If
End Sub

Private Sub PopulaCampos()
  If Not rsItem.EOF And Not rsItem.BOF Then
    codtxt.Text = rsItem!cod_item
    itemtxt.Text = Trim(rsItem!Item & "")
    qtdtxt.Text = rsItem!qtd_estoque
  End If
  
  If rsItem.EOF Then
    b_last.Enabled = False
    b_next.Enabled = False
  Else
    b_last.Enabled = True
    b_next.Enabled = True
  End If
  
  If rsItem.BOF Then
    b_first.Enabled = False
    b_prior.Enabled = False
  Else
    b_first.Enabled = True
    b_prior.Enabled = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsItem.Close
End Sub

Private Sub atualiza()
  rsItem.Requery
  rsItem.MoveFirst
  While Trim(rsItem!cod_item) <> Trim(codtxt.Text)
    rsItem.MoveNext
  Wend
  PopulaCampos
End Sub

Private Sub picX_Click()
Unload Me
End Sub
