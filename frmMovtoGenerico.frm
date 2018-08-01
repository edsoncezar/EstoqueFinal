VERSION 5.00
Begin VB.Form frmMovtoGenerico 
   BorderStyle     =   0  'None
   Caption         =   "Movimentos de Estoque"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4890
      TabIndex        =   21
      Top             =   2310
      Width           =   1215
   End
   Begin VB.OptionButton radSaida 
      BackColor       =   &H00808080&
      Caption         =   "Saída"
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
      Left            =   2340
      TabIndex        =   10
      Top             =   1575
      Width           =   975
   End
   Begin VB.OptionButton radEntrada 
      BackColor       =   &H00808080&
      Caption         =   "Entrada"
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
      Left            =   1140
      TabIndex        =   9
      Top             =   1545
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox datatxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1140
      TabIndex        =   7
      Top             =   1110
      Width           =   2295
   End
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1140
      TabIndex        =   5
      Top             =   750
      Width           =   975
   End
   Begin VB.TextBox qtdMovtotxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   2310
      Width           =   1215
   End
   Begin VB.TextBox codItemtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3660
      TabIndex        =   2
      Top             =   1890
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   1890
      Width           =   2415
   End
   Begin Controle_de_Estoque.xpcmdbutton b_first 
      Height          =   315
      Left            =   4260
      TabIndex        =   13
      Top             =   630
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
   Begin Controle_de_Estoque.xpcmdbutton b_prior 
      Height          =   315
      Left            =   4740
      TabIndex        =   14
      Top             =   630
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
   Begin Controle_de_Estoque.xpcmdbutton b_next 
      Height          =   315
      Left            =   5220
      TabIndex        =   15
      Top             =   630
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
   Begin Controle_de_Estoque.xpcmdbutton b_last 
      Height          =   315
      Left            =   5700
      TabIndex        =   16
      Top             =   630
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
   Begin Controle_de_Estoque.xpcmdbutton b_salvar 
      Height          =   375
      Left            =   780
      TabIndex        =   17
      Top             =   2850
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
   Begin Controle_de_Estoque.xpcmdbutton b_incluir 
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   2850
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
   Begin Controle_de_Estoque.xpcmdbutton b_excluir 
      Height          =   375
      Left            =   3300
      TabIndex        =   19
      Top             =   2850
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
   Begin Controle_de_Estoque.xpcmdbutton b_fechar 
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   2850
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
      Left            =   2400
      TabIndex        =   22
      Top             =   2355
      Width           =   2415
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   5700
      Picture         =   "frmMovtoGenerico.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   6075
      Picture         =   "frmMovtoGenerico.frx":0A55
      Top             =   3360
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   5940
      Picture         =   "frmMovtoGenerico.frx":1365
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   6260
      Picture         =   "frmMovtoGenerico.frx":1E0F
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmMovtoGenerico.frx":279E
      Top             =   3375
      Width           =   255
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Movimentos de Estoque"
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
      TabIndex        =   12
      Top             =   210
      Width           =   2250
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmMovtoGenerico.frx":3044
      Top             =   150
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmMovtoGenerico.frx":39AE
      Top             =   0
      Width           =   345
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
      Left            =   375
      TabIndex        =   11
      Top             =   1560
      Width           =   405
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
      Left            =   375
      TabIndex        =   8
      Top             =   1155
      Width           =   435
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
      Left            =   375
      TabIndex        =   6
      Top             =   795
      Width           =   630
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
      Left            =   375
      TabIndex        =   4
      Top             =   2355
      Width           =   450
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
      Left            =   375
      TabIndex        =   3
      Top             =   1935
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmMovtoGenerico.frx":4458
      Top             =   0
      Width           =   8505
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   240
      Picture         =   "frmMovtoGenerico.frx":5BCD
      Top             =   3300
      Width           =   8505
   End
End
Attribute VB_Name = "frmMovtoGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMovto As ADODB.Recordset
Dim rsItem As ADODB.Recordset


Private Sub b_excluir_Click()
  If codtxt.Text <> Empty Then
    SQLString = "DELETE FROM estoque WHERE cod_movto = " & Val(codtxt.Text)
    fecharRS
    rs.Open SQLString, Con
    MsgBox "Movimento Excluído", vbExclamation, "Mensagem"
    b_first_Click
    preenche
    atualiza
  End If
End Sub

Private Sub b_incluir_Click()
  CarregaCod
  If codtxt.Text <> Empty Then
    SQLString = "INSERT INTO estoque (cod_movto, cod_usuario, cod_item, dat_movto) VALUES (" & codtxt.Text & ", " & cod_usuario & ", " & codItemtxt.Text & " ,'" & Date & "')"
    fecharRS
    rs.Open SQLString, Con
    atualiza
  End If
End Sub

Private Sub CarregaCod()
  fecharRS
  SQLString = "SELECT MAX(cod_movto+1) as COD from estoque "
  rs.Open SQLString, Con
  If rs!COD & "" <> Empty Then
    codtxt.Text = rs!COD
  Else
    codtxt.Text = "1"
  End If
End Sub



Private Sub b_fechar_Click()
Unload Me
End Sub

Private Sub b_first_Click()
  rsMovto.MoveFirst
  preenche
End Sub

Private Sub b_last_Click()
  rsMovto.MoveLast
  preenche
End Sub

Private Sub b_next_Click()
  rsMovto.MoveNext
  preenche
End Sub

Private Sub b_prior_Click()
  rsMovto.MovePrevious
  preenche
End Sub

Private Sub b_salvar_Click()
Dim tipo As String

  If Val(qtdMovtotxt.Text) > 0 Then
    If radEntrada.Value = True Then
      tipo = "E"
    Else
      tipo = "S"
    End If
      
    SQLString = "UPDATE estoque SET "
    SQLString = SQLString & "cod_item = " & codItemtxt.Text & ", "
    SQLString = SQLString & "cod_usuario = " & cod_usuario & ", "
    SQLString = SQLString & "qtd_movto = " & qtdMovtotxt.Text & ", "
    SQLString = SQLString & "tipo_movto = '" & tipo & "' "
    SQLString = SQLString & "WHERE cod_movto = " & codtxt.Text & " "
    
    fecharRS
    rs.Open SQLString, Con
    
    MsgBox "Registro Alterado!", vbInformation
    atualiza
  Else
    MsgBox "Preencha a qtde do movimento!", vbExclamation
  End If
End Sub

Private Sub cboItem_LostFocus()
  rsItem.MoveFirst
  If Trim(cboItem.Text) <> Empty Then
    While (Not rsItem.EOF) And (Trim(rsItem!Item) <> Trim(cboItem.Text))
      rsItem.MoveNext
    Wend
    codItemtxt.Text = rsItem!cod_item
  End If
End Sub

Private Sub codtxt_LostFocus()

If codtxt.Text <> Empty Then
  SQLString = "SELECT count(*) as QT FROM estoque e, tab_item i WHERE e.cod_item = i.cod_item and e.cod_movto = " & Val(codtxt.Text)
  fecharRS
  rs.Open SQLString, Con
  
  If rs!QT > 0 Then
    atualiza
    preenche
  Else
    MsgBox "Registro não encontrado!", vbExclamation
  End If
End If

End Sub

Private Sub f_fechar_Click()
  Unload Me
End Sub





Private Sub Form_Activate()
  If codtxt.Text <> Empty Then
    atualiza
  End If
  preenche
End Sub

Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18
  
  Set rsItem = New ADODB.Recordset
  Set rsMovto = New ADODB.Recordset
  SQLString = "SELECT * FROM tab_item ORDER BY item"
  
  rsItem.Open SQLString, Con
  cboItem.Clear
  cboItem.Text = ""
  
  While Not rsItem.EOF
    cboItem.AddItem (rsItem!Item)
    rsItem.MoveNext
  Wend
  
  SQLString = "SELECT e.cod_movto,  e.cod_item, e.dat_movto,  e.tipo_movto,  i.item,  e.qtd_movto "
  SQLString = SQLString + " FROM estoque e, tab_item i "
  SQLString = SQLString + " WHERE e.cod_item = i.cod_item "
  rsMovto.Open SQLString, Con
  
End Sub


Private Sub preenche()

If Not rsMovto.EOF And Not rsMovto.BOF Then
  
  If Trim(rsMovto!tipo_movto) = "E" Then
    radEntrada.Value = True
  Else
    radSaida.Value = True
  End If
  
  codtxt.Text = (rsMovto!cod_movto)
  datatxt.Text = (rsMovto!dat_movto)
  cboItem.Text = ""
  cboItem.SelText = rsMovto!Item
  qtdMovtotxt.Text = rsMovto!qtd_movto
  codItemtxt.Text = rsMovto!cod_item
  
  cboItem.SetFocus
Else
  MsgBox "Não há registro não encontrado!", vbInformation
End If

End Sub


Private Sub atualiza()
  rsMovto.Requery
  rsMovto.MoveFirst
  While Trim(rsMovto!cod_movto) <> Trim(codtxt.Text)
    rsMovto.MoveNext
  Wend
  
  If rsMovto.EOF Then
    rsMovto.MoveFirst
  End If
  preenche
End Sub


Private Sub picX_Click()
Unload Me
End Sub
