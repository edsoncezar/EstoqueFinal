VERSION 5.00
Begin VB.Form frmAlteracaoMovto 
   BorderStyle     =   0  'None
   Caption         =   "Alteração de Movimentos de Estoque"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Controle_de_Estoque.xpcmdbutton f_fechar 
      Height          =   375
      Left            =   4320
      TabIndex        =   16
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
      Left            =   2580
      TabIndex        =   15
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
   Begin Controle_de_Estoque.xpcmdbutton b_salvar 
      Height          =   375
      Left            =   810
      TabIndex        =   14
      Top             =   2880
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
      Left            =   4920
      TabIndex        =   12
      Top             =   2370
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
      Height          =   195
      Left            =   2310
      TabIndex        =   10
      Top             =   1560
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
      Left            =   1110
      TabIndex        =   9
      Top             =   1530
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox datatxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1110
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1110
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox qtdMovtotxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   2370
      Width           =   1215
   End
   Begin VB.TextBox codItemtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   3690
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1110
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   5700
      Picture         =   "frmAlteracaoMovto.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   6075
      Picture         =   "frmAlteracaoMovto.frx":0A55
      Top             =   3360
      Width           =   285
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmAlteracaoMovto.frx":1365
      Top             =   3375
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   6255
      Picture         =   "frmAlteracaoMovto.frx":1C0B
      Top             =   450
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmAlteracaoMovto.frx":259A
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmAlteracaoMovto.frx":3044
      Top             =   150
      Width           =   105
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   5940
      Picture         =   "frmAlteracaoMovto.frx":39AE
      Top             =   0
      Width           =   420
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
      Left            =   330
      TabIndex        =   17
      Top             =   210
      Width           =   3525
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
      Left            =   2430
      TabIndex        =   13
      Top             =   2415
      Width           =   2415
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
      Left            =   285
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
      Left            =   285
      TabIndex        =   8
      Top             =   1125
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
      Left            =   285
      TabIndex        =   6
      Top             =   765
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
      Left            =   285
      TabIndex        =   4
      Top             =   2430
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
      Left            =   285
      TabIndex        =   3
      Top             =   1965
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmAlteracaoMovto.frx":4458
      Top             =   0
      Width           =   8505
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   240
      Picture         =   "frmAlteracaoMovto.frx":5BCD
      Top             =   3300
      Width           =   8505
   End
End
Attribute VB_Name = "frmAlteracaoMovto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsItem As ADODB.Recordset
Dim qtdAnt As Integer


Private Sub b_limpar_Click()
  cboItem.SelText = ""
  codItemtxt.Text = ""
  qtdMovtotxt.Text = ""
  datatxt.Text = ""
  codtxt.Text = ""
  codtxt.SetFocus
  b_salvar.Enabled = False
End Sub


Private Sub b_salvar_Click()
Dim tipo As String

  If Val(qtdMovtotxt.Text) > 0 Then
  
    If radEntrada.Value = True Then
      tipo = "E"
    Else
      tipo = "S"
    End If
    
    If (tipo = "S" And Val(qtdMovtotxt.Text) > (Val(qtdDisp.Text) + qtdAnt)) Then
      MsgBox "A quantidade do movimento é maior do que o disponivel em estoque. Corrija!", vbExclamation
      Exit Sub
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
  SQLString = "SELECT e.cod_movto,  e.cod_item, e.dat_movto,  e.tipo_movto,  i.item,  e.qtd_movto "
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

    cboItem.Text = ""
    cboItem.SelText = rs!Item
    qtdAnt = rs!qtd_movto
    qtdMovtotxt.Text = rs!qtd_movto
    codItemtxt.Text = rs!cod_item

    b_salvar.Enabled = True
    cboItem.SetFocus
  
    SQLString = "SELECT "
    SQLString = SQLString & "  i.item, "
    SQLString = SQLString & "SUM(IIF(e.tipo_movto = 'S', qtd_Movto * (-1) , qtd_Movto)) AS saldo "
    SQLString = SQLString & "FROM tab_item i, "
    SQLString = SQLString & "  Estoque e "
    SQLString = SQLString & "WHERE e.cod_item = i.cod_item "
    SQLString = SQLString & "  and e.cod_item = " & codItemtxt.Text & " "
    SQLString = SQLString & "GROUP BY item "
    SQLString = SQLString & "ORDER BY item "

    fecharRS
    rs.Open SQLString, Con

    If rs.EOF Or rs.BOF Then
      qtdDisp.Text = "0"
    Else
      qtdDisp.Text = rs!Saldo
    End If
  Else
    MsgBox "Registro não encontrado!", vbInformation
    b_salvar.Enabled = False
    b_limpar_Click
  End If
End If
End Sub

Private Sub f_fechar_Click()
  codtxt.Text = ""
  cboItem.Text = ""
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
  Conexao
  Set rsItem = New ADODB.Recordset
  SQLString = "SELECT * FROM tab_item ORDER BY item"
  fecharRS
  rsItem.Open SQLString, Con
  cboItem.Clear
  cboItem.Text = ""
  
  While Not rsItem.EOF
    cboItem.AddItem (rsItem!Item)
    rsItem.MoveNext
  Wend
End Sub



Private Sub picX_Click()
Unload Me
End Sub
