VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEntradaSaida 
   BorderStyle     =   0  'None
   Caption         =   "Entradas/Saídas"
   ClientHeight    =   7635
   ClientLeft      =   7410
   ClientTop       =   0
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Controle_de_Estoque.xpcmdbutton b_fechar 
      Height          =   375
      Left            =   4020
      TabIndex        =   28
      Top             =   4080
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
   Begin Controle_de_Estoque.xpcmdbutton b_excluir 
      Height          =   375
      Left            =   2520
      TabIndex        =   27
      Top             =   4080
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
   Begin Controle_de_Estoque.xpcmdbutton Command1 
      Height          =   375
      Left            =   1020
      TabIndex        =   26
      Top             =   4080
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      Caption         =   "xpcmdbutton1"
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
   Begin Controle_de_Estoque.xpcmdbutton b_mov 
      Height          =   405
      Left            =   4200
      TabIndex        =   25
      Top             =   1860
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   714
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
   Begin Controle_de_Estoque.xpcmdbutton b_confirmar 
      Height          =   375
      Left            =   1440
      TabIndex        =   24
      Top             =   3240
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   661
      Caption         =   "&Confirmar"
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
      Left            =   6570
      TabIndex        =   23
      Top             =   660
      Width           =   615
      _ExtentX        =   1085
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
      Left            =   5940
      TabIndex        =   22
      Top             =   660
      Width           =   615
      _ExtentX        =   1085
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
      Left            =   5310
      TabIndex        =   21
      Top             =   660
      Width           =   615
      _ExtentX        =   1085
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
      Left            =   4680
      TabIndex        =   20
      Top             =   660
      Width           =   615
      _ExtentX        =   1085
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
   Begin VB.TextBox codmovtotxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1830
      TabIndex        =   18
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox datatxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1830
      TabIndex        =   15
      Top             =   2820
      Width           =   2295
   End
   Begin VB.TextBox usuariotxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1830
      TabIndex        =   14
      Top             =   2460
      Width           =   2295
   End
   Begin VB.TextBox qtd1txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1830
      TabIndex        =   11
      Top             =   1740
      Width           =   2295
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2415
      Left            =   540
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4530
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.OptionButton saida 
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
      Height          =   495
      Left            =   5940
      TabIndex        =   8
      Top             =   3210
      Width           =   1095
   End
   Begin VB.OptionButton entrada 
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
      Height          =   495
      Left            =   4740
      TabIndex        =   7
      Top             =   3210
      Width           =   975
   End
   Begin VB.ListBox listaprod 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1785
      Left            =   4620
      TabIndex        =   6
      Top             =   1050
      Width           =   2655
   End
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1830
      TabIndex        =   5
      Top             =   1020
      Width           =   855
   End
   Begin VB.TextBox qtdtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1830
      TabIndex        =   4
      Top             =   2100
      Width           =   2295
   End
   Begin VB.TextBox itemtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1830
      TabIndex        =   2
      Top             =   1380
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Tipo de Operação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4620
      TabIndex        =   9
      Top             =   2970
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Alteração"
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
      Left            =   300
      TabIndex        =   13
      Top             =   3810
      Width           =   7095
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entradas e Saídas"
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
      Left            =   390
      TabIndex        =   29
      Top             =   210
      Width           =   1710
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   6900
      Picture         =   "frmentradasaida.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   395
      Left            =   7228
      Picture         =   "frmentradasaida.frx":0A55
      Top             =   7238
      Width           =   285
   End
   Begin VB.Image Image11 
      Height          =   4245
      Left            =   7410
      Picture         =   "frmentradasaida.frx":1365
      Top             =   3000
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmentradasaida.frx":1CF4
      Top             =   7258
      Width           =   255
   End
   Begin VB.Image Image10 
      Height          =   4245
      Left            =   0
      Picture         =   "frmentradasaida.frx":259A
      Top             =   3060
      Width           =   105
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmentradasaida.frx":2F04
      Top             =   570
      Width           =   105
   End
   Begin VB.Image Image9 
      Height          =   4245
      Left            =   7405
      Picture         =   "frmentradasaida.frx":386E
      Top             =   570
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmentradasaida.frx":41FD
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   7088
      Picture         =   "frmentradasaida.frx":4CA7
      Top             =   0
      Width           =   420
   End
   Begin VB.Label label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Movimentação:"
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
      Left            =   420
      TabIndex        =   19
      Top             =   690
      Width           =   1290
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   420
      TabIndex        =   17
      Top             =   2850
      Width           =   435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
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
      Left            =   420
      TabIndex        =   16
      Top             =   2490
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
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
      Left            =   420
      TabIndex        =   12
      Top             =   1770
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   420
      TabIndex        =   3
      Top             =   2130
      Width           =   1080
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
      Left            =   420
      TabIndex        =   1
      Top             =   1050
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
      Left            =   420
      TabIndex        =   0
      Top             =   1410
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   0
      Picture         =   "frmentradasaida.frx":5751
      Top             =   0
      Width           =   8505
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   0
      Picture         =   "frmentradasaida.frx":6EC6
      Top             =   7190
      Width           =   8505
   End
End
Attribute VB_Name = "frmentradasaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim data, en As String
Dim QTDE, ab As Integer
Dim a, x, y, ex As Integer
Dim bb As Integer
Dim rsItem As ADODB.Recordset

Private Sub b_alterar_Click()
If codtxt.Text <> Empty And itemtxt <> Empty And qtd1txt.Text <> Empty Then
SQLString = "update estoque set cod_item=" & Val(codtxt.Text) & ", qtd_movto=" & Val(qtd1txt.Text) & " where cod_movto=" & (codmovtotxt.Text) & ""
fecharRS
rs.Open SQLString, Con
MsgBox "Registro Alterado", vbExclamation, "Mensagem"
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
If ex = 1 And codtxt.Text <> Empty And itemtxt.Text <> Empty And codmovtotxt.Text <> Empty Then
SQLString = "select qtd_estoque from tab_item where item ='" & itemtxt.Text & "'"
fecharRS
rs.Open SQLString, Con
a = rs!qtd_estoque
    If entrada.Value = True Then
    SQLString = "update tab_item set qtd_estoque=" & Val(MSHFlexGrid1.TextMatrix(x, 3)) - a & ""
    Else
    If saida.Value = True Then
    SQLString = "update tab_item set qtd_estoque=" & Val(MSHFlexGrid1.TextMatrix(x, 3)) + a & ""
    End If
    End If
    fecharRS
    rs.Open SQLString, Con
    SQLString = "DELETE FROM estoque WHERE cod_movto = " & Val(codmovtotxt.Text)
    fecharRS
    rs.Open SQLString, Con
    MsgBox "Item Excluído", vbExclamation, "Mensagem"
    b_first_Click
    atualiza
    codmovtotxt.Text = Empty
    itemtxt.Text = Empty
    qtd1txt.Text = Empty
    usuariotxt.Text = Empty
    datatxt.Text = Empty
            End If
     SQLString = "select * from estoque"
fecharRS
rs.Open SQLString, Con
If rs.RecordCount <= 0 Then
codmovtotxt.Text = 1
Else
SQLString = "select max(cod_movto) as QTDE From estoque"
fecharRS
rs.Open SQLString, Con
codmovtotxt.Text = rs!QTDE + 1
End If
End Sub

Private Sub b_first_Click()
  rsItem.MoveFirst
  
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

Private Sub b_mov_Click()
If listaprod.Text <> Empty Then
SQLString = "select * from tab_item where item='" & listaprod.Text & "'"
fecharRS
rs.Open SQLString, Con
codtxt.Text = rs!cod_item
itemtxt.Text = rs!Item
qtdtxt.Text = rs!qtd_estoque
qtd1txt.Text = Empty
qtd1txt.SetFocus
ex = 0

End If
End Sub

Private Sub b_confirmar_Click()
ex = 0
data = Date
If codtxt.Text <> Empty Then
If entrada.Value = True Then
en = "E"
bb = Val(qtd1txt.Text) + Val(qtdtxt.Text)
Else
If saida.Value = True Then
en = "S"
If Val(qtd1txt.Text) > Val(qtdtxt.Text) Then
MsgBox "Não há essa quantidade em estoque", vbExclamation, "Mensagem"
qtd1txt.Text = Empty
GoTo aa
End If
bb = Val(qtdtxt.Text) - Val(qtd1txt.Text)
Else
MsgBox "Escolha o tipo de operação"
GoTo aa
End If
End If
SQLString = "insert into estoque values(" & Val(codmovtotxt.Text) & ", " & Val(codtxt.Text) & ", " & "1" & "," & Val(qtd1txt.Text) & ", '" & en & "','" & data & "')"
fecharRS
rs.Open SQLString, Con
SQLString = "update tab_item set qtd_estoque=" & bb & " where cod_item =" & Val(codtxt.Text) & ""
rs.Open SQLString, Con
MsgBox "Movimentação Inserida", 0, "Mensagem"
codmovtotxt.Text = Val(codmovtotxt.Text) + 1
End If
atualiza
aa:
End Sub

Private Sub MSHFlexGrid1_Click()
SQLString = "select * from estoque"
fecharRS
rs.Open SQLString, Con
If rs.RecordCount > 0 Then
x = MSHFlexGrid1.RowSel
If x = 0 Then
x = 1
End If
codmovtotxt.Text = MSHFlexGrid1.TextMatrix(x, 0)
codtxt.Text = MSHFlexGrid1.TextMatrix(x, 1)
ab = MSHFlexGrid1.TextMatrix(x, 1)
SQLString = "select * from tab_item where cod_item = " & ab & ""
fecharRS
rs.Open SQLString, Con
itemtxt.Text = rs!Item
qtd1txt.Text = MSHFlexGrid1.TextMatrix(x, 3)
usuariotxt.Text = MSHFlexGrid1.TextMatrix(x, 2)
datatxt.Text = MSHFlexGrid1.TextMatrix(x, 5)
If MSHFlexGrid1.TextMatrix(x, 4) = "E" Then
entrada.Value = True
Else
saida.Value = True
End If
ex = 1
End If
End Sub

Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18

MSHFlexGrid1.Rows = 3
  SQLString = "SELECT * FROM tab_item"
  Set rsItem = New ADODB.Recordset
  rsItem.Open SQLString, Con

  If rsItem.EOF Or rsItem.BOF Then
    MsgBox "Registro não Encontrado!"
  Else
    rsItem.MoveFirst
    While rsItem.EOF = False
    listaprod.AddItem (rsItem!Item)
    rsItem.MoveNext
    Wend
        
  End If
SQLString = " select * from estoque"
fecharRS
rs.Open SQLString, Con
Set MSHFlexGrid1.DataSource = rs
MSHFlexGrid1.TextMatrix(0, 0) = "Código"
MSHFlexGrid1.TextMatrix(0, 1) = "Item"
MSHFlexGrid1.TextMatrix(0, 2) = "Usuário"
MSHFlexGrid1.TextMatrix(0, 3) = "Quantidade"
MSHFlexGrid1.TextMatrix(0, 4) = "Tipo Movimento"
SQLString = "select * from estoque"
fecharRS
rs.Open SQLString, Con
If rs.RecordCount = 0 Then
codmovtotxt.Text = 1
Else
SQLString = "select max(cod_movto) as QTDE From estoque"
fecharRS
rs.Open SQLString, Con
codmovtotxt.Text = rs!QTDE + 1
End If
datatxt = Date
End Sub

Private Sub codtxt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    itemtxt.SetFocus
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsItem.Close
End Sub

Private Sub atualiza()
SQLString = "select * from estoque"
fecharRS
rs.Open SQLString, Con

  Set MSHFlexGrid1.DataSource = rs
  MSHFlexGrid1.TextMatrix(0, 0) = "Código"
MSHFlexGrid1.TextMatrix(0, 1) = "Item"
MSHFlexGrid1.TextMatrix(0, 2) = "Usuário"
MSHFlexGrid1.TextMatrix(0, 3) = "Quantidade"
MSHFlexGrid1.TextMatrix(0, 4) = "Tipo Movimento"
   

End Sub


Private Sub picX_Click()
Unload Me
End Sub
