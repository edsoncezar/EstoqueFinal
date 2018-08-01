VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmsaldos 
   BorderStyle     =   0  'None
   Caption         =   "Consulta Entradas/Saídas "
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Controle_de_Estoque.xpcmdbutton fechar 
      Height          =   375
      Left            =   3390
      TabIndex        =   11
      Top             =   2730
      Width           =   1005
      _ExtentX        =   1773
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
   Begin Controle_de_Estoque.xpcmdbutton saidas 
      Height          =   375
      Left            =   2340
      TabIndex        =   10
      Top             =   2730
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      Caption         =   "&Saídas"
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
   Begin Controle_de_Estoque.xpcmdbutton entradas 
      Height          =   375
      Left            =   1290
      TabIndex        =   9
      Top             =   2730
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      Caption         =   "&Entradas"
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
   Begin Controle_de_Estoque.xpcmdbutton listar1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2730
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      Caption         =   "&Totais"
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
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   4530
      TabIndex        =   1
      Top             =   600
      Width           =   3255
      _Version        =   524288
      _ExtentX        =   5741
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2003
      Month           =   11
      Day             =   5
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2775
      Left            =   210
      TabIndex        =   0
      Top             =   3150
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Tipo de Movimento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   270
      TabIndex        =   2
      Top             =   720
      Width           =   2175
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808080&
         Caption         =   "Entradas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808080&
         Caption         =   "Saídas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Pesquisa por produto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   270
      TabIndex        =   3
      Top             =   1920
      Width           =   3975
      Begin VB.ComboBox itemcombo 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   7530
      Picture         =   "frmsaldos.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image10 
      Height          =   4245
      Left            =   8048
      Picture         =   "frmsaldos.frx":0A55
      Top             =   1800
      Width           =   105
   End
   Begin VB.Image Image9 
      Height          =   4245
      Left            =   0
      Picture         =   "frmsaldos.frx":13E4
      Top             =   1830
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmsaldos.frx":1D4E
      Top             =   6045
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmsaldos.frx":25F4
      Top             =   420
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmsaldos.frx":2F5E
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   8055
      Picture         =   "frmsaldos.frx":3A08
      Top             =   420
      Width           =   105
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   7740
      Picture         =   "frmsaldos.frx":4397
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   7875
      Picture         =   "frmsaldos.frx":4E41
      Top             =   6030
      Width           =   285
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta Entradas/Saídas "
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
      Width           =   2475
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -570
      Picture         =   "frmsaldos.frx":5751
      Top             =   0
      Width           =   8505
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   120
      Picture         =   "frmsaldos.frx":6EC6
      Top             =   5970
      Width           =   8505
   End
End
Attribute VB_Name = "frmsaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bb As String
Private Sub Calendar1_Click()
fecharRS
bb = Calendar1.Value
If Option1.Value = True Then
  bc = "E"
  SQLString = "SELECT e.cod_movto,  i.item, "
  SQLString = SQLString & "IIF(e.tipo_movto = 'E','Entrada', 'Saída') AS tipo, qtd_movto, dat_movto "
  SQLString = SQLString & "FROM estoque e, tab_item i WHERE e.cod_item = i.cod_item "
  SQLString = SQLString & " and dat_movto& '' = '" & bb & "' and tipo_movto='" & bc & "'"
Else
  If Option2.Value = True Then
    bc = "S"
    SQLString = "SELECT e.cod_movto,  i.item, "
    SQLString = SQLString & "IIF(e.tipo_movto = 'E','Entrada', 'Saída') AS tipo, qtd_movto,dat_movto "
    SQLString = SQLString & "FROM estoque e, tab_item i WHERE e.cod_item = i.cod_item "
    SQLString = SQLString & " and  dat_movto&'' = '" & bb & "' and tipo_movto='" & bc & "'"
  Else
    SQLString = "SELECT e.cod_movto,  i.item, "
    SQLString = SQLString & "IIF(e.tipo_movto = 'E','Entrada', 'Saída') AS tipo, qtd_movto,dat_movto "
    SQLString = SQLString & "FROM estoque e, tab_item i WHERE e.cod_item = i.cod_item "
    SQLString = SQLString & " and dat_movto&'' = '" & bb & "' "
  End If
End If

rs.Open SQLString, Con
If rs.EOF Then
  MsgBox "Nenhum Registro Encontrado", vbInformation, "Mensagem"
Else
  If rs.RecordCount > 0 Then
    flex
  End If
End If
End Sub

Private Sub entradas_Click()
fecharRS
MSHFlexGrid1.Clear
SQLString = "SELECT e.cod_movto,  i.item, "
SQLString = SQLString & "IIF(e.tipo_movto = 'E','Entrada', 'Saída') AS tipo, qtd_movto,dat_movto "
SQLString = SQLString & "FROM estoque e, tab_item i WHERE e.cod_item = i.cod_item "
SQLString = SQLString & " and tipo_movto='" & "E" & "'"

rs.Open SQLString, Con

If rs.EOF Then
  MsgBox "Nenhuma Movimentação encontrada", vbExclamation, "Mensagem"
Else
  flex
End If
End Sub

Private Sub fechar_Click()
Unload Me
End Sub



Private Sub itemcombo_click()
SQLString = "select cod_item from tab_item where item='" & itemcombo.Text & "'"
fecharRS
rs.Open SQLString, Con
bb = Val(rs!cod_item)

SQLString = "SELECT e.cod_movto,  i.item, "
SQLString = SQLString & "IIF(e.tipo_movto = 'E','Entrada', 'Saída') AS tipo, qtd_movto, dat_movto "
SQLString = SQLString & "FROM estoque e, tab_item i WHERE e.cod_item = i.cod_item "
SQLString = SQLString & " and e.cod_item = " & bb & ""

fecharRS
rs.Open SQLString, Con

If rs.EOF Then
  MsgBox "Nenhum registro encontrado", vbExclamation, "Mensagem"
Else
  Set MSHFlexGrid1.DataSource = rs
Call flex
End If
End Sub

Private Sub listar1_click()
fecharRS
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 6
SQLString = "SELECT e.cod_movto,  i.item, "
SQLString = SQLString & "IIF(e.tipo_movto = 'E','Entrada', 'Saída') AS tipo, qtd_movto,dat_movto "
SQLString = SQLString & "FROM estoque e, tab_item i WHERE e.cod_item = i.cod_item "

rs.Open SQLString, Con
If rs.EOF Then
  MsgBox "Nenhuma Movimentação encontrada", vbExclamation, "Mensagem"
Else
  Set MSHFlexGrid1.DataSource = rs
  Call flex
End If
End Sub

Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18
fecharRS
SQLString = "select item from tab_item"
rs.Open SQLString, Con
If Not rs.EOF Then
  rs.MoveFirst
  While rs.EOF = False
    itemcombo.AddItem (rs!Item)
    rs.MoveNext
  Wend
End If
End Sub

Private Sub picX_Click()
Unload Me
End Sub

Private Sub saidas_Click()
fecharRS
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 6
SQLString = "select * from estoque where "

SQLString = "SELECT e.cod_movto,  i.item, "
SQLString = SQLString & "IIF(e.tipo_movto = 'E','Entrada', 'Saída') AS tipo, qtd_movto, dat_movto "
SQLString = SQLString & "FROM estoque e, tab_item i WHERE e.cod_item = i.cod_item "
SQLString = SQLString & " and tipo_movto='" & "S" & "'"
rs.Open SQLString, Con

If rs.EOF Then
  MsgBox "Nenhuma Movimentação encontrada", vbExclamation, "Mensagem"
Else
  Set MSHFlexGrid1.DataSource = rs
  Call flex
End If

End Sub

Private Function flex()
    Set MSHFlexGrid1.DataSource = rs
    MSHFlexGrid1.TextMatrix(0, 0) = "Código"
    MSHFlexGrid1.TextMatrix(0, 1) = "Item"
    MSHFlexGrid1.TextMatrix(0, 2) = "Tipo Movimento"
    MSHFlexGrid1.TextMatrix(0, 3) = "Quantidade"
    MSHFlexGrid1.TextMatrix(0, 4) = "Data"
End Function
