VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MsFlxGrd.Ocx"
Begin VB.Form consulta 
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox codtxt 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   4215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "consulta.frx":0000
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4260
      _Version        =   393216
      ScrollTrack     =   -1  'True
      MousePointer    =   1
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Dagoberto\Meus documentos\thiago\Meus arquivos recebidos\projetolp\base_estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "estoque"
      Top             =   3000
      Width           =   4095
   End
   Begin VB.ListBox List1 
      Columns         =   1
      DataField       =   "cod_item"
      DataSource      =   "Data1"
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   5055
   End
   Begin VB.Menu a 
      Caption         =   "&Consulta"
      Begin VB.Menu b 
         Caption         =   "Item"
         Shortcut        =   ^T
      End
      Begin VB.Menu c 
         Caption         =   "Fabricante"
         Shortcut        =   ^F
      End
      Begin VB.Menu j 
         Caption         =   "Código"
         Shortcut        =   ^O
      End
      Begin VB.Menu d 
         Caption         =   "Itens em limite"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu e 
      Caption         =   "&Tabelas"
      Begin VB.Menu f 
         Caption         =   "Produtos"
         Shortcut        =   ^P
      End
      Begin VB.Menu g 
         Caption         =   "Empregados"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu h 
      Caption         =   "&Sobre"
      Begin VB.Menu i 
         Caption         =   "Ajuda"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_Click()
codtxt.Visible = True
Data1.Visible = False
MSFlexGrid1.Visible = False
List1.Visible = True
Frame1.Visible = True
Frame1.Caption = "Item a ser Pesquisado"
End Sub
Private Sub codtxt_Change()
rs.Close
SQLString = "select * from estoque "
rs.Open SQLString, Con
If rs.EOF Then
Else
While rs.EOF = False
List1.AddItem (rs!coditem & " " & rs!Item & " " & rs!qtd_estoque)
rs.MoveNext
Wend
End If
End Sub
Private Sub Data1_Validate(Action As Integer, Save As Integer)
List1.Clear
End Sub

Private Sub f_Click()
Data1.Visible = True
MSFlexGrid1.Visible = True
List1.Visible = False
codtxt.Visible = False
Frame1.Visible = False
Data1.RecordSource = estoque
End Sub

Private Sub Form_Load()
Conexao
List1.Visible = True
codtxt.Visible = True
Frame1.Visible = True
MSFlexGrid1.Visible = False
Data1.Visible = True
SQLString = "select * from tab_item "
rs.Open SQLString, Con
rs.MoveFirst
If rs.EOF = True Then
Else
rs.MoveFirst
While rs.EOF = False
List1.AddItem (rs!iten)
rs.MoveNext
Wend
End If
End Sub

