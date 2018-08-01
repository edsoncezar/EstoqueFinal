VERSION 5.00
Begin VB.Form frmUsuarioGenerico 
   BorderStyle     =   0  'None
   Caption         =   "Movimentos de Estoque"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Ativo?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2520
      TabIndex        =   6
      Top             =   1020
      Width           =   840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Perfil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3720
      TabIndex        =   3
      Top             =   1020
      Width           =   1695
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808080&
         Caption         =   "Admnistrador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   240
         Width           =   1470
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808080&
         Caption         =   "Usuário"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.TextBox senhatxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox nometxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1620
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1620
      TabIndex        =   0
      Top             =   990
      Width           =   750
   End
   Begin Controle_de_Estoque.xpcmdbutton b_first 
      Height          =   315
      Left            =   3420
      TabIndex        =   10
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
      Left            =   3900
      TabIndex        =   11
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
      Left            =   4380
      TabIndex        =   12
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
      Left            =   4860
      TabIndex        =   13
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
      Left            =   210
      TabIndex        =   16
      Top             =   2550
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
   Begin Controle_de_Estoque.xpcmdbutton b_inclusao 
      Height          =   375
      Left            =   1470
      TabIndex        =   17
      Top             =   2550
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
      Left            =   2730
      TabIndex        =   18
      Top             =   2550
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
      Left            =   3990
      TabIndex        =   15
      Top             =   2550
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
   Begin VB.Image picX 
      Height          =   315
      Left            =   4920
      Picture         =   "frmUsuarioGenerico.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmUsuarioGenerico.frx":0A55
      Top             =   3045
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   5295
      Picture         =   "frmUsuarioGenerico.frx":12FB
      Top             =   3030
      Width           =   285
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmUsuarioGenerico.frx":1C0B
      Top             =   150
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmUsuarioGenerico.frx":2575
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   5475
      Picture         =   "frmUsuarioGenerico.frx":301F
      Top             =   330
      Width           =   105
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   5160
      Picture         =   "frmUsuarioGenerico.frx":39AE
      Top             =   0
      Width           =   420
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acesso Genérico a Usuário"
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
      TabIndex        =   14
      Top             =   210
      Width           =   2550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Senha "
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
      Left            =   180
      TabIndex        =   9
      Top             =   1980
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Nome "
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
      Left            =   180
      TabIndex        =   8
      Top             =   1500
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Código de Perfil"
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
      Left            =   180
      TabIndex        =   7
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   120
      Picture         =   "frmUsuarioGenerico.frx":4458
      Top             =   2970
      Width           =   8505
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmUsuarioGenerico.frx":530A
      Top             =   0
      Width           =   8505
   End
End
Attribute VB_Name = "frmusuariogenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim perfil As Integer
Dim ativo As String
Dim rsUsuario As ADODB.Recordset

Private Sub b_excluir_Click()
If codtxt.Text <> Empty Then
  SQLString = "Delete FROM tab_usuario WHERE cod_usuario = " & Val(codtxt.Text) & " or nome ='" & nometxt & "'"
  fecharRS
  rs.Open SQLString, Con
  MsgBox "Usuário Excluído", vbExclamation, "Mensagem"
  b_first_Click
End If
atualiza

End Sub



Private Sub b_first_Click()
rsUsuario.MoveFirst
preenche
End Sub

Private Sub b_inclusao_Click()
b_limpar_Click
CarregaCod
If codtxt.Text <> Empty Then
  SQLString = "select cod_usuario from tab_usuario where cod_usuario=" & Val(codtxt.Text) & ""
  fecharRS
  rs.Open SQLString, Con
  If rs.RecordCount > 0 Then
    MsgBox "Código já cadastrado", vbExclamation, "Mensagem"
  Else
    If Check1.Value = 1 Then
      ativo = "S"
    Else
      ativo = "N"
    End If

    If Option1.Value = True Then
      perfil = 1
    Else
      perfil = 2
    End If

    SQLString = "insert into tab_usuario (cod_usuario) values(" & codtxt.Text & ")"
    fecharRS
    rs.Open SQLString, Con
    MsgBox "Registro Inserido", 0, "OK"
  End If
End If
atualiza
End Sub

Private Sub b_last_Click()
  rsUsuario.MoveLast
  preenche
End Sub

Private Sub b_limpar_Click()
  codtxt.Text = ""
  nometxt.Text = ""
  senhatxt.Text = ""
  Option1.Value = False
  Option2.Value = False
  Check1.Value = False
  nometxt.SetFocus
End Sub

Private Sub b_fechar_Click()
 Unload Me
End Sub

Private Sub b_next_Click()
  rsUsuario.MoveNext
  preenche
End Sub

Private Sub b_prior_Click()
  rsUsuario.MovePrevious
  preenche
End Sub


Private Sub b_salvar_Click()
Dim ativo, perfil As String
If Check1.Value = 1 Then
  ativo = "S"
Else
  ativo = "N"
End If

If Option1.Value = True Then
  perfil = "1"
Else
  perfil = "2"
End If


SQLString = "UPDATE tab_usuario SET "
SQLString = SQLString & " nome = '" & nometxt.Text & "',"
SQLString = SQLString & " senha = '" & senhatxt.Text & "',"
SQLString = SQLString & " cod_perfil = '" & perfil & "',"
SQLString = SQLString & " ativo = '" & ativo & "' "
SQLString = SQLString & " WHERE cod_usuario = " & codtxt.Text & " "
fecharRS
rs.Open SQLString, Con
MsgBox "Registro Alterado", vbExclamation, "Mensagem"
rsUsuario.MoveFirst
atualiza
End Sub

Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18

Conexao
SQLString = "SELECT * FROM tab_usuario"
  Set rsUsuario = New ADODB.Recordset
  rsUsuario.Open SQLString, Con

  If rsUsuario.EOF Or rsUsuario.BOF Then
    MsgBox "Registro não Encontrado!"
  Else
    rsUsuario.MoveFirst
    preenche
  End If


End Sub

Private Sub preenche()

  If rsUsuario.EOF Then
    rsUsuario.MoveFirst
  Else
    If rsUsuario.BOF Then
      rsUsuario.MoveLast
    End If
  End If

  codtxt.Text = rsUsuario!cod_usuario & ""
  nometxt.Text = rsUsuario!nome & ""
  senhatxt.Text = rsUsuario!senha & ""
  
  If rsUsuario!cod_Perfil = 1 Then
    Option1.Value = True
  Else
    Option2.Value = True
  End If
  
  If rsUsuario!ativo = "S" Then
    Check1.Value = 1
  Else
    Check1.Value = 0
  End If
  
  If rsUsuario.EOF Then
    b_last.Enabled = False
    b_next.Enabled = False
  Else
    b_last.Enabled = True
    b_next.Enabled = True
  End If
  
  If rsUsuario.BOF Then
    b_first.Enabled = False
    b_prior.Enabled = False
  Else
    b_first.Enabled = True
    b_prior.Enabled = True
  End If


End Sub

Private Sub atualiza()
  rsUsuario.Requery
  rsUsuario.MoveFirst
  While Trim(rsUsuario!cod_usuario) <> Trim(codtxt.Text)
    rsUsuario.MoveNext
  Wend
  preenche
End Sub


Private Sub picX_Click()
  Unload Me
End Sub

Private Sub CarregaCod()
  fecharRS
  SQLString = "select MAX(cod_usuario+1) as COD from tab_usuario "
  rs.Open SQLString, Con
  codtxt.Text = rs!COD
End Sub
