VERSION 5.00
Begin VB.Form frmAlteracaoUsuario 
   BorderStyle     =   0  'None
   Caption         =   "Alteração de Usuário"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Controle_de_Estoque.xpcmdbutton b_Fechar 
      Height          =   375
      Left            =   3540
      TabIndex        =   12
      Top             =   2190
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
   Begin Controle_de_Estoque.xpcmdbutton b_Limpar 
      Height          =   375
      Left            =   1965
      TabIndex        =   11
      Top             =   2190
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
   Begin Controle_de_Estoque.xpcmdbutton b_Salvar 
      Height          =   375
      Left            =   390
      TabIndex        =   10
      Top             =   2190
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
      Left            =   3600
      TabIndex        =   4
      Top             =   780
      Width           =   1665
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
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1245
      End
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.CheckBox Check1 
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
      Height          =   225
      Left            =   2520
      TabIndex        =   3
      Top             =   780
      Width           =   840
   End
   Begin VB.TextBox senhatxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1140
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1620
      Width           =   2115
   End
   Begin VB.TextBox nometxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1140
      TabIndex        =   1
      Top             =   1200
      Width           =   2145
   End
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1140
      TabIndex        =   0
      Top             =   780
      Width           =   750
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alteração de Usuários"
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
      TabIndex        =   13
      Top             =   210
      Width           =   2085
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   4740
      Picture         =   "frmAlteracaoUsuario.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   4980
      Picture         =   "frmAlteracaoUsuario.frx":0A55
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   5115
      Picture         =   "frmAlteracaoUsuario.frx":14FF
      Top             =   2730
      Width           =   285
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmAlteracaoUsuario.frx":1E0F
      Top             =   2745
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmAlteracaoUsuario.frx":26B5
      Top             =   150
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmAlteracaoUsuario.frx":301F
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   5295
      Picture         =   "frmAlteracaoUsuario.frx":3AC9
      Top             =   450
      Width           =   105
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
      Height          =   300
      Left            =   300
      TabIndex        =   9
      Top             =   780
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
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
      TabIndex        =   8
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
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
      TabIndex        =   7
      Top             =   1620
      Width           =   585
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmAlteracaoUsuario.frx":4458
      Top             =   0
      Width           =   8505
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   120
      Picture         =   "frmAlteracaoUsuario.frx":5BCD
      Top             =   2670
      Width           =   8505
   End
End
Attribute VB_Name = "frmAlteracaoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
  MsgBox "Registro Salvo", vbInformation
End Sub

Private Sub b_fechar_Click()
  Unload Me
End Sub

Private Sub b_limpar_Click()
  codtxt.Text = ""
  nometxt.Text = ""
  senhatxt.Text = ""
  Option1.Value = False
  Option2.Value = False
  Check1.Value = False
  codtxt.SetFocus
End Sub

Private Sub Form_Activate()
  b_limpar_Click
End Sub


Private Sub codtxt_LostFocus()
If codtxt.Text <> Empty Then
  SQLString = "SELECT * FROM tab_usuario WHERE cod_usuario = " & Val(codtxt.Text)
  fecharRS
  rs.Open SQLString, Con

  If rs.EOF Or rs.BOF Then
    MsgBox "Registro não Encontrado!"
    b_Salvar.Enabled = False
    b_limpar_Click
  Else
    b_Salvar.Enabled = True
    codtxt.Text = rs!cod_usuario
    nometxt.Text = rs!nome
    senhatxt.Text = rs!senha
   
    If rs!cod_Perfil = 1 Then
      Option1.Value = True
    Else
      Option2.Value = True
    End If
    
    If rs!ativo = "S" Then
      Check1.Value = 1
    Else
      Check1.Value = 0
    End If
  End If
End If

End Sub

Private Sub Form_Load()
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18
End Sub

Private Sub nometxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  senhatxt.SetFocus
  KeyAscii = 0
End If
End Sub


Private Sub picX_Click()
Unload Me
End Sub
