VERSION 5.00
Begin VB.Form frmInclusaoUsuario 
   BorderStyle     =   0  'None
   Caption         =   "Inclusão de Usuário"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Controle_de_Estoque.xpcmdbutton b_fechar 
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
   Begin Controle_de_Estoque.xpcmdbutton b_limpar 
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
   Begin Controle_de_Estoque.xpcmdbutton b_incluir 
      Height          =   375
      Left            =   390
      TabIndex        =   10
      Top             =   2190
      Width           =   1425
      _ExtentX        =   2514
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
   Begin VB.TextBox codtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1140
      TabIndex        =   8
      Top             =   810
      Width           =   855
   End
   Begin VB.CheckBox chkAtivo 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   810
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
      Left            =   3540
      TabIndex        =   2
      Top             =   810
      Width           =   1665
      Begin VB.OptionButton radAdmin 
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
         Left            =   150
         MaskColor       =   &H00808080&
         TabIndex        =   4
         Top             =   240
         Width           =   1485
      End
      Begin VB.OptionButton radUsu 
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
         MaskColor       =   &H00808080&
         TabIndex        =   3
         Top             =   720
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.TextBox senhatxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1140
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1650
      Width           =   2235
   End
   Begin VB.TextBox nometxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1140
      TabIndex        =   0
      Top             =   1215
      Width           =   2235
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inclusão de Usuários"
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
      Width           =   1965
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmInclusaoUsuario.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   4740
      Picture         =   "frmInclusaoUsuario.frx":0AAA
      Top             =   150
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   5115
      Picture         =   "frmInclusaoUsuario.frx":14FF
      Top             =   2730
      Width           =   285
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmInclusaoUsuario.frx":1E0F
      Top             =   2745
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   4980
      Picture         =   "frmInclusaoUsuario.frx":26B5
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   4245
      Left            =   0
      Picture         =   "frmInclusaoUsuario.frx":315F
      Top             =   180
      Width           =   105
   End
   Begin VB.Image Image6 
      Height          =   4245
      Left            =   5295
      Picture         =   "frmInclusaoUsuario.frx":3AC9
      Top             =   480
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
      Height          =   225
      Left            =   300
      TabIndex        =   9
      Top             =   825
      Width           =   630
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
      Top             =   1695
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome: "
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
      TabIndex        =   6
      Top             =   1260
      Width           =   585
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   150
      Picture         =   "frmInclusaoUsuario.frx":4458
      Top             =   2670
      Width           =   8505
   End
   Begin VB.Image Image8 
      Height          =   585
      Left            =   -840
      Picture         =   "frmInclusaoUsuario.frx":530A
      Top             =   0
      Width           =   8505
   End
End
Attribute VB_Name = "frmInclusaoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim perfil As Integer
Dim ativo As String

Private Sub b_fechar_Click()
  Unload Me
End Sub

Private Sub b_incluir_Click()
'Inclusão
If nometxt.Text <> Empty And senhatxt.Text <> Empty Then
 
  If chkAtivo.Value = 1 Then
    ativo = "S"
  Else
    ativo = "N"
  End If

  If radAdmin.Value = True Then
    perfil = 1
  Else
    perfil = 2
  End If
     
  SQLString = "insert into tab_usuario values(" & codtxt.Text & ",'" & nometxt.Text & "', '" & senhatxt.Text & "'," & perfil & ", '" & ativo & "')"
  fecharRS
  rs.Open SQLString, Con
  MsgBox "Registro Inserido", 0, "OK"
  b_limpar_Click
Else
  MsgBox "Preencha todos os Campos!!", vbInformation
End If

End Sub

Private Sub b_limpar_Click()
  nometxt.Text = ""
  senhatxt.Text = ""
  radAdmin.Value = 0
  radUsu.Value = 1
  chkAtivo.Value = 1
  CarregaCod
  nometxt.SetFocus
End Sub


Private Sub CarregaCod()
  fecharRS
  SQLString = "SELECT MAX(cod_usuario+1) as COD FROM tab_usuario "
  rs.Open SQLString, Con
  
  codtxt.Text = Trim(Str(rs!COD))
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

Private Sub picX_Click()
Unload Me
End Sub
