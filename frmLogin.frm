VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Senha"
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   210
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Controle_de_Estoque.xpcmdbutton cmdCancelar 
      Height          =   375
      Left            =   1710
      TabIndex        =   5
      Top             =   1380
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      Caption         =   "&Cancelar"
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
   Begin Controle_de_Estoque.xpcmdbutton cmdLogin 
      Height          =   375
      Left            =   270
      TabIndex        =   4
      Top             =   1380
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   661
      Caption         =   "&Login"
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
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1050
      Width           =   1395
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   0
      Top             =   630
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
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
      Left            =   360
      TabIndex        =   6
      Top             =   210
      Width           =   540
   End
   Begin VB.Image picX 
      Height          =   315
      Left            =   2430
      Picture         =   "frmLogin.frx":0000
      Top             =   180
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   2745
      Picture         =   "frmLogin.frx":0A55
      Top             =   1890
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   2610
      Picture         =   "frmLogin.frx":1365
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   2205
      Left            =   2925
      Picture         =   "frmLogin.frx":1E0F
      Top             =   540
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmLogin.frx":2700
      Top             =   1905
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmLogin.frx":2FA6
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image6 
      Height          =   2190
      Left            =   0
      Picture         =   "frmLogin.frx":3A50
      Top             =   570
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   3
      Top             =   1050
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   2
      Top             =   645
      Width           =   765
   End
   Begin VB.Image Image9 
      Height          =   585
      Left            =   300
      Picture         =   "frmLogin.frx":4311
      Top             =   0
      Width           =   3915
   End
   Begin VB.Image Image8 
      Height          =   450
      Left            =   240
      Picture         =   "frmLogin.frx":531E
      Top             =   1830
      Width           =   4095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
  End
End Sub


Private Sub cmdLogin_Click()

Dim NomeUsuario As String
Dim DesUsuario As Byte
Dim erro As Byte
administrador = "N"

SQLString = "SELECT * FROM tab_usuario WHERE Nome = '" & txtUsuario.Text & "'"
fecharRS
rs.Open SQLString, Con

    If rs.EOF Then
        erro = erro + 1
        If erro = 3 Then
            MsgBox "Seu Cadastro foi Desativado e o programa será fechado", vbExclamation, "Desativado"
        End If
    Else
        If rs!senha = txtSenha.Text Then
            If rs!ativo = "N" Then
                MsgBox "Usuário Inativo! Fale com o Administrador", vbInformation, "Usuário Inativo!!!"
                Exit Sub
            Else
                If rs!cod_Perfil = 1 Then
                    admnistrador = "S"
                End If
                cod_usuario = rs!cod_usuario
                fecharRS
                Unload Me
                frmPrincipal.Show 1
            End If
        Else
            erro = erro + 1
            If NomeUsuario = txtUsuario.Text Then
                DesUsuario = DesUsuario + 1
            Else
                DesUsuario = 1
                NomeUsuario = txtUsuario.Text
            End If
            If DesUsuario >= 3 And erro >= 3 Then
                MsgBox "Seu Cadastro foi Desativado e o programa será fechado", vbExclamation, "Desativado"
                SQLString = "UPDATE tabela SET bloqueio = 'S' where nome = '" & txtUsuario.Text & "'"
                fecharRS
                rs.Open , SQLString, Con
                End
            ElseIf erro >= 3 And DesUsuario < 3 Then
                MsgBox "Seu Cadastro foi Desativado e o programa será fechado", vbExclamation, "Desativado"
            End If
            txtSenha.Text = ""
            txtUsuario.Text = ""
        End If
    End If
 
End Sub

Private Sub Form_Load()
Conexao
Me.BackColor = &H808080 'cor do form
'chama a função para arredondar os cantos
'area
Retangulo Me.hWnd, 18
txtSenha.PasswordChar = "*"
End Sub


Private Sub picX_Click()
End
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    cmdLogin_Click
  End If
End Sub
