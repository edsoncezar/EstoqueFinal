VERSION 5.00
Begin VB.Form frmalteracaousuario 
   Caption         =   "Alteração  Usuário"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Limpar"
      Height          =   435
      Left            =   2040
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fechar"
      Height          =   435
      Left            =   3720
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Perfil"
      Height          =   1200
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   1605
      Begin VB.OptionButton Option2 
         Caption         =   "Usuário"
         Height          =   330
         Left            =   210
         TabIndex        =   7
         Top             =   720
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Admnistrador"
         Height          =   330
         Left            =   210
         TabIndex        =   6
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ativo?"
      Height          =   225
      Left            =   2520
      TabIndex        =   4
      Top             =   360
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2145
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código de Perfil"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome "
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha "
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   510
   End
End
Attribute VB_Name = "frmalteracaousuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
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
SQLString = SQLString & " nome = '" & Text2.Text & "',"
SQLString = SQLString & " senha = '" & Text3.Text & "',"
SQLString = SQLString & " cod_perfil = '" & perfil & "',"
SQLString = SQLString & " ativo = '" & ativo & "' "
SQLString = SQLString & " WHERE cod_usuario = " & Text1.Text & " "

fecharRS
rs.Open SQLString, Con

MsgBox "Registro Alterado"


End Sub

Private Sub Command2_Click()
 Unload Me


End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Option1.Value = False
Option2.Value = False
Check1.Value = False
Text1.SetFocus

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Option1.Value = False
Option2.Value = False
Check1.Value = False

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
KeyAscii = 0
End If

End Sub

Private Sub Text1_LostFocus()
SQLString = "SELECT * FROM tab_usuario WHERE cod_usuario = " & Val(Text1.Text)
fecharRS
rs.Open SQLString, Con

If rs.EOF Or rs.BOF Then
  MsgBox "Registro não Encontrado!"
Else
  Text1.Text = rs!cod_usuario
  Text2.Text = rs!nome
  Text3.Text = rs!senha
   
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

End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
KeyAscii = 0
End If
End Sub


Private Sub Text2_LostFocus()
SQLString = "SELECT * FROM tab_usuario WHERE cod_usuario = " & Str(Val(Text1.Text))
fecharRS
rs.Open SQLString, Con

If rs.EOF Or rs.BOF Then
  MsgBox "Registro não Encontrado!"
Else
  Text1.Text = rs!cod_usuario
  Text2.Text = rs!nome
  Text3.Text = rs!senha
   
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
End Sub
