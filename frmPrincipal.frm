VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Sistema de Controle de Estoque "
   ClientHeight    =   5880
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   8025
      Left            =   0
      Picture         =   "frmPrincipal.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12000
   End
   Begin VB.Menu Arquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu Sair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu Cadastros 
      Caption         =   "&Cadastros"
      Begin VB.Menu Usuarios 
         Caption         =   "Usuários"
         Begin VB.Menu Usuarios_Alt 
            Caption         =   "Alteração"
         End
         Begin VB.Menu Usuarios_Excl 
            Caption         =   "Exclusão"
         End
         Begin VB.Menu Usuarios_ 
            Caption         =   "Inserção"
         End
         Begin VB.Menu Usuarios_ConsGen 
            Caption         =   "Consulta Genérica"
         End
         Begin VB.Menu Usuarios_Gen 
            Caption         =   "Acesso Genérico"
         End
      End
      Begin VB.Menu Itens 
         Caption         =   "Itens"
         Begin VB.Menu Itens_Alt 
            Caption         =   "Alteração"
         End
         Begin VB.Menu Itens_Excl 
            Caption         =   "Exclusão"
         End
         Begin VB.Menu Itens_Ins 
            Caption         =   "Inserção"
         End
         Begin VB.Menu Itens_ConsGen 
            Caption         =   "Consulta Genérica"
         End
         Begin VB.Menu Itens_Gen 
            Caption         =   "Acesso Genérico"
         End
      End
      Begin VB.Menu Estoque 
         Caption         =   "Estoque"
         Begin VB.Menu Estoque_Alt 
            Caption         =   "Alteração"
         End
         Begin VB.Menu Estoque_Excl 
            Caption         =   "Exclusão"
         End
         Begin VB.Menu Estoque_Ins 
            Caption         =   "Inserção"
         End
         Begin VB.Menu Movimentos_ConsGen 
            Caption         =   "Consulta Genérica"
         End
         Begin VB.Menu Estoque_Gen 
            Caption         =   "Acesso Genérico"
         End
      End
   End
   Begin VB.Menu Saldos 
      Caption         =   "Saldos"
      Begin VB.Menu Saldo_Atual 
         Caption         =   "Saldo Atual"
      End
      Begin VB.Menu Saldos_Prod 
         Caption         =   "Movimentos por Produto"
      End
   End
   Begin VB.Menu Help1 
      Caption         =   "Help"
      Begin VB.Menu Help 
         Caption         =   "Help"
      End
      Begin VB.Menu sobre 
         Caption         =   "Sobre o SIE"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Estoque_Alt_Click()
  frmAlteracaoMovto.Show 1
End Sub

Private Sub Estoque_Excl_Click()
  frmExclusaoMovto.Show 1
End Sub

Private Sub Estoque_Gen_Click()
   frmMovtoGenerico.Show 1
End Sub

Private Sub Estoque_Ins_Click()
  frmInclusaoMovto.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Con.Close
End Sub

Private Sub Help_Click()
  frmHelp.Show 1
End Sub

Private Sub Itens_Alt_Click()
  frmAlteracaoItem.Show 1
End Sub

Private Sub Itens_ConsGen_Click()
  frmCGItem.Show 1
End Sub

Private Sub Itens_Excl_Click()
  frmexclusaoitem.Show 1
End Sub

Private Sub Itens_Gen_Click()
  frmItemGenerico.Show 1
End Sub

Private Sub Itens_Ins_Click()
  frmInclusaoItem.Show 1
End Sub

Private Sub Movimentos_ConsGen_Click()
  frmCGMOvto.Show 1
End Sub

Private Sub Sair_Click()
  End
End Sub

Private Sub Saldo_Atual_Click()
  frmSaldoAtual.Show 1
End Sub

Private Sub Saldos_Prod_Click()
  frmsaldos.Show 1
End Sub

Private Sub sobre_Click()
  frmSobre.Show 1
End Sub

Private Sub Usuarios__Click()
  frmInclusaoUsuario.Show 1
End Sub

Private Sub Usuarios_Alt_Click()
  frmAlteracaoUsuario.Show 1
End Sub

Private Sub Usuarios_ConsGen_Click()
  frmCGusuario.Show 1
End Sub

Private Sub Usuarios_Excl_Click()
  frmExclusaoUsuario.Show 1
End Sub

Private Sub Usuarios_Gen_Click()
  frmusuariogenerico.Show 1
End Sub
