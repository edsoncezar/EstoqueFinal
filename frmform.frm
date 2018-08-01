VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image8 
      Height          =   450
      Left            =   240
      Picture         =   "frmform.frx":0000
      Top             =   2670
      Width           =   4095
   End
   Begin VB.Image Image7 
      Height          =   315
      Left            =   3990
      Picture         =   "frmform.frx":0BA8
      Top             =   180
      Width           =   315
   End
   Begin VB.Image Image6 
      Height          =   2190
      Left            =   0
      Picture         =   "frmform.frx":15FD
      Top             =   570
      Width           =   105
   End
   Begin VB.Image Image5 
      Height          =   2205
      Left            =   4478
      Picture         =   "frmform.frx":1EBE
      Top             =   540
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   570
      Left            =   0
      Picture         =   "frmform.frx":27AF
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   4170
      Picture         =   "frmform.frx":3259
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmform.frx":3D03
      Top             =   2738
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   4305
      Picture         =   "frmform.frx":45A9
      Top             =   2730
      Width           =   285
   End
   Begin VB.Image Image9 
      Height          =   585
      Left            =   300
      Picture         =   "frmform.frx":4EB9
      Top             =   0
      Width           =   3915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Me.BackColor = &H808080 'Apenas para destacar a cor
  'Coloca o formulário com os cantos arredondados
  'e fator 80 de área
  Retangulo Me.hWnd, 18
End Sub
