VERSION 5.00
Begin VB.Form frmSobre 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Sobre o SIE"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Controle_de_Estoque.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   5010
      TabIndex        =   14
      Top             =   2850
      Width           =   1545
      _ExtentX        =   2725
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Erick José da Cruz Roberto"
      Height          =   195
      Index           =   5
      Left            =   2970
      TabIndex        =   13
      Top             =   2370
      Width           =   1935
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "35982"
      Height          =   195
      Left            =   6090
      TabIndex        =   12
      Top             =   2310
      Width           =   450
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nº:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6090
      TabIndex        =   11
      Top             =   750
      Width           =   285
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aluno:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2970
      TabIndex        =   10
      Top             =   810
      Width           =   555
   End
   Begin VB.Line Line1 
      X1              =   2970
      X2              =   6570
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "35887"
      Height          =   195
      Left            =   6090
      TabIndex        =   9
      Top             =   1830
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thiago Brito"
      Height          =   195
      Index           =   4
      Left            =   2970
      TabIndex        =   8
      Top             =   1890
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "35839"
      Height          =   195
      Left            =   6090
      TabIndex        =   7
      Top             =   1590
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "35154"
      Height          =   195
      Left            =   6090
      TabIndex        =   6
      Top             =   1110
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "35967"
      Height          =   195
      Left            =   6090
      TabIndex        =   5
      Top             =   2070
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edson Moreira César"
      Height          =   195
      Index           =   3
      Left            =   2970
      TabIndex        =   4
      Top             =   1650
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leandro Aparecido dos Santos Cardozo"
      Height          =   195
      Index           =   2
      Left            =   2970
      TabIndex        =   3
      Top             =   1170
      Width           =   2820
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "César Thiago Xavier Duarte"
      Height          =   195
      Index           =   1
      Left            =   2970
      TabIndex        =   2
      Top             =   1410
      Width           =   1965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Samuel Henrique Matioli"
      Height          =   195
      Index           =   0
      Left            =   2970
      TabIndex        =   1
      Top             =   2130
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "SIE - Sistema Integrado de Estoque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2970
      TabIndex        =   0
      Top             =   210
      Width           =   3705
   End
   Begin VB.Image Image1 
      DragMode        =   1  'Automatic
      Height          =   3390
      Left            =   -30
      Picture         =   "frmSobre.frx":0000
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   3180
   End
End
Attribute VB_Name = "frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub xpcmdbutton1_Click()
  Unload Me
End Sub
