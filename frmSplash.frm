VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3405
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   367
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4440
      Top             =   2760
   End
   Begin VB.PictureBox picpgb2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   960
      ScaleHeight     =   19
      ScaleMode       =   0  'User
      ScaleWidth      =   239
      TabIndex        =   0
      Top             =   2940
      Width           =   3585
   End
   Begin VB.Line Line2 
      DrawMode        =   4  'Mask Not Pen
      X1              =   424
      X2              =   424
      Y1              =   168
      Y2              =   8
   End
   Begin VB.Image imgpgb1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Picture         =   "frmSplash.frx":6622
      Top             =   2400
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim distance As Integer
    
    
Private Sub Form_Load()
    distance = 4
    Horizontal Me, RGB(131, 166, 244), RGB(33, 120, 224)
    picpgb2.PaintPicture imgpgb1, 0, 0, 4, 19, 0, 0, 4, 19
    picpgb2.PaintPicture imgpgb1, 4, 0, picpgb2.Width - 9, 19, 4, 0, 10, 19
    picpgb2.PaintPicture imgpgb1, picpgb2.Width - 5, 0, 5, 19, 14, 0, 5, 19
End Sub

Private Sub Form_Terminate()
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    For i = 1 To 2
        picpgb2.PaintPicture imgpgb1.Picture, distance, 4, 8, 12, 23, 5, 8, 12
        distance = distance + 10
    Next i
    If distance > picpgb2.Width - 5 Then
        Timer1.Enabled = False
        Unload Me
        
        frmLogin.Show
    End If
End Sub

