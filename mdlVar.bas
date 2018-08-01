Attribute VB_Name = "mdlVar"
Option Explicit
'Variaveis conexão ADO
Public rs As ADODB.Recordset
Public Con As ADODB.Connection
Public SQLString As String
'Variaveis
Public administrador As String
Public cod_usuario As Integer
'Variaveis arredondamento do form
Public Declare Function CreateRoundRectRgn Lib _
        "gdi32" (ByVal X1 As Long, ByVal Y1 As _
        Long, ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" _
        (ByVal hWnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long
Public Declare Function GetClientRect Lib "user32" _
        (ByVal hWnd As Long, lpRect As Rect) As Long
Public Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'conexão ADO
Public Sub Conexao()

  Set Con = New ADODB.Connection
  Con.CursorLocation = adUseClient
  Set rs = New ADODB.Recordset
  Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\base_estoque.MDB"

End Sub

'verifica se o record set está aberto
Public Function fecharRS()
    If rs.State = adStateOpen Then
        rs.Close
    End If
End Function

'funcão que arredonda o form
Public Sub Retangulo(m_hWnd As Long, Fator As Byte)
  Dim RGN As Long
  Dim RC As Rect
  Call GetClientRect(m_hWnd, RC)
  RGN = CreateRoundRectRgn(RC.Left, RC.Top, RC.Right, _
                           RC.Bottom, Fator, Fator)
  SetWindowRgn m_hWnd, RGN, True
End Sub


Public Function Horizontal(Newform As Form, Colour1 As ColorConstants, Colour2 As ColorConstants)
    
   
    Dim VR, VG, VB As Single
    Dim Color1, Color2 As Long
    Dim r, G, b, R2, G2, B2, x As Integer
    Dim Temp As Long

    Color1 = Colour1
    Color2 = Colour2

    Temp = (Color1 And 255)
    r = Temp And 255
    Temp = Int(Color1 / 256)
    G = Temp And 255
    Temp = Int(Color1 / 65536)
    b = Temp And 255
    Temp = (Color2 And 255)
    R2 = Temp And 255
    Temp = Int(Color2 / 256)
    G2 = Temp And 255
    Temp = Int(Color2 / 65536)
    B2 = Temp And 255

    VR = Abs(r - R2) / Newform.ScaleWidth
    VG = Abs(G - G2) / Newform.ScaleWidth
    VB = Abs(b - B2) / Newform.ScaleWidth

    If R2 < r Then VR = -VR
    If G2 < G Then VG = -VG
    If B2 < b Then VB = -VB

    For x = 0 To Newform.ScaleWidth
        R2 = r + VR * x
        G2 = G + VG * x
        B2 = b + VB * x
        Newform.Line (x, 0)-(x, Newform.ScaleHeight), RGB(R2, G2, B2)
    Next x
End Function
