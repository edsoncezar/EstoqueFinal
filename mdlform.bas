Attribute VB_Name = "Module1"
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

Public Sub Retangulo(m_hWnd As Long, Fator As Byte)
  Dim RGN As Long
  Dim RC As Rect
  Call GetClientRect(m_hWnd, RC)
  RGN = CreateRoundRectRgn(RC.Left, RC.Top, RC.Right, _
                           RC.Bottom, Fator, Fator)
  SetWindowRgn m_hWnd, RGN, True
End Sub


