Attribute VB_Name = "ModuleSplash"
Public Function Horizontal(Newform As Form, Colour1 As ColorConstants, Colour2 As ColorConstants)
    
    'Adapitado by Edilson Souza, GoKu
    'Fórum Web - http://www.forumweb.com.br
    
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
