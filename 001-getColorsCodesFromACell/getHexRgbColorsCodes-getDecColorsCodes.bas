Attribute VB_Name = "Module1"
Function getHexRgbColorsCodes(FCell As Range) As String
    
    'Code HEX
    Dim hexColor As String
    hexColor = CStr(FCell.Interior.Color)
    hexColor = Right("000000" & Hex(hexColor), 6)
    
    'Code RGB
    Dim rgbColor As Long
    Dim R As Long, G As Long, B As Long
    rgbColor = FCell.Interior.Color
    R = rgbColor Mod 256
    G = (rgbColor \ 256) Mod 256
    B = (rgbColor \ 65536) Mod 256
    
    ' _ (Espace Underscore) : pour continuer le code VBA sur une nouvelle ligne
    'Chr(10): pour revenir à la ligne dans une cellule
    
    getHexRgbColorsCodes = "HEX " & Right(hexColor, 2) & Mid(hexColor, 3, 2) & Left(hexColor, 2) _
                           & Chr(10) _
                           & "RGB " & R & " " & G & " " & B _
                           & Chr(10) _

End Function
Function getDecColorsCodes(FCell As Range, Optional Opt As Integer = 0) As String

    'Code DEC
    Dim decColor As Long
    Dim R As Long, G As Long, B As Long
    decColor = FCell.Interior.Color
    R = decColor Mod 256
    G = (decColor \ 256) Mod 256
    B = (decColor \ 65536) Mod 256
    Select Case Opt
        Case 1
            getDecColorsCodes = R
        Case 2
            getDecColorsCodes = G
        Case 3
            getDecColorsCodes = B
        Case Else
            getDecColorsCodes = "DEC " & decColor
    End Select
End Function

