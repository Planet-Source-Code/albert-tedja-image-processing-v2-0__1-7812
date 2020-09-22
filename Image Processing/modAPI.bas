Attribute VB_Name = "modAPI"
'WINDOWS API DECLARATIONS
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'DECLARES PUBLIC VARIABLES
Public indeks As Integer
Public picforms(0 To 99) As New frmPicture
Public fpath As String
Public hBMPSour(0 To 99) As Long
Public hDCSour(0 To 99) As Long
Public hBMPDest(0 To 99) As Long
Public hDCDest(0 To 99) As Long
Public iCancel As Boolean
Public currDir As String

Function GetRed(cValue As Long) As Long 'A function that is used to get RED value
    GetRed = cValue Mod 256
End Function

Function GetGreen(cValue As Long) As Long   'A function that is used to get GREEN value
    GetGreen = Int((cValue / 256)) Mod 256
End Function

Function GetBlue(cValue As Long) As Long    'A function that is used to get BLUE value
    GetBlue = Int(cValue / 65536)
End Function

Sub Lighten(pfIndex As Integer) 'Codes to make your picture lighter
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    For x = 0 To pX
        For y = 0 To pY
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            red2 = red + 20
            green2 = green + 20
            blue2 = blue + 20
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
        'Copy picture into picture box
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
        'Refresh picture in memory
End Sub

Sub Darken(pfIndex As Integer)  'Codes to make your picture darker
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    For x = 0 To pX
        For y = 0 To pY
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            red2 = red - 20
            green2 = green - 20
            blue2 = blue - 20
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
        'Copy picture into picture box
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
        'Refresh picture in memory
End Sub

Sub Grayscaling(pfIndex As Integer) 'Codes to make your picture in grayscale
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    For x = 0 To pX
        For y = 0 To pY
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            red2 = Int((red + green + blue) / 3)
            green2 = red2
            blue2 = red2
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
        'Copy picture into picture box
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
        'Refresh picture in memory
End Sub

Sub Inverting(pfIndex As Integer)   'Codes to invert color of your picture
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval As Long
    Dim red As Long, green As Long, blue As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    For x = 0 To pX
        For y = 0 To pY
            colorval = GetPixel(hDCSour(pfIndex), x, y)
            red = GetRed(colorval)
            green = GetGreen(colorval)
            blue = GetBlue(colorval)
            
            red2 = 255 - red
            green2 = 255 - green
            blue2 = 255 - blue
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
        'Copy picture into picture box
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
        'Refresh picture in memory
End Sub

Sub Blurring(pfIndex As Integer)    'Codes to blur your picture
    Dim pX As Long, pY As Long
    Dim x As Long, y As Long
    Dim colorval(8) As Long
    Dim red(8) As Long, green(8) As Long, blue(8) As Long
    Dim red2 As Long, green2 As Long, blue2 As Long
    Dim i As Integer
    
    pX = mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth - 1
    pY = mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight - 1
    For x = 1 To pX
        For y = 1 To pY
            colorval(0) = GetPixel(hDCSour(pfIndex), x - 1, y - 1)
            colorval(1) = GetPixel(hDCSour(pfIndex), x - 1, y)
            colorval(2) = GetPixel(hDCSour(pfIndex), x - 1, y + 1)
            colorval(3) = GetPixel(hDCSour(pfIndex), x, y - 1)
            colorval(4) = GetPixel(hDCSour(pfIndex), x, y)
            colorval(5) = GetPixel(hDCSour(pfIndex), x, y + 1)
            colorval(6) = GetPixel(hDCSour(pfIndex), x + 1, y - 1)
            colorval(7) = GetPixel(hDCSour(pfIndex), x + 1, y)
            colorval(8) = GetPixel(hDCSour(pfIndex), x + 1, y + 1)
                'Get color value in 3x3 pixels box
            For i = 0 To 8
                red(i) = GetRed(colorval(i))
                green(i) = GetGreen(colorval(i))
                blue(i) = GetBlue(colorval(i))
                    'Get red, green, and blue values for each pixel
                red2 = red2 + red(i)
                green2 = green2 + green(i)
                blue2 = blue2 + blue(i)
                    'Make a sum of those red, green, and blue values
            Next i
            
            red2 = Int(red2 / 9)
            green2 = Int(green2 / 9)
            blue2 = Int(blue2 / 9)
                'Average those sums
            
            If red2 >= 255 Then red2 = 255
            If green2 >= 255 Then green2 = 255
            If blue2 >= 255 Then blue2 = 255
            If red2 <= 0 Then red2 = 0
            If green2 <= 0 Then green2 = 0
            If blue2 <= 0 Then blue2 = 0
            
            SetPixel hDCDest(pfIndex), x, y, RGB(red2, green2, blue2)
        Next y
    Next x
    BitBlt mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, hDCDest(pfIndex), 0, 0, vbSrcCopy
        'Copy picture into picture box
    mdiImgProcess.ActiveForm.pcbPicture.Refresh
    BitBlt hDCSour(pfIndex), 0, 0, mdiImgProcess.ActiveForm.pcbPicture.ScaleWidth, mdiImgProcess.ActiveForm.pcbPicture.ScaleHeight, mdiImgProcess.ActiveForm.pcbPicture.hdc, 0, 0, vbSrcCopy
        'Refresh picture in memory
End Sub

