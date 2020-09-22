Attribute VB_Name = "basThumbViewer"
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Sub ViewImage(ByVal strFile As String, picTemp As PictureBox, picTarget As PictureBox, Optional intQuality As Integer = 0)
    Dim x&, y&, x1&, y1&, x2&, y2&, z1!, R&, G&, B&
    Dim TempOffsetX&, TempOffsetY&, TempX&, TempY&, factor!, Lower!, Upper!, Skip&
    Dim rgbColor&, PixelCount&, Count&
    Dim sNoPreview$
    On Error GoTo ErrorHandler
    
    'Set default stuffs
    picTarget.Cls
    picTarget.AutoRedraw = True
    picTemp.Visible = False
    picTemp.AutoSize = True
    
    'get target sizing info
    x = picTarget.Width
    y = picTarget.Height
    
    'Load the image
    If strFile <> "" Then picTemp.Picture = LoadPicture(strFile)
    
    'get source sizing info
    x1 = picTemp.Width
    y1 = picTemp.Height
    
    'Determine conversion ratio to use
    z1 = IIf(x / x1 * y1 < y, x / x1, y / y1)
    
    'Calculate new image size
    x1 = x1 * z1
    y1 = y1 * z1
    
    'Correct invalid quality settings
    intQuality = intQuality Mod 5
    
    If intQuality > 0 Then
        'Calculate pixel range
        factor = 1 / z1
        If factor > 1 Then
            Upper = (factor - 1) / 2
            Lower = -Upper
        Else
            Lower = -factor / 2
            Upper = factor / 2
        End If
        
        'Calculate pixel skip number
        Skip = CInt(factor - (factor * intQuality / 4)) + 1
        
        'Calculate image size in pixels
        x1 = x1 / 15
        y1 = y1 / 15
        
        'Calculate top-left corner
        x2 = ((x / 15) - x1) / 2
        y2 = ((y / 15) - y1) / 2
                
        'Draw Thumbnail
        For y = 1 To y1 - 1
            DoEvents
            For x = 1 To x1 - 1
                R = 0
                G = 0
                B = 0
                PixelCount = 0
                Count = 0
                TempX = x * factor + Lower
                TempY = y * factor + Lower
                'Capture pixel Range
                For TempOffsetX = Lower To Upper
                    For TempOffsetY = Lower To Upper
                        'Check for pixel skip
                        If Count Mod Skip = 0 Then
                            rgbColor = GetPixel(picTemp.hdc, TempX + TempOffsetX, TempY + TempOffsetY)
                            If rgbColor >= 0 Then
                                R = R + rgbColor Mod 256
                                G = G + (rgbColor \ 256) Mod 256
                                B = B + (rgbColor \ 65536) Mod 256
                                PixelCount = PixelCount + 1
                            End If
                        End If
                        Count = Count + 1
                    Next
                Next
                'Draw average color to pixel
                If PixelCount > 0 Then
                    'Faster, but doesn't can't use AutoRedraw :/
                    'SetPixel picTarget.hdc, x2 + x, y2 + y, RGB(R / PixelCount, G / PixelCount, B / PixelCount)
                    TempX = (x2 + x - 1) * 15
                    TempY = (y2 + y - 1) * 15
                    picTarget.PSet (TempX, TempY), RGB(R / PixelCount, G / PixelCount, B / PixelCount)
                End If
            Next
        Next
    Else
        'Draw Thumbnail
        picTarget.PaintPicture picTemp.Picture, (x - x1) / 2, (y - y1) / 2, x1, y1
    End If
    Exit Sub
    
ErrorHandler:
    'set temp image to nothing
    picTemp.Picture = LoadPicture()
    'Display default error message
    sNoPreview = "No Preview Available"
    picTarget.CurrentX = x / 2 - picTarget.TextWidth(sNoPreview) / 2
    picTarget.CurrentY = y / 2 - picTarget.TextHeight(sNoPreview) / 2
    picTarget.Print sNoPreview
End Sub
