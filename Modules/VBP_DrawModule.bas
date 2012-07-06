Attribute VB_Name = "Drawing"
'***************************************************************************
'PhotoDemon Drawing Routines
'©2000-2012 Tanner Helland
'Created: 4/3/01
'Last updated: 04/July/12
'Last update: Rewrote DrawPreviewImage so edge pixels weren't being missed
'
'Miscellaneous drawing routines that don't fit elsewhere.  At present, this includes rendering preview images,
' drawing the canvas background of image forms, and a gradient-rendering sub (used primarily on the histogram form).
'
'***************************************************************************

Option Explicit

'Variables for working with the image previews
Public PreviewWidth As Long, PreviewHeight As Long, PreviewX As Long, PreviewY As Long


'Used to draw the main image onto a preview picture box
Public Sub DrawPreviewImage(ByRef DstPicture As PictureBox)
    GetImageData
    'Dim DWidth As Single, DHeight As Single
    Dim DWidth As Long, DHeight As Long
    
    Dim dstWidth As Single, dstHeight As Single
    dstWidth = DstPicture.ScaleWidth
    dstHeight = DstPicture.ScaleHeight
    
    Dim srcAspect As Single, dstAspect As Single
    srcAspect = PicWidthL / PicHeightL
    dstAspect = dstWidth / dstHeight
    
    If srcAspect > dstAspect Then
        DWidth = DstPicture.ScaleWidth
        DHeight = CSng(PicHeightL / PicWidthL) * DWidth + 0.5
        PreviewY = CInt((DstPicture.ScaleHeight - DHeight) / 2)
        PreviewX = 0
        SetStretchBltMode DstPicture.hdc, STRETCHBLT_HALFTONE
        StretchBlt DstPicture.hdc, 0, PreviewY, DWidth, DHeight, FormMain.ActiveForm.BackBuffer.hdc, 0, 0, PicWidthL, PicHeightL, vbSrcCopy
    Else
        DHeight = DstPicture.ScaleHeight
        DWidth = CSng(PicWidthL / PicHeightL) * DHeight + 0.5
        PreviewX = CInt((DstPicture.ScaleWidth - DWidth) / 2)
        PreviewY = 0
        SetStretchBltMode DstPicture.hdc, STRETCHBLT_HALFTONE
        StretchBlt DstPicture.hdc, PreviewX, 0, DWidth, DHeight, FormMain.ActiveForm.BackBuffer.hdc, 0, 0, PicWidthL, PicHeightL, vbSrcCopy
    End If
    
    PreviewWidth = DWidth - 1
    PreviewHeight = DHeight - 1
    
    DstPicture.Picture = DstPicture.Image
    DstPicture.Refresh
End Sub

'A simple routine to draw the canvas background; the public CanvasBackground variable is used to determine
' draw mode: -1 is a checkerboard effect, any other value is treated as an RGB long
Public Sub DrawSpecificCanvas(ByRef dstForm As Form)

    '-1 indicates the user wants a checkboard background pattern
    If CanvasBackground = -1 Then

        Dim stepIntervalX As Long, stepIntervalY As Long
        stepIntervalX = dstForm.PicCH.ScaleWidth
        stepIntervalY = dstForm.PicCH.ScaleHeight

        If dstForm.ScaleMode = 3 Then
            For x = 0 To dstForm.FrontBuffer.ScaleWidth Step stepIntervalX
            For y = 0 To dstForm.FrontBuffer.ScaleHeight Step stepIntervalY
                BitBlt dstForm.FrontBuffer.hdc, x, y, stepIntervalX, stepIntervalY, dstForm.PicCH.hdc, 0, 0, vbSrcCopy
            Next y
            Next x
            dstForm.FrontBuffer.Picture = dstForm.FrontBuffer.Image
        End If
        
    'Any other value is treated as an RGB long
    Else
    
        dstForm.FrontBuffer.Picture = LoadPicture("")
        dstForm.FrontBuffer.BackColor = CanvasBackground
    
    End If
    
End Sub

'Draw a gradient from Color1 to Color 2 (RGB longs) on a specified picture box
Public Sub DrawGradient(ByVal DstPicBox As Object, ByVal Color1 As Long, ByVal Color2 As Long)

    'Calculation variables (used to interpolate between the gradient colors)
    Dim VR As Single, VG As Single, VB As Single
    
    'Red, green, and blue variables for each gradient color
    Dim r As Long, g As Long, b As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    
    'Extract the red, green, and blue values from the gradient colors (which were passed as Longs)
    r = ExtractR(Color1)
    g = ExtractG(Color1)
    b = ExtractB(Color1)
    r2 = ExtractR(Color2)
    g2 = ExtractG(Color2)
    b2 = ExtractB(Color2)
    
    'Width and height variables are faster than repeated access of .ScaleWidth/Height properties
    Dim tmpHeight As Long
    Dim tmpWidth As Long
    tmpWidth = DstPicBox.ScaleWidth
    tmpHeight = DstPicBox.ScaleHeight

    'Create a calculation variable, which will be used to determine the interpolation step between
    ' each gradient color
    VR = Abs(r - r2) / tmpHeight
    VG = Abs(g - g2) / tmpHeight
    VB = Abs(b - b2) / tmpHeight
    
    'If the second color is less than the first value, make the step negative
    If r2 < r Then VR = -VR
    If g2 < g Then VG = -VG
    If b2 < b Then VB = -VB
    
    'Run a loop across the picture box, changing the gradient color according to the step calculated earlier
    For y = 0 To tmpHeight
        r2 = r + VR * y
        g2 = g + VG * y
        b2 = b + VB * y
        
        DstPicBox.Line (0, y)-(tmpWidth, y), RGB(r2, g2, b2)
    Next y

End Sub
