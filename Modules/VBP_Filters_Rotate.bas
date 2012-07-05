Attribute VB_Name = "Filters_Rotate"
'***************************************************************************
'Filter (Rotation) Interface
'©2000-2012 Tanner Helland
'Created: 1/25/03
'Last updated: 4/November/07
'Last update: Conversion to DIB section broke some things I didn't previously notice... o_O
'             Now it's all better.  Phew!
'
'Runs all rotation-style filters.  Includes flip and mirror as well.
'It now uses DIB sections for 90 and 270 rotation, so this is some fast s*. :)
'
'***************************************************************************

Option Explicit

'Flip an image horizontally
Public Sub MenuFlip()
    Message "Flipping image..."
    GetImageData
    PicWidthL = PicWidthL + 1
    PicHeightL = PicHeightL + 1
    FormMain.ActiveForm.BackBuffer.PaintPicture FormMain.ActiveForm.BackBuffer.Picture, 0, 0, PicWidthL, PicHeightL, 0, PicHeightL, PicWidthL, -PicHeightL, vbSrcCopy
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer.Refresh
    
    Message "Finished. "
    
    ScrollViewport FormMain.ActiveForm
End Sub

Public Sub MenuMirror()
    Message "Mirroring image..."
    GetImageData
    PicWidthL = PicWidthL + 1
    PicHeightL = PicHeightL + 1
    FormMain.ActiveForm.BackBuffer.PaintPicture FormMain.ActiveForm.BackBuffer.Picture, 0, 0, PicWidthL, PicHeightL, PicWidthL, 0, -PicWidthL, PicHeightL, vbSrcCopy
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer.Refresh
    
    Message "Finished. "
    
    ScrollViewport FormMain.ActiveForm
End Sub

Public Sub MenuRotate90Clockwise()
    Message "Rotating image clockwise..."
    'ImageData() will store the original image data
    GetImageData True
    'Clear out the current picture box and prepare the 2nd buffer
    FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
    FormMain.ActiveForm.BackBuffer.Cls
    FormMain.ActiveForm.BackBuffer.Width = PicHeightL + 3
    FormMain.ActiveForm.BackBuffer.Height = PicWidthL + 3
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer2.Width = FormMain.ActiveForm.BackBuffer.Width
    FormMain.ActiveForm.BackBuffer2.Height = FormMain.ActiveForm.BackBuffer.Height
    FormMain.ActiveForm.BackBuffer2.Picture = FormMain.ActiveForm.BackBuffer2.Image
    DoEvents
    'ImageData2() will store the new (translated) data
    GetImageData2 True
    SetProgBarMax PicWidthL
    'Perform the translation
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        ImageData2((PicWidthL * 3) - QuickVal + 2, y) = ImageData(y * 3 + 2, x)
        ImageData2((PicWidthL * 3) - QuickVal + 1, y) = ImageData(y * 3 + 1, x)
        ImageData2((PicWidthL * 3) - QuickVal, y) = ImageData(y * 3, x)
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData2 True
    'Transfer the picture from the 2nd buffer to the main buffer
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer2.Picture
    'Save some memory by shrinking the 2nd buffer
    FormMain.ActiveForm.BackBuffer2.Picture = LoadPicture("")
    FormMain.ActiveForm.BackBuffer2.Height = 1
    FormMain.ActiveForm.BackBuffer2.Width = 1
    SetProgBarVal cProgBar.Max
    'Manually verify the values of PicWidthL and PicHeightL
    PicWidthL = FormMain.ActiveForm.BackBuffer.ScaleWidth - 1
    PicHeightL = FormMain.ActiveForm.BackBuffer.ScaleHeight - 1
    DisplaySize PicWidthL + 1, PicHeightL + 1
    
    Message "Finished. "
    
    PrepareViewport FormMain.ActiveForm
    SetProgBarVal 0
End Sub

Public Sub MenuRotate180()
    'Rotating 180 degrees can be accomplished by flipping and then mirroring
    'an image.  So instead of writing up code to do this, I just cheat and combine
    'those two routines into one.
    Message "Rotating image..."
    Process Flip, , , , , , , , , , , False
    Process Mirror, , , , , , , , , , , False
    
    Message "Finished. "
    
End Sub

Public Sub MenuRotate270Clockwise()
    Message "Rotating image counter-clockwise..."
    'ImageData() will store the original image data
    GetImageData True
    'Clear out the current picture box and prepare the 2nd buffer
    FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
    FormMain.ActiveForm.BackBuffer.Cls
    FormMain.ActiveForm.BackBuffer.Width = PicHeightL + 3
    FormMain.ActiveForm.BackBuffer.Height = PicWidthL + 3
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer2.Width = FormMain.ActiveForm.BackBuffer.Width
    FormMain.ActiveForm.BackBuffer2.Height = FormMain.ActiveForm.BackBuffer.Height
    FormMain.ActiveForm.BackBuffer2.Picture = FormMain.ActiveForm.BackBuffer2.Image
    DoEvents
    'ImageData2() will store the new (translated) data
    GetImageData2 True
    SetProgBarMax PicWidthL
    'Perform the translation
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        ImageData2(QuickVal + 2, y) = ImageData((PicHeightL - y) * 3 + 2, x)
        ImageData2(QuickVal + 1, y) = ImageData((PicHeightL - y) * 3 + 1, x)
        ImageData2(QuickVal, y) = ImageData((PicHeightL - y) * 3, x)
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData2 True
    'Transfer the picture from the 2nd buffer to the main buffer
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer2.Picture
    'Save some memory by shrinking the 2nd buffer
    FormMain.ActiveForm.BackBuffer2.Picture = LoadPicture("")
    FormMain.ActiveForm.BackBuffer2.Height = 1
    FormMain.ActiveForm.BackBuffer2.Width = 1
    SetProgBarVal cProgBar.Max
    'Manually verify the values of PicWidthL and PicHeightL
    PicWidthL = FormMain.ActiveForm.BackBuffer.ScaleWidth - 1
    PicHeightL = FormMain.ActiveForm.BackBuffer.ScaleHeight - 1
    DisplaySize PicWidthL + 1, PicHeightL + 1
    
    Message "Finished. "
    
    PrepareViewport FormMain.ActiveForm
    SetProgBarVal 0
End Sub
