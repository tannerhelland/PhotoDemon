Attribute VB_Name = "Filters_Area"
'***************************************************************************
'Filter (Area) Interface
'©2000-2012 Tanner Helland
'Created: 12/June/01
'Last updated: 20/Apr/12
'Last update: copied over optimizations and comments from my standalone Custom Filter project (on tannerhelland.com)
'
'Holder for generalized area filters.  Also contains the DoFilter routine, which is central to running
' custom filters (as well as some of the intrinsic PhotoDemon ones).
'
'***************************************************************************

Option Explicit

'The omnipotent DoFilter routine - it takes whatever is in FM() - the "filter matrix" and applies it to the image
Public Sub DoFilter(ByRef FilterType As String, Optional ByVal InvertResult As Boolean = False)
    
    GetImageData
    
    'Note that the only purpose of the FilterType string is to display this message
    Message "Applying " & FilterType & " filter..."
    
    'C and D are like X and Y - they are additional loop variables used for sub-loops
    Dim c As Long, d As Long
    
    'CalcVar determines the size of each sub-loop (so that we don't waste time running a 5x5 matrix on 3x3 filters)
    Dim CalcVar As Long
    CalcVar = (FilterSize \ 2)
    
    'Temporary red, green, and blue values
    Dim tR As Long, tG As Long, tB As Long
    
    'iFM() will hold the contents of FM() - the filter matrix; I don't use FM in case other events want to access it
    Dim iFM() As Long
    
    'Resize iFM according to the size of the filter matrix, then copy over the contents of FM()
    If FilterSize = 3 Then ReDim iFM(-1 To 1, -1 To 1) As Long Else ReDim iFM(-2 To 2, -2 To 2) As Long
    iFM = FM
    
    'FilterWeightA and FilterBiasA are copies of the global FilterWeight and FilterBias variables; again, we don't use the originals in case other events
    ' want to access them
    Dim FilterWeightA As Long, FilterBiasA As Long
    FilterWeightA = FilterWeight
    FilterBiasA = FilterBias
    
    'FilterWeightTemp will be reset for every pixel, and decremented appropriately when attempting to calculate the value for pixels
    ' outside the image perimeter
    Dim FilterWeightTemp As Long
    
    'Temporary calculation variables
    Dim CalcX As Long, CalcY As Long
    
    'tData holds the processed image data; at the end of the filter processing it will get copied over the original image data
    ReDim tData(0 To (PicWidthL * 3) + 3, 0 To PicHeightL + 1)
    
    'TempRef is like QuickX below, but for sub-loops
    Dim TempRef As Long
    
    SetProgBarMax PicWidthL
    
    Dim QuickVal As Long
    
    'Now that we're ready, loop through the image, calculating pixel values as we go
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        
        'Reset our values upon beginning analysis on a new pixel
        tR = 0
        tG = 0
        tB = 0
        FilterWeightTemp = FilterWeightA
        
        'Run a sub-loop around the current pixel
        For c = x - CalcVar To x + CalcVar
            TempRef = c * 3
        For d = y - CalcVar To y + CalcVar
        
            CalcX = c - x
            CalcY = d - y
            
            'If no filter value is being applied to this pixel, ignore it (GoTo's aren't generally a part of good programming, but they ARE convenient :)
            If iFM(CalcX, CalcY) = 0 Then GoTo NextCustomFilterPixel
            
            'If this pixel lies outside the image perimeter, ignore it and adjust FilterWeight accordingly
            If c < 0 Or d < 0 Or c > PicWidthL Or d > PicHeightL Then
                FilterWeightTemp = FilterWeightTemp - iFM(CalcX, CalcY)
                GoTo NextCustomFilterPixel
            End If
            
            'Adjust red, green, and blue according to the values in the filter matrix (FM)
            tR = tR + (ImageData(TempRef + 2, d) * iFM(CalcX, CalcY))
            tG = tG + (ImageData(TempRef + 1, d) * iFM(CalcX, CalcY))
            tB = tB + (ImageData(TempRef, d) * iFM(CalcX, CalcY))

NextCustomFilterPixel:  Next d
        Next c
        
        'If a weight has been set, apply it now
        If (FilterWeightA <> 1) And (FilterWeightTemp <> 0) Then
            tR = tR \ FilterWeightTemp
            tG = tG \ FilterWeightTemp
            tB = tB \ FilterWeightTemp
        End If
        
        'If a bias has been specified, apply it now
        If FilterBiasA <> 0 Then
            tR = tR + FilterBiasA
            tG = tG + FilterBiasA
            tB = tB + FilterBiasA
        End If
        
        'Make sure all values are between 0 and 255
        ByteMeL tR
        ByteMeL tG
        ByteMeL tB
        
        'If inversion is specified, apply it now
        If InvertResult = True Then
            tR = 255 - tR
            tG = 255 - tG
            tB = 255 - tB
        End If
        
        'Finally, remember the new value in our tData array
        tData(QuickVal + 2, y) = tR
        tData(QuickVal + 1, y) = tG
        tData(QuickVal, y) = tB
        
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    
    SetProgBarVal cProgBar.Max
    
    'Copy tData over the original pixel data
    TransferImageData
    
    'Draw the updated image to the screen
    SetImageData
    
End Sub

Public Sub FilterSoften()
    Message "Softening image..."
    Dim tR As Integer, tG As Integer, tB As Integer
    Dim Divisor As Integer
    SetProgBarMax PicWidthL
    Dim QuickVal As Long, QuickVal2 As Long
    Dim c As Long, d As Long
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = 0
        tG = 0
        tB = 0
        For c = x - 1 To x + 1
            If c < 0 Or c > PicWidthL Then GoTo 400
            QuickVal2 = c * 3
        For d = y - 1 To y + 1
            If d < 0 Or d > PicHeightL Then GoTo 300
            If c = x And d = y Then GoTo 300
            Divisor = Divisor + 1
            tR = tR + ImageData(QuickVal2 + 2, d)
            tG = tG + ImageData(QuickVal2 + 1, d)
            tB = tB + ImageData(QuickVal2, d)
300     Next d
400     Next c
        tR = tR + ImageData(QuickVal + 2, y) * Divisor
        tG = tG + ImageData(QuickVal + 1, y) * Divisor
        tB = tB + ImageData(QuickVal, y) * Divisor
        Divisor = Divisor * 2
        tR = tR \ Divisor
        tG = tG \ Divisor
        tB = tB \ Divisor
        Divisor = 0
        ByteMe tR
        ByteMe tG
        ByteMe tB
        tData(QuickVal + 2, y) = tR
        tData(QuickVal + 1, y) = tG
        tData(QuickVal, y) = tB
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    SetProgBarVal cProgBar.Max
    TransferImageData
    SetImageData
End Sub

Public Sub FilterSoftenMore()
    Message "Softening image 2x..."
    Dim tR As Integer, tG As Integer, tB As Integer
    Dim Divisor As Integer
    Dim c As Long, d As Long
    SetProgBarMax PicWidthL
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
    Dim QuickVal As Long, QuickVal2 As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = 0
        tG = 0
        tB = 0
        For c = x - 2 To x + 2
            If c < 0 Or c > PicWidthL Then GoTo 401
            QuickVal2 = x * 3
        For d = y - 2 To y + 2
            If d < 0 Or d > PicHeightL Then GoTo 301
            If c = x And d = y Then GoTo 301
            Divisor = Divisor + 1
            tR = tR + ImageData(QuickVal2 + 2, d)
            tG = tG + ImageData(QuickVal2 + 1, d)
            tB = tB + ImageData(QuickVal2, d)
301     Next d
401     Next c
        tR = tR + ImageData(QuickVal + 2, y) * Divisor
        tG = tG + ImageData(QuickVal + 1, y) * Divisor
        tB = tB + ImageData(QuickVal, y) * Divisor
        Divisor = Divisor * 2
        tR = tR \ Divisor
        tG = tG \ Divisor
        tB = tB \ Divisor
        Divisor = 0
        ByteMe tR
        ByteMe tG
        ByteMe tB
        tData(QuickVal + 2, y) = tR
        tData(QuickVal + 1, y) = tG
        tData(QuickVal, y) = tB
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    SetProgBarVal cProgBar.Max
    TransferImageData
    SetImageData
End Sub

Public Sub FilterBlur()
    Message "Blurring image..."
    Dim tR As Integer, tG As Integer, tB As Integer
    Dim c As Long, d As Long
    Dim Divisor As Byte
    SetProgBarMax PicWidthL
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
    Dim QuickVal As Long, QuickVal2 As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = 0
        tG = 0
        tB = 0
        For c = x - 1 To x + 1
            If c < 0 Or c > PicWidthL Then GoTo 402
            QuickVal2 = c * 3
        For d = y - 1 To y + 1
            If d < 0 Or d > PicHeightL Then GoTo 302
            Divisor = Divisor + 1
            tR = tR + ImageData(QuickVal2 + 2, d)
            tG = tG + ImageData(QuickVal2 + 1, d)
            tB = tB + ImageData(QuickVal2, d)
302     Next d
402     Next c
        tR = tR \ Divisor
        tG = tG \ Divisor
        tB = tB \ Divisor
        Divisor = 0
        ByteMe tR
        ByteMe tG
        ByteMe tB
        tData(QuickVal + 2, y) = tR
        tData(QuickVal + 1, y) = tG
        tData(QuickVal, y) = tB
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    SetProgBarVal cProgBar.Max
    TransferImageData
    SetImageData
End Sub

Public Sub FilterBlurMore()
    Message "Blurring image 2x..."
    Dim tR As Integer, tG As Integer, tB As Integer
    Dim Divisor As Byte
    Dim c As Long, d As Long
    SetProgBarMax PicWidthL
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
    Dim QuickVal As Long, QuickVal2 As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = 0
        tG = 0
        tB = 0
        For c = x - 2 To x + 2
            If c < 0 Or c > PicWidthL Then GoTo 403
            QuickVal2 = c * 3
        For d = y - 2 To y + 2
            If d < 0 Or d > PicHeightL Then GoTo 303
            Divisor = Divisor + 1
            tR = tR + ImageData(QuickVal2 + 2, d)
            tG = tG + ImageData(QuickVal2 + 1, d)
            tB = tB + ImageData(QuickVal2, d)
303     Next d
403     Next c
        tR = tR \ Divisor
        tG = tG \ Divisor
        tB = tB \ Divisor
        Divisor = 0
        ByteMe tR
        ByteMe tG
        ByteMe tB
        tData(QuickVal + 2, y) = tR
        tData(QuickVal + 1, y) = tG
        tData(QuickVal, y) = tB
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    SetProgBarVal cProgBar.Max
    TransferImageData
    SetImageData
End Sub

Public Sub FilterSharpen()
    Message "Sharpening image..."
    Dim tR As Integer, tG As Integer, tB As Integer
    Dim c As Long, d As Long
    SetProgBarMax PicWidthL
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
    Dim QuickVal As Long, QuickVal2 As Long
    For x = 1 To PicWidthL - 1
        QuickVal = x * 3
    For y = 1 To PicHeightL - 1
        tR = 0
        tG = 0
        tB = 0
        For c = x - 1 To x + 1
            If c < 0 Or c > PicWidthL Then GoTo 405
            QuickVal2 = c * 3
        For d = y - 1 To y + 1
            If d < 0 Or d > PicHeightL Then GoTo 305
            If c = x And d = y Then
                tR = tR + ImageData(QuickVal2 + 2, d) * 15
                tG = tG + ImageData(QuickVal2 + 1, d) * 15
                tB = tB + ImageData(QuickVal2, d) * 15
            Else
                tR = tR - ImageData(QuickVal2 + 2, d)
                tG = tG - ImageData(QuickVal2 + 1, d)
                tB = tB - ImageData(QuickVal2, d)
            End If
305     Next d
405     Next c
        tR = tR \ 7
        tG = tG \ 7
        tB = tB \ 7
        ByteMe tR
        ByteMe tG
        ByteMe tB
        tData(QuickVal + 2, y) = tR
        tData(QuickVal + 1, y) = tG
        tData(QuickVal, y) = tB
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    SetProgBarVal cProgBar.Max
    TransferImageData
    SetImageData
End Sub

Public Sub FilterSharpenMore()
    Message "2x sharpening image..."
    Dim tR As Integer, tG As Integer, tB As Integer
    SetProgBarMax PicWidthL
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
    Dim QuickVal As Long
    For x = 1 To PicWidthL - 1
        QuickVal = x * 3
    For y = 1 To PicHeightL - 1
        tR = 0
        tG = 0
        tB = 0
            'Main pixel
            tR = tR + ImageData(QuickVal + 2, y) * 5
            tG = tG + ImageData(QuickVal + 1, y) * 5
            tB = tB + ImageData(QuickVal, y) * 5
            'Outer pixels
            tR = tR - ImageData((x - 1) * 3 + 2, y)
            tG = tG - ImageData((x - 1) * 3 + 1, y)
            tB = tB - ImageData((x - 1) * 3, y)
            tR = tR - ImageData(QuickVal + 2, y - 1)
            tG = tG - ImageData(QuickVal + 1, y - 1)
            tB = tB - ImageData(QuickVal, y - 1)
            tR = tR - ImageData((x + 1) * 3 + 2, y)
            tG = tG - ImageData((x + 1) * 3 + 1, y)
            tB = tB - ImageData((x + 1) * 3, y)
            tR = tR - ImageData(QuickVal + 2, y + 1)
            tG = tG - ImageData(QuickVal + 1, y + 1)
            tB = tB - ImageData(QuickVal, y + 1)
        ByteMe tR
        ByteMe tG
        ByteMe tB
        tData(QuickVal + 2, y) = tR
        tData(QuickVal + 1, y) = tG
        tData(QuickVal, y) = tB
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    SetProgBarVal cProgBar.Max
    TransferImageData
    SetImageData
End Sub

Public Sub FilterUnsharp()
    Message "Unsharpening image..."
    Dim tR As Integer, tG As Integer, tB As Integer
    Dim iR As Integer, iG As Integer, iB As Integer
    Dim c As Long, d As Long
    Dim Divisor As Byte
    SetProgBarMax PicWidthL
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
    Dim QuickVal As Long, QuickVal2 As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = 0
        tG = 0
        tB = 0
        For c = x - 1 To x + 1
            If c < 0 Or c > PicWidthL Then GoTo 405
            QuickVal2 = c * 3
        For d = y - 1 To y + 1
            If d < 0 Or d > PicHeightL Then GoTo 305
            Divisor = Divisor + 1
            tR = tR + ImageData(QuickVal2 + 2, d)
            tG = tG + ImageData(QuickVal2 + 1, d)
            tB = tB + ImageData(QuickVal2, d)
305     Next d
405     Next c
        iR = ImageData(QuickVal + 2, y) * 2
        iG = ImageData(QuickVal + 1, y) * 2
        iB = ImageData(QuickVal, y) * 2
        tR = tR \ Divisor
        tG = tG \ Divisor
        tB = tB \ Divisor
        Divisor = 0
        tR = iR - tR
        tG = iG - tG
        tB = iB - tB
        ByteMe tR
        ByteMe tG
        ByteMe tB
        tData(QuickVal + 2, y) = tR
        tData(QuickVal + 1, y) = tG
        tData(QuickVal, y) = tB
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    SetProgBarVal cProgBar.Max
    TransferImageData
    SetImageData
End Sub

Public Sub FilterGridBlur()
    GetImageData
    Dim tR As Long, tB As Long, tG As Long
    Dim xCalc As Long
    Dim rax() As Long, gax() As Long, bax() As Long
    Dim ray() As Long, gay() As Long, bay() As Long
    ReDim rax(0 To PicWidthL) As Long, gax(0 To PicWidthL) As Long, bax(0 To PicWidthL) As Long
    ReDim ray(0 To PicHeightL) As Long, gay(0 To PicHeightL), bay(0 To PicHeightL)
    
    Message "Generating grids..."
    
    Dim QuickVal As Long
    
    'Generate the x averaging variables
    For x = 0 To PicWidthL
        tR = 0
        tG = 0
        tB = 0
        QuickVal = x * 3
        For y = 0 To PicHeightL
            tR = tR + ImageData(QuickVal + 2, y)
            tG = tG + ImageData(QuickVal + 1, y)
            tB = tB + ImageData(QuickVal, y)
        Next y
        rax(x) = tR
        gax(x) = tG
        bax(x) = tB
    Next x
    
    'Generate the y averaging variables
    For y = 0 To PicHeightL
        tR = 0
        tG = 0
        tB = 0
        For x = 0 To PicWidthL
            QuickVal = x * 3
            tR = tR + ImageData(QuickVal + 2, y)
            tG = tG + ImageData(QuickVal + 1, y)
            tB = tB + ImageData(QuickVal, y)
        Next x
        ray(y) = tR
        gay(y) = tG
        bay(y) = tB
    Next y

    'Apply grid data to the image
    Message "Grid blurring image..."
    SetProgBarMax PicWidthL
    xCalc = PicWidthL + PicHeightL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = (rax(x) + ray(y)) \ xCalc
        tG = (gax(x) + gay(y)) \ xCalc
        tB = (bax(x) + bay(y)) \ xCalc
        ImageData(QuickVal + 2, y) = tR
        ImageData(QuickVal + 1, y) = tG
        ImageData(QuickVal, y) = tB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'A very, very gentle softening effect
Public Sub FilterAntialias()
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, 0) = 1
    FM(1, 0) = 1
    FM(0, -1) = 1
    FM(0, 1) = 1
    FM(0, 0) = 6
    FilterWeight = 10
    FilterBias = 0
    DoFilter "Antialias"
End Sub

'3x3 Gaussian blur
Public Sub FilterGaussianBlur()
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = 1
    FM(0, -1) = 2
    FM(1, -1) = 1
    FM(-1, 0) = 2
    FM(0, 0) = 4
    FM(1, 0) = 2
    FM(-1, 1) = 1
    FM(0, 1) = 2
    FM(1, 1) = 1
    FilterWeight = 16
    FilterBias = 0
    DoFilter "Gaussian Blur "
End Sub

'5x5 Gaussian blur
Public Sub FilterGaussianBlurMore()
    FilterSize = 5
    ReDim FM(-2 To 2, -2 To 2) As Long
    
    FM(-2, -2) = 1
    FM(-1, -2) = 4
    FM(0, -2) = 7
    FM(1, -2) = 4
    FM(2, -2) = 1
    
    FM(-2, -1) = 4
    FM(-1, -1) = 16
    FM(0, -1) = 26
    FM(1, -1) = 16
    FM(2, -1) = 4
    
    FM(-2, 0) = 7
    FM(-1, 0) = 26
    FM(0, 0) = 41
    FM(1, 0) = 26
    FM(2, 0) = 7
    
    FM(-2, 1) = 4
    FM(-1, 1) = 16
    FM(0, 1) = 26
    FM(1, 1) = 16
    FM(2, 1) = 4
    
    FM(-2, 2) = 1
    FM(-1, 2) = 4
    FM(0, 2) = 7
    FM(1, 2) = 4
    FM(2, 2) = 1
    
    FilterWeight = 273
    FilterBias = 0
    DoFilter "Strong Gaussian Blur "
End Sub

Public Sub FilterIsometric()
    Message "Preparing conversion tables..."
    
    'Get the current image data and prepare all the picture boxes
    GetImageData True
    Dim hWidth As Long
    Dim oWidth As Long, oHeight As Long
    oWidth = PicWidthL
    oHeight = PicHeightL
    hWidth = (PicWidthL \ 2)
    
    PicWidthL = PicHeightL + PicWidthL + 1
    PicHeightL = PicWidthL \ 2
    
    FormMain.ActiveForm.BackBuffer.AutoSize = False
    FormMain.ActiveForm.BackBuffer.Width = PicWidthL + 3
    FormMain.ActiveForm.BackBuffer.Height = PicHeightL + 3
    FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
    FormMain.ActiveForm.BackBuffer2.Width = FormMain.ActiveForm.BackBuffer.Width
    FormMain.ActiveForm.BackBuffer2.Height = FormMain.ActiveForm.BackBuffer.Height
    FormMain.ActiveForm.BackBuffer2.Picture = FormMain.ActiveForm.BackBuffer.Picture
    
    DoEvents
    
    GetImageData2 True
    
    'Display the new size
    DisplaySize PicWidthL + 1, PicHeightL + 1
    
    'Perform the translation
    Message "Generating isometric image..."
    SetProgBarMax PicWidthL
    
    Dim TX As Long, TY As Long, QuickVal As Long, QuickVal2 As Long
    
    For x = 0 To PicWidthL
    For y = 0 To PicHeightL
        
        QuickVal2 = x * 3
        TX = getIsometricX(x, y, hWidth)
        
        QuickVal = TX * 3
        TY = getIsometricY(x, y, hWidth)
        
        If (TX >= 0 And TX <= oWidth And TY >= 0 And TY <= oHeight) Then
            ImageData2(QuickVal2 + 2, y) = ImageData(QuickVal + 2, TY)
            ImageData2(QuickVal2 + 1, y) = ImageData(QuickVal + 1, TY)
            ImageData2(QuickVal2, y) = ImageData(QuickVal, TY)
        Else
            ImageData2(QuickVal2 + 2, y) = 255
            ImageData2(QuickVal2 + 1, y) = 255
            ImageData2(QuickVal2, y) = 255
        End If
    
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetProgBarVal cProgBar.Max
    
    SetImageData2 True
    
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer2.Picture
    FormMain.ActiveForm.BackBuffer2.Picture = LoadPicture("")
    FormMain.ActiveForm.BackBuffer2.Width = 1
    FormMain.ActiveForm.BackBuffer2.Height = 1
    
    SetProgBarVal 0
    
    FitOnScreen
End Sub

'These two functions translate a normal (x,y) coordinate to an isometric plane
Private Function getIsometricX(ByVal xc As Long, ByVal yc As Long, ByVal tWidth As Long) As Long
    getIsometricX = (xc / 2) - yc + tWidth
End Function
Private Function getIsometricY(ByVal xc As Long, ByVal yc As Long, ByVal tWidth As Long) As Long
    getIsometricY = (xc / 2) + yc - tWidth
End Function

'Temporary arrays are necessary for many area transformations - this handles the transfer between the temp array and ImageData()
Public Sub TransferImageData()
    Message "Transferring data..."
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        ImageData(QuickVal + 2, y) = tData(QuickVal + 2, y)
        ImageData(QuickVal + 1, y) = tData(QuickVal + 1, y)
        ImageData(QuickVal, y) = tData(QuickVal, y)
    Next y
    Next x
    
    Erase tData
    
End Sub
