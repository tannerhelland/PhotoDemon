VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Drop Shadow class
'Copyright ©2011-2013 by Tanner Helland
'Created: 24/October/12
'Last updated: 24/October/12
'Last update: initial build
'
'Because PhotoDemon has frequent reason to render drop shadows, I have finally created my own drop
' shadow renderer.  The goal is to allow rendering of drop shadows of any size and offset onto any
' surface.  At first this will just be used for drawing a drop shadow around images (e.g. onto the
' background canvas), but in the future this could also be used for the full image, selections,
' and possibly even text.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Four images are used for the straight stretches of drop shadow - two for vertical, two for horizontal
Private singleHShadow As pdLayer, singleHShadow2 As pdLayer
Private singleVShadow As pdLayer, singleVShadow2 As pdLayer

'Four more images are used for the corners.
Private nwShadow As pdLayer, neShadow As pdLayer
Private swShadow As pdLayer, seShadow As pdLayer

'Create the images necessary for rendering a drop shadow around a square object.
Public Sub initializeSquareShadow(ByVal shadowRadius As Long, ByVal shadowStrength As Byte, ByVal backColor As Long)

    'Note that the shadow strength is just a value between 255 and 0, which will be subtracted from pixel
    ' color values in order to create a "shadow" illusion.
    
    'Create a look-up table of possible drop shadow blend values
    Dim shadowLookup(0 To 255) As Byte
    Dim i As Long
    Dim tmpCalc As Long
    
    For i = 0 To 255
        tmpCalc = Abs(i - shadowStrength)
        If tmpCalc > 255 Then tmpCalc = 255
        shadowLookup(i) = CByte(tmpCalc)
    Next i

    Set singleHShadow = New pdLayer
    Set singleVShadow = New pdLayer
    
    'Initialize shadows that are (1 x shadowRadius) and (shadowRadius x 1) in size
    singleHShadow.createBlank 1, shadowRadius, 24, backColor
    singleVShadow.createBlank shadowRadius, 1, 24, backColor
    
    'Now, draw a smoothly alpha-blended shadow onto each of those
    
    'HORIZONTAL
    
    'Get a pointer to the DIB data
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, singleHShadow
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Loop through the image, alpha-blending a drop shadow effect as we go
    Dim loopSize As Long
    loopSize = shadowRadius - 1
    For i = 0 To loopSize
        iData(2, i) = BlendColors(iData(2, i), shadowLookup(iData(2, i)), CSng(i / loopSize))
        iData(1, i) = BlendColors(iData(1, i), shadowLookup(iData(1, i)), CSng(i / loopSize))
        iData(0, i) = BlendColors(iData(0, i), shadowLookup(iData(0, i)), CSng(i / loopSize))
    Next i
    
    'Release our DIB pointer
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    
    'VERTICAL
    
    'Get a pointer to the DIB data
    prepSafeArray tmpSA, singleVShadow
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    Dim QuickX As Long
    
    'Loop through the image, alpha-blending a drop shadow effect as we go
    For i = 0 To loopSize
        QuickX = i * 3
        iData(QuickX + 2, 0) = BlendColors(iData(QuickX + 2, 0), shadowLookup(iData(QuickX + 2, 0)), CSng(i / loopSize))
        iData(QuickX + 1, 0) = BlendColors(iData(QuickX + 1, 0), shadowLookup(iData(QuickX + 1, 0)), CSng(i / loopSize))
        iData(QuickX, 0) = BlendColors(iData(QuickX, 0), shadowLookup(iData(QuickX, 0)), CSng(i / loopSize))
    Next i
    
    'Release our DIB pointer
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    
    'We now want to create mirror copies of those two shadows
    Set singleHShadow2 = New pdLayer
    Set singleVShadow2 = New pdLayer
    
    singleHShadow2.createFromExistingLayer singleHShadow
    singleVShadow2.createFromExistingLayer singleVShadow
    
    StretchBlt singleHShadow2.getLayerDC, 0, 0, 1, shadowRadius, singleHShadow.getLayerDC, 0, shadowRadius - 1, 1, -shadowRadius, vbSrcCopy
    StretchBlt singleVShadow2.getLayerDC, 0, 0, shadowRadius, 1, singleVShadow.getLayerDC, shadowRadius - 1, 0, -shadowRadius, 1, vbSrcCopy
    
    'Now, we need to render the first of four corner shadows.  The render will only be done once, and each subsequent corner
    ' will simply be a rotation of this initial shadow.
    Set seShadow = New pdLayer
    seShadow.createBlank shadowRadius, shadowRadius, 24, backColor
    
    'Loop through the image, using a distance formula to determine opacity
    prepSafeArray tmpSA, seShadow
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Loop through the image, alpha-blending a drop shadow effect as we go
    Dim x As Long, y As Long
    Dim distPixels As Double
    
    For x = 0 To loopSize
        QuickX = x * 3
    For y = 0 To loopSize
        distPixels = distanceTwoPoints(0, 0, x, y) / (shadowRadius - 1)
        If distPixels > 1 Then distPixels = 1
        iData(QuickX + 2, y) = BlendColors(shadowLookup(iData(QuickX + 2, y)), iData(QuickX + 2, y), distPixels)
        iData(QuickX + 1, y) = BlendColors(shadowLookup(iData(QuickX + 1, y)), iData(QuickX + 1, y), distPixels)
        iData(QuickX, y) = BlendColors(shadowLookup(iData(QuickX, y)), iData(QuickX, y), distPixels)
    Next y
    Next x
        
    'Release our DIB pointer
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
        
    'Now, corner number 2
    Set swShadow = New pdLayer
    swShadow.createBlank shadowRadius, shadowRadius, 24, backColor
    
    'Loop through the image, using a distance formula to determine opacity
    prepSafeArray tmpSA, swShadow
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Loop through the image, alpha-blending a drop shadow effect as we go
    For x = 0 To loopSize
        QuickX = (loopSize - x) * 3
    For y = 0 To loopSize
        distPixels = distanceTwoPoints(0, 0, x, y) / (shadowRadius - 1)
        If distPixels > 1 Then distPixels = 1
        iData(QuickX + 2, y) = BlendColors(shadowLookup(iData(QuickX + 2, y)), iData(QuickX + 2, y), distPixels)
        iData(QuickX + 1, y) = BlendColors(shadowLookup(iData(QuickX + 1, y)), iData(QuickX + 1, y), distPixels)
        iData(QuickX, y) = BlendColors(shadowLookup(iData(QuickX, y)), iData(QuickX, y), distPixels)
    Next y
    Next x
        
    'Release our DIB pointer
    CopyMemory ByVal VarPtrArray(iData), 0&, 4

    'Now, corner number 3
    Set nwShadow = New pdLayer
    nwShadow.createBlank shadowRadius, shadowRadius, 24, backColor
    
    'Loop through the image, using a distance formula to determine opacity
    prepSafeArray tmpSA, nwShadow
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Loop through the image, alpha-blending a drop shadow effect as we go
    For x = 0 To loopSize
        QuickX = (loopSize - x) * 3
    For y = 0 To loopSize
        distPixels = distanceTwoPoints(0, 0, x, y) / (shadowRadius - 1)
        If distPixels > 1 Then distPixels = 1
        iData(QuickX + 2, loopSize - y) = BlendColors(shadowLookup(iData(QuickX + 2, loopSize - y)), iData(QuickX + 2, loopSize - y), distPixels)
        iData(QuickX + 1, loopSize - y) = BlendColors(shadowLookup(iData(QuickX + 1, loopSize - y)), iData(QuickX + 1, loopSize - y), distPixels)
        iData(QuickX, loopSize - y) = BlendColors(shadowLookup(iData(QuickX, loopSize - y)), iData(QuickX, loopSize - y), distPixels)
    Next y
    Next x
        
    'Release our DIB pointer
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    
    'And finally, corner number 4
    Set neShadow = New pdLayer
    neShadow.createBlank shadowRadius, shadowRadius, 24, backColor
    
    'Loop through the image, using a distance formula to determine opacity
    prepSafeArray tmpSA, neShadow
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Loop through the image, alpha-blending a drop shadow effect as we go
    For x = 0 To loopSize
        QuickX = x * 3
    For y = 0 To loopSize
        distPixels = distanceTwoPoints(0, 0, x, y) / (shadowRadius - 1)
        If distPixels > 1 Then distPixels = 1
        iData(QuickX + 2, loopSize - y) = BlendColors(shadowLookup(iData(QuickX + 2, loopSize - y)), iData(QuickX + 2, loopSize - y), distPixels)
        iData(QuickX + 1, loopSize - y) = BlendColors(shadowLookup(iData(QuickX + 1, loopSize - y)), iData(QuickX + 1, loopSize - y), distPixels)
        iData(QuickX, loopSize - y) = BlendColors(shadowLookup(iData(QuickX, loopSize - y)), iData(QuickX, loopSize - y), distPixels)
    Next y
    Next x
        
    'Release our DIB pointer
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    
End Sub

'Return a DC to one of the class's internal shadow images.  Possible values include:
' 0 - Top edge
' 1 - Bottom edge
' 2 - Left edge
' 3 - Right edge
' 4 - Top-left corner
' 5 - Top-right corner
' 6 - Bottom-left corner
' 7 - Bottom-right corner
Public Function getShadowDC(ByVal dcToReturn As Long) As Long

    Select Case dcToReturn
    
        Case 0
            getShadowDC = singleHShadow.getLayerDC
        Case 1
            getShadowDC = singleHShadow2.getLayerDC
        Case 2
            getShadowDC = singleVShadow.getLayerDC
        Case 3
            getShadowDC = singleVShadow2.getLayerDC
        Case 4
            getShadowDC = nwShadow.getLayerDC
        Case 5
            getShadowDC = neShadow.getLayerDC
        Case 6
            getShadowDC = swShadow.getLayerDC
        Case 7
            getShadowDC = seShadow.getLayerDC
    End Select

End Function
