VERSION 5.00
Begin VB.Form toolpanel_ColorPicker 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14265
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "Toolpanel_ColorPicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   951
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   1935
      Index           =   0
      Left            =   1440
      Top             =   960
      Visible         =   0   'False
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   3413
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   255
         Index           =   0
         Left            =   120
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         Caption         =   "after clicking"
      End
      Begin PhotoDemon.pdButtonStrip btsSampleMerged 
         Height          =   945
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   1667
         Caption         =   "sample from"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdCheckBox chkAfter 
         Height          =   345
         Left            =   210
         TabIndex        =   4
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         Caption         =   "return to previous tool"
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   3120
         TabIndex        =   6
         Top             =   1395
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdPictureBox picSample 
      Height          =   810
      Left            =   15
      Top             =   15
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1429
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   255
      Index           =   0
      Left            =   6120
      Top             =   60
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   450
      Caption         =   "red"
   End
   Begin PhotoDemon.pdDropDown cboColorSpace 
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdSlider sldRadius 
      Height          =   375
      Left            =   1500
      TabIndex        =   0
      Top             =   390
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   661
      FontSizeCaption =   10
      Max             =   100
      ScaleStyle      =   1
      NotchPosition   =   2
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   255
      Index           =   1
      Left            =   8040
      Top             =   60
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   450
      Caption         =   "green"
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   255
      Index           =   2
      Left            =   9960
      Top             =   60
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   450
      Caption         =   "blue"
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   255
      Index           =   3
      Left            =   11880
      Top             =   60
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   450
      Caption         =   "alpha"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   0
      Left            =   7200
      Top             =   60
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   1
      Left            =   9120
      Top             =   60
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   2
      Left            =   11040
      Top             =   60
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   3
      Left            =   12960
      Top             =   60
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   255
      Index           =   4
      Left            =   6120
      Top             =   495
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   450
      Caption         =   "red"
   End
   Begin PhotoDemon.pdDropDown cboColorSpace 
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   435
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   255
      Index           =   5
      Left            =   8040
      Top             =   495
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   450
      Caption         =   "green"
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   255
      Index           =   6
      Left            =   9960
      Top             =   495
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   450
      Caption         =   "blue"
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   255
      Index           =   7
      Left            =   11880
      Top             =   495
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   450
      Caption         =   "alpha"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   4
      Left            =   7200
      Top             =   495
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   5
      Left            =   9120
      Top             =   495
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   6
      Left            =   11040
      Top             =   495
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   7
      Left            =   12960
      Top             =   495
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      Caption         =   "sample radius"
      Value           =   0   'False
   End
End
Attribute VB_Name = "toolpanel_ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Color-Picker Tool Panel
'Copyright 2013-2026 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 07/November/21
'Last update: migrate to new flyout-driven UI
'
'Color pickers are pretty straightforward tools: sample pixels from the image, and reflect the results on-screen.
' The main purpose of this tool is to "stay out of the damn way", I think!
'
'PD provides a standard assortment of options, and two separate color views (so you can see e.g. RGB and HSV
' values simultaneously).  I may add a third view in the future, as there's plenty of free space on modern displays.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The toolpanel for this dialog makes it easy to see multiple color space values at once.
Private Enum PD_ColorPickerSpaces
    cps_RGBA = 0
    cps_RGBAPercent = 1
    cps_HSV = 2
    cps_CMYK = 3
    cps_ColorSpaceCount = 4
End Enum

#If False Then
    Private Const cps_RGBA = 0, cps_RGBAPercent = 1, cps_HSV = 2, cps_CMYK = 3, cps_ColorSpaceCount = 4
#End If

'Translated text for all color spaces.  These strings are populated when the toolbar is loaded; this greatly
' improves rendering performance when translations are active.
Private m_ColorNames() As String
Private m_NullTextString As String
Private m_StringsInitialized As Boolean

'Last-passed mouse coordinates.  To spare repeat processing when zoomed-in, we cache these and only update
' our color samples if they change.
Private m_ImgX As Single, m_ImgY As Single

'If the current cursor position is OOB, this will be set to TRUE.  (Similarly, if no images are loaded,
' this will also be set to TRUE.)
Private m_NoColorAvailable As Boolean

'The current values of the last-selected color are cached, so the user can toggle color space modes without
' losing color data.  Note that we use RGBA notation here, because the values returned from the canvas are
' already translated into the current RGBA working space.
Private m_Red As Single, m_Green As Single, m_Blue As Single, m_Alpha As Single

'If we need to sample an area of the source image (or if we are sampling merged colors), we'll need a temporary
' DIB to store the results.
Private m_SampleDIB As pdDIB

'Preview DIB of the current color (displayed right there in the toolbox)
Private m_PreviewDIB As pdDIB

'Flyout manager
Private WithEvents m_Flyout As pdFlyout
Attribute m_Flyout.VB_VarHelpID = -1

'The value of all controls on this form are saved and loaded to file by this class
' (Normally this is declared WithEvents, but this dialog doesn't require custom settings behavior.)
Private m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'Mouse interactions will call into this function, supplying the x/y coordinates (in the current image space)
' of the current mouse operation.  This function will then translate those coordinates, using the current
' tool settings, into usable color values.
Public Sub NotifyCanvasXY(ByVal mouseButtonDown As Boolean, ByVal imgX As Single, ByVal imgY As Single, ByRef srcCanvas As pdCanvas)
    
    Dim initColorAvailable As Boolean
    initColorAvailable = m_NoColorAvailable
    
    Dim sampleRadius As Long
    sampleRadius = sldRadius.Value
    
    'First, make sure we have a valid image to check!
    m_NoColorAvailable = Not PDImages.IsImageActive()
    
    'Next, ignore color retrieval if these coordinates match our last ones
    If (imgX = m_ImgX) And (imgY = m_ImgY) And (Not mouseButtonDown) Then Exit Sub
    m_ImgX = imgX
    m_ImgY = imgY
    
    Dim sampleLeft As Long, sampleTop As Long, sampleRight As Long, sampleBottom As Long
    Dim sampleWidth As Long, sampleHeight As Long
    
    'If previous steps determined that a color isn't available at this position, we have no further work to do.
    If (Not m_NoColorAvailable) Then
    
        'Grab a color from the correct source.
        If (btsSampleMerged.ListIndex = 0) And (PDImages.GetActiveImage.GetNumOfLayers > 1) Then
            
            'Before proceeding, ensure the mouse pointer lies within the image.
            If (m_ImgX < 0) Or (m_ImgY < 0) Or (m_ImgX > PDImages.GetActiveImage.Width) Or (m_ImgY > PDImages.GetActiveImage.Height) Then
                m_NoColorAvailable = True
            Else
                
                'We need to retrieve a composited rect from the image's compositor, at the size of the requested
                ' sample radius (if any).
                
                'First, figure out the area to sample
                sampleLeft = Int(imgX) - sampleRadius
                sampleTop = Int(imgY) - sampleRadius
                If (sampleLeft < 0) Then sampleLeft = 0
                If (sampleTop < 0) Then sampleTop = 0
                
                sampleRight = Int(imgX) + sampleRadius
                sampleBottom = Int(imgY) + sampleRadius
                If (sampleRight > PDImages.GetActiveImage.Width) Then sampleRight = PDImages.GetActiveImage.Width
                If (sampleBottom > PDImages.GetActiveImage.Height) Then sampleBottom = PDImages.GetActiveImage.Height
                
                'Cover the special case of "sample radius = 0"
                If (sampleRight < sampleLeft + 1) Then sampleRight = sampleLeft + 1
                If (sampleBottom < sampleTop + 1) Then sampleBottom = sampleTop + 1
                
                sampleWidth = sampleRight - sampleLeft
                sampleHeight = sampleBottom - sampleTop
                        
                'Make a local copy of the pixel data
                If (m_SampleDIB Is Nothing) Then Set m_SampleDIB = New pdDIB
                m_SampleDIB.CreateBlank sampleWidth, sampleHeight, 32, 0, 0
                
                Dim dstRectF As RectF, srcRectF As RectF
                With dstRectF
                    .Left = 0
                    .Top = 0
                    .Width = sampleWidth
                    .Height = sampleHeight
                End With
                
                With srcRectF
                    .Left = sampleLeft
                    .Top = sampleTop
                    .Width = sampleWidth
                    .Height = sampleHeight
                End With
                
                PDImages.GetActiveImage.GetCompositedRect m_SampleDIB, dstRectF, srcRectF, GP_IM_NearestNeighbor, False, CLC_ColorSample
                
                'Find an average!
                FindAverageValues
                
            End If
            
        'Current layer only...
        Else
        
            Dim layerX As Single, layerY As Single
            Drawing.ConvertImageCoordsToLayerCoords_Full PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, imgX, imgY, layerX, layerY
            
            Dim srcRGBA As RGBQuad
            If Layers.GetRGBAPixelFromLayer(PDImages.GetActiveImage.GetActiveLayerIndex, Int(layerX), Int(layerY), srcRGBA) Then
            
                'A valid color was found!  Fill our module-level color values.
                Dim unPremult As Single
                
                'If the current sampling radius is 1, we can use the return as-is
                If (sldRadius.Value = 0) Then
                
                    If (srcRGBA.Alpha > 0#) Then unPremult = (255# / srcRGBA.Alpha) Else unPremult = 0#
                    
                    With srcRGBA
                        m_Red = .Red * unPremult
                        m_Green = .Green * unPremult
                        m_Blue = .Blue * unPremult
                        m_Alpha = .Alpha
                    End With
                
                'If sampling is active, we need to retrieve a larger area from the source layer,
                ' then manually calculate an average color.
                Else
                    
                    'Figure out the area to sample
                    sampleLeft = Int(layerX) - sampleRadius
                    sampleTop = Int(layerY) - sampleRadius
                    If (sampleLeft < 0) Then sampleLeft = 0
                    If (sampleTop < 0) Then sampleTop = 0
                    
                    sampleRight = Int(layerX) + sampleRadius
                    sampleBottom = Int(layerY) + sampleRadius
                    If (sampleRight > PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)) Then sampleRight = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
                    If (sampleBottom > PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)) Then sampleBottom = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
                    
                    sampleWidth = sampleRight - sampleLeft
                    sampleHeight = sampleBottom - sampleTop
                    
                    'Make a local copy of the pixel data
                    If (m_SampleDIB Is Nothing) Then Set m_SampleDIB = New pdDIB
                    m_SampleDIB.CreateBlank sampleWidth, sampleHeight, 32, 0, 0
                    GDI.BitBltWrapper m_SampleDIB.GetDIBDC, 0, 0, sampleWidth, sampleHeight, PDImages.GetActiveImage.GetActiveDIB.GetDIBDC, sampleLeft, sampleTop, vbSrcCopy
                    
                    'Find an average!
                    FindAverageValues
                
                End If
            
            Else
                m_NoColorAvailable = True
            End If
        
        End If
        
    End If
    
    'If the mouse is down, update the current color accordingly.
    If (mouseButtonDown And (Not m_NoColorAvailable)) Then layerpanel_Colors.SetCurrentColor m_Red, m_Green, m_Blue
    
    'Update the display as necessary
    If (Not m_NoColorAvailable) Or (initColorAvailable <> m_NoColorAvailable) Then UpdateUIText
    
End Sub

Public Sub NotifyMouseReleased()
    If chkAfter.Value And (g_PreviousTool <> TOOL_UNDEFINED) And (Not Tools.GetToolAltState()) Then toolbar_Toolbox.SelectNewTool g_PreviousTool
End Sub

'Find the average color value of the pixels in the (already prepared) m_SampleDIB object.
Private Sub FindAverageValues()

    If (m_SampleDIB Is Nothing) Then Exit Sub
    
    Dim x As Long, y As Long, xFinal As Long, yFinal As Long
    xFinal = (m_SampleDIB.GetDIBWidth - 1) * 4
    yFinal = m_SampleDIB.GetDIBHeight - 1
    
    Dim lineOfPixels() As Byte, tmpSA As SafeArray1D
    m_SampleDIB.WrapArrayAroundScanline lineOfPixels, tmpSA, 0
    
    Dim pxPtr As Long, pxWidth As Long
    pxPtr = m_SampleDIB.GetDIBPointer
    pxWidth = m_SampleDIB.GetDIBStride
    
    Dim rTotal As Long, gTotal As Long, bTotal As Long, aTotal As Long
    
    For y = 0 To yFinal
        tmpSA.pvData = pxPtr + y * pxWidth
    For x = 0 To xFinal Step 4
        bTotal = bTotal + lineOfPixels(x)
        gTotal = gTotal + lineOfPixels(x + 1)
        rTotal = rTotal + lineOfPixels(x + 2)
        aTotal = aTotal + lineOfPixels(x + 3)
    Next x
    Next y
    
    m_SampleDIB.UnwrapArrayFromDIB lineOfPixels
    
    Dim pxDivisor As Single
    pxDivisor = 1# / (m_SampleDIB.GetDIBWidth * m_SampleDIB.GetDIBHeight)
    
    m_Blue = CSng(bTotal) * pxDivisor
    m_Green = CSng(gTotal) * pxDivisor
    m_Red = CSng(rTotal) * pxDivisor
    m_Alpha = CSng(aTotal) * pxDivisor
    
    'Finally, un-premultiply the color values
    If (m_Alpha > 0!) Then
        pxDivisor = 255# / m_Alpha
        m_Red = m_Red * pxDivisor
        m_Green = m_Green * pxDivisor
        m_Blue = m_Blue * pxDivisor
    End If

End Sub

Private Sub UpdateUIText()
    
    'If we haven't pulled localized strings from the translation engine yet, bail
    If (Not m_StringsInitialized) Then Exit Sub
    
    Dim i As Long, j As Long, curCategory As Long
    Dim textChanged As Boolean, atLeastOneTextChanged As Boolean
    
    'Regardless of color settings, we always start by filling the color name labels
    For i = cboColorSpace.lBound To cboColorSpace.UBound
        
        curCategory = cboColorSpace(i).ListIndex
        If (curCategory < 0) Then curCategory = 0
        
        For j = 0 To 3
            textChanged = Strings.StringsNotEqual(lblColor(j + i * 4).Caption, m_ColorNames(curCategory, j) & ":")
            If textChanged Then
                lblColor(j + i * 4).Caption = m_ColorNames(curCategory, j) & ":"
                atLeastOneTextChanged = True
            End If
        Next j
        
    Next i
    
    'If captions were changed, reflow the layout horizontally.  (This produces a better UI because
    ' the difference in width between captions like "red" and "saturation" can be huge, and there's
    ' no good one-size-fits-all solution across all localizations.)
    If atLeastOneTextChanged Then
        
        'Use a pdFont object for precise text measurements
        Dim cFont As pdFont
        Set cFont = New pdFont
        cFont.SetFontSize 10    'Size is hard-coded to match "best-case" font size of these labels
        cFont.SetFontBold False
        
        Dim xOffset As Long
        xOffset = cboColorSpace(0).GetLeft + cboColorSpace(0).GetWidth + Interface.FixDPI(8)
        
        Dim padBetweenColors As Long
        padBetweenColors = Interface.FixDPI(10)
        
        'Iterate through controls horizontally
        For j = 0 To 3
            
            'Find the larger of the two columns and set horizontal width of *both* to match.
            Dim maxWidth As Long, testWidth As Long
            maxWidth = cFont.GetWidthOfString(m_ColorNames(IIf(cboColorSpace(0).ListIndex < 0, 0, cboColorSpace(0).ListIndex), j) & ":")
            testWidth = cFont.GetWidthOfString(m_ColorNames(IIf(cboColorSpace(1).ListIndex < 0, 0, cboColorSpace(1).ListIndex), j) & ":")
            If (testWidth > maxWidth) Then maxWidth = testWidth
            
            'Add a tiny bit of padding
            maxWidth = maxWidth + 2
            
            'Reflow the title labels, while also fixing max width
            lblColor(j).SetPositionAndSize xOffset, lblColor(j).GetTop, maxWidth, lblColor(j).GetHeight
            lblColor(4 + j).SetPositionAndSize xOffset, lblColor(4 + j).GetTop, maxWidth, lblColor(4 + j).GetHeight
            
            'Reflow their matching value labels
            xOffset = xOffset + maxWidth + Interface.FixDPI(2)
            lblValue(j).SetLeft xOffset
            lblValue(4 + j).SetLeft xOffset
            testWidth = cFont.GetWidthOfString("100.0%") + 2
            lblValue(j).SetWidth testWidth
            lblValue(4 + j).SetWidth testWidth
            
            'Add padding before moving to the next entry
            xOffset = xOffset + lblValue(j).GetWidth + padBetweenColors
            
        Next j
        
    End If
                    
    'If a color isn't available, blank all labels
    If m_NoColorAvailable Then
        
        For i = cboColorSpace.lBound To cboColorSpace.UBound
            For j = 0 To 3
                lblValue(j + i * 4).Caption = m_NullTextString
            Next j
        Next i
        
    Else
        
        'Iterate through all color space dropdowns, and update their text accordingly
        For i = cboColorSpace.lBound To cboColorSpace.UBound
        
            curCategory = cboColorSpace(i).ListIndex
            If (curCategory < 0) Then curCategory = 0
            
            Dim idxLabel As Long
            idxLabel = i * 4
            
            Select Case curCategory
                
                Case cps_RGBA
                    
                    'Color values are easy in RGB!
                    lblValue(idxLabel).Caption = Int(m_Red)
                    lblValue(idxLabel + 1).Caption = Int(m_Green)
                    lblValue(idxLabel + 2).Caption = Int(m_Blue)
                    lblValue(idxLabel + 3).Caption = Int(m_Alpha)
                    
                Case cps_RGBAPercent
                
                    lblValue(idxLabel).Caption = Format$(m_Red / 255#, "0.0%")
                    lblValue(idxLabel + 1).Caption = Format$(m_Green / 255#, "0.0%")
                    lblValue(idxLabel + 2).Caption = Format$(m_Blue / 255#, "0.0%")
                    lblValue(idxLabel + 3).Caption = Format$(m_Alpha / 255#, "0.0%")
                    
                Case cps_HSV
                
                    Dim cHue As Double, cSat As Double, cVal As Double
                    Colors.fRGBtoHSV m_Red / 255#, m_Green / 255#, m_Blue / 255#, cHue, cSat, cVal
                    
                    lblValue(idxLabel).Caption = Format$((cHue * 360#), "#0.0") & ChrW$(&HB0&)
                    lblValue(idxLabel + 1).Caption = Format$(cSat, "0.0%")
                    lblValue(idxLabel + 2).Caption = Format$(cVal, "0.0%")
                    lblValue(idxLabel + 3).Caption = Format$(m_Alpha / 255#, "0.0%")
                    
                Case cps_CMYK
                    
                    Dim rTmp As Double, gTmp As Double, bTmp As Double
                    rTmp = m_Red / 255#
                    gTmp = m_Green / 255#
                    bTmp = m_Blue / 255#
                    
                    Dim cK As Double, mK As Double, yK As Double, bK As Double
                    bK = 1# - PDMath.Max3Float(rTmp, gTmp, bTmp)
                    
                    If (bK < 1#) Then
                        cK = (1# - rTmp - bK) / (1# - bK)
                        mK = (1# - gTmp - bK) / (1# - bK)
                        yK = (1# - bTmp - bK) / (1# - bK)
                    Else
                        cK = 0#
                        mK = 0#
                        yK = 0#
                    End If
                    
                    lblValue(idxLabel).Caption = Format$(cK, "0.0%")
                    lblValue(idxLabel + 1).Caption = Format$(mK, "0.0%")
                    lblValue(idxLabel + 2).Caption = Format$(yK, "0.0%")
                    lblValue(idxLabel + 3).Caption = Format$(bK, "0.0%")
                    
            End Select
        
        Next i
        
    End If
    
    'Finally, paint the new color preview
    RegenerateColorSampleBox
    
End Sub

Private Sub RegenerateColorSampleBox(Optional ByVal redrawImmediately As Boolean = True)

    Dim sampleWidth As Long, sampleHeight As Long
    sampleWidth = picSample.GetWidth
    sampleHeight = picSample.GetHeight
    
    If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
    If (m_PreviewDIB.GetDIBWidth <> sampleWidth) Or (m_PreviewDIB.GetDIBHeight <> sampleHeight) Then
        m_PreviewDIB.CreateBlank sampleWidth, sampleHeight, 32, 0, 255
    Else
        m_PreviewDIB.ResetDIB 0
    End If
    
    'Checkerboard first (for the opacity region)
    GDI_Plus.GDIPlusFillDIBRect_Pattern m_PreviewDIB, 0!, 0!, sampleWidth, sampleHeight, g_CheckerboardPattern, , True
    
    'All subsequent renders only operate if a valid color has been selected
    If (Not m_NoColorAvailable) Then
        
        'Opaque color next
        Dim tmpSurface As pd2DSurface
        Set tmpSurface = New pd2DSurface
        tmpSurface.WrapSurfaceAroundPDDIB m_PreviewDIB
        
        Dim tmpBrush As pd2DBrush
        Drawing2D.QuickCreateSolidBrush tmpBrush, RGB(m_Red, m_Green, m_Blue), m_Alpha * (100# / 255#)
        PD2D.FillRectangleI tmpSurface, tmpBrush, 0, 0, sampleWidth, sampleHeight
        
        '"Pure" color next
        Drawing2D.QuickCreateSolidBrush tmpBrush, RGB(m_Red, m_Green, m_Blue), 100#
        PD2D.FillRectangleI tmpSurface, tmpBrush, 0, 0, sampleWidth, sampleHeight \ 2
        
        'Finally, draw a neutral-color border around the control
        If (Not g_Themer Is Nothing) Then
            Dim tmpPen As pd2DPen: Set tmpPen = New pd2DPen
            tmpPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
            tmpPen.SetPenLineJoin P2_LJ_Miter
            PD2D.DrawRectangleI tmpSurface, tmpPen, 0, 0, sampleWidth - 1, sampleHeight - 1
        End If
        
    End If
    
    'Free our pd2D objects and flip the buffer to the screen
    Set tmpBrush = Nothing: Set tmpSurface = Nothing
    
    If redrawImmediately Then
        Dim pichDC As Long
        picSample.StartPaint pichDC, sampleWidth, sampleHeight
        GDI.BitBltWrapper pichDC, 0, 0, sampleWidth, sampleHeight, m_PreviewDIB.GetDIBDC, 0, 0, vbSrcCopy
        picSample.EndPaint True
    End If
    
End Sub

Private Sub btsSampleMerged_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub btsSampleMerged_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sldRadius.hWndSpinner
    Else
        newTargetHwnd = Me.chkAfter.hWnd
    End If
End Sub

Private Sub cboColorSpace_Click(Index As Integer)
    UpdateUIText
End Sub

Private Sub cboColorSpace_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        If (Index = 0) Then
            newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
        Else
            newTargetHwnd = Me.cboColorSpace(0).hWnd
        End If
    Else
        If (Index = 0) Then
            newTargetHwnd = Me.cboColorSpace(1).hWnd
        Else
            newTargetHwnd = Me.ttlPanel(0).hWnd
        End If
    End If
End Sub

Private Sub chkAfter_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub chkAfter_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsSampleMerged.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
    End If
End Sub

Private Sub cmdFlyoutLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    If (Not m_Flyout Is Nothing) Then m_Flyout.UpdateLockStatus Me.cntrPopOut(Index).hWnd, cmdFlyoutLock(Index).Value, cmdFlyoutLock(Index)
End Sub

Private Sub cmdFlyoutLock_GotFocusAPI(Index As Integer)
    UpdateFlyout Index, True
End Sub

Private Sub cmdFlyoutLock_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then newTargetHwnd = Me.chkAfter.hWnd Else newTargetHwnd = Me.cboColorSpace(0).hWnd
End Sub

Private Sub Form_Load()

    Tools.SetToolBusyState True
    
    Dim i As Long
    For i = cboColorSpace.lBound To cboColorSpace.UBound
        cboColorSpace(i).AddItem "RGB", 0
        cboColorSpace(i).AddItem "RGB %", 1
        cboColorSpace(i).AddItem "HSV", 2
        cboColorSpace(i).AddItem "CMYK", 3
    Next i
    
    'At present, we default to "RGB" in the first color area, and "HSV" in the second
    cboColorSpace(0).ListIndex = cps_RGBA
    cboColorSpace(1).ListIndex = cps_HSV
    
    btsSampleMerged.AddItem "image", 0
    btsSampleMerged.AddItem "layer", 1
    btsSampleMerged.ListIndex = 0
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
    Tools.SetToolBusyState False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If Not (m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If

    'Failsafe only
    If (Not m_Flyout Is Nothing) Then m_Flyout.HideFlyout
    Set m_Flyout = Nothing
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Calculate individual color names on a per-space basis, while accounting for translations
    ReDim m_ColorNames(0 To cps_ColorSpaceCount - 1, 0 To 3) As String
    m_ColorNames(cps_RGBA, 0) = g_Language.TranslateMessage("red")
    m_ColorNames(cps_RGBA, 1) = g_Language.TranslateMessage("green")
    m_ColorNames(cps_RGBA, 2) = g_Language.TranslateMessage("blue")
    m_ColorNames(cps_RGBA, 3) = g_Language.TranslateMessage("opacity")
    
    m_ColorNames(cps_RGBAPercent, 0) = m_ColorNames(cps_RGBA, 0)
    m_ColorNames(cps_RGBAPercent, 1) = m_ColorNames(cps_RGBA, 1)
    m_ColorNames(cps_RGBAPercent, 2) = m_ColorNames(cps_RGBA, 2)
    m_ColorNames(cps_RGBAPercent, 3) = m_ColorNames(cps_RGBA, 3)
    
    m_ColorNames(cps_HSV, 0) = g_Language.TranslateMessage("hue")
    m_ColorNames(cps_HSV, 1) = g_Language.TranslateMessage("saturation")
    m_ColorNames(cps_HSV, 2) = g_Language.TranslateMessage("value")
    m_ColorNames(cps_HSV, 3) = g_Language.TranslateMessage("opacity")
    
    m_ColorNames(cps_CMYK, 0) = g_Language.TranslateMessage("cyan")
    m_ColorNames(cps_CMYK, 1) = g_Language.TranslateMessage("magenta")
    m_ColorNames(cps_CMYK, 2) = g_Language.TranslateMessage("yellow")
    m_ColorNames(cps_CMYK, 3) = g_Language.TranslateMessage("key (black)")
    
    m_NullTextString = g_Language.TranslateMessage("n/a")
    m_StringsInitialized = True
    
    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me
    
    'As language settings may have changed, we now need to redraw the current UI text
    UpdateUIText

End Sub

Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

Private Sub picSample_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (m_PreviewDIB Is Nothing) Then RegenerateColorSampleBox
    GDI.BitBltWrapper targetDC, 0, 0, m_PreviewDIB.GetDIBWidth, m_PreviewDIB.GetDIBHeight, m_PreviewDIB.GetDIBDC, 0, 0, vbSrcCopy
End Sub

Private Sub sldRadius_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub sldRadius_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(0).hWnd
    Else
        newTargetHwnd = Me.btsSampleMerged.hWnd
    End If
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

'Update the actively displayed flyout (if any).  Note that the flyout manager will automatically
' hide any other open flyouts, as necessary.
Private Sub UpdateFlyout(ByVal flyoutIndex As Long, Optional ByVal newState As Boolean = True)
    
    'Ensure we have a flyout manager
    If (m_Flyout Is Nothing) Then Set m_Flyout = New pdFlyout
    
    'Exit if we're already in the process of synchronizing
    If m_Flyout.GetFlyoutSyncState() Then Exit Sub
    m_Flyout.SetFlyoutSyncState True
    
    'Ensure we have a flyout manager, then raise the corresponding panel
    If newState Then
        If (flyoutIndex <> m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.ShowFlyout Me, ttlPanel(flyoutIndex), cntrPopOut(flyoutIndex), flyoutIndex, Interface.FixDPI(-8)
    Else
        If (flyoutIndex = m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.HideFlyout
    End If
    
    'Update titlebar state(s) to match
    Dim i As Long
    For i = ttlPanel.lBound To ttlPanel.UBound
        If (i = m_Flyout.GetFlyoutTrackerID()) Then
            If (Not ttlPanel(i).Value) Then ttlPanel(i).Value = True
        Else
            If ttlPanel(i).Value Then ttlPanel(i).Value = False
        End If
    Next i
    
    'Clear the synchronization flag before exiting
    m_Flyout.SetFlyoutSyncState False
    
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboColorSpace(1).hWnd
    Else
        newTargetHwnd = Me.sldRadius.hWnd
    End If
End Sub
