VERSION 5.00
Begin VB.Form FormPhotoFilters 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Photo filter"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   778
   Begin PhotoDemon.pdColorSelector csColor 
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   3360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1720
      Caption         =   "color"
   End
   Begin PhotoDemon.pdListBoxOD lstFilters 
      Height          =   3015
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6800
      Caption         =   "filter"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltDensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   4440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "density"
      Min             =   1
      Max             =   100
      SliderTrackStyle=   2
      Value           =   30
      GradientColorLeft=   16777215
      NotchPosition   =   2
      NotchValueCustom=   30
   End
   Begin PhotoDemon.pdCheckBox chkLuminance 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   5280
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   661
      Caption         =   "preserve luminance"
   End
End
Attribute VB_Name = "FormPhotoFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Photo Filter Application Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 06/June/13
'Last updated: 27/October/20
'Last update: greatly simplify tool to bring it in line with Photoshop
'
'Photo filter simulation tool.  For a brief overview of how photographic filters theoretically work,
' consult Wikipedia: https://en.wikipedia.org/wiki/Photographic_filter
'
'The list of available filters mirrors Photoshop's.  There's not much reason to provide this tool
' other than to "fill a gap" in Photoshop tutorials that leverage it.  I've at least tried to improve
' the UI and make the tool a little more interesting visually than Photoshop's, and we also perform
' our conversion in L*ab color space for ideal results (something PS does not, to my knowledge).
'
'For a better understanding of "true" photo filters (e.g. actual hardware that you stick on a
' camera lens), consider these useful resources:
'
'http://www.karmalimbo.com/aro/pics/filters/transmision%20of%20wratten%20filters.pdf
'https://en.wikipedia.org/wiki/Wratten_85#Reference_table
'http://www.redisonellis.com/wratten.html
'http://www.vistaview360.com/cameras/filters_by_the_numbers.htm
'https://en.wikipedia.org/wiki/Photographic_filter
'http://photo.net/learn/optics/edscott/cf000020.htm
'http://www.olympusmicro.com/primer/photomicrography/bwfilters.html
'http://web.archive.org/web/20091028192325/http://www.geocities.com/cokinfiltersystem/color_corection.htm
'http://www.filmcentre.co.uk/faqs_filter.htm
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Specialized type for holding photographic filter information
Private Type WrattenFilter
    Name As String
    RGBColor As Long
End Type

'All current available filters
Private m_Filters() As WrattenFilter
Private m_numOfFilters As Long

'Height of each filter content block, in pixels at 96 DPI.
' (This will be automatically scaled at run-time, as necessary.)
Private Const BLOCKHEIGHT As Long = 28

'Font object for rendering list captions
Private m_TitleFont As pdFont

'Avoid stack crashes
Private m_SourceIsColor As Boolean, m_SourceIsList As Boolean

Private Sub chkLuminance_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Photo filter", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

'Adding new filters is as simple as passing additional values through this sub
Private Sub AddPhotoFilter(ByVal filterName As String, ByVal filterRGB As Long)
    With m_Filters(m_numOfFilters)
        .Name = filterName
        .RGBColor = filterRGB
    End With
    m_numOfFilters = m_numOfFilters + 1
    ReDim Preserve m_Filters(0 To m_numOfFilters) As WrattenFilter
End Sub

Private Sub csColor_ColorChanged()
    If (Not m_SourceIsList) Then
        m_SourceIsColor = True
        lstFilters.ListIndex = m_numOfFilters - 1
        m_SourceIsColor = False
    End If
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.SetPreviewStatus False
    
    'Initialize a custom font object for names
    Set m_TitleFont = New pdFont
    m_TitleFont.SetFontBold False
    m_TitleFont.SetFontSize 10
    m_TitleFont.CreateFontObject
    m_TitleFont.SetTextAlignment vbLeftJustify
    
    m_numOfFilters = 0
    ReDim m_Filters(0) As WrattenFilter
    
    'Add a comprehensive list of Wratten-type filters and their corresponding RGB values.  In the future, this may be moved to an external
    ' (or embedded) XML file.  That would certainly make editing the list easier, though I don't anticipate needing to change it much.
    AddPhotoFilter g_Language.TranslateMessage("Warming filter (%1)", "85"), RGB(236, 138, 0)
    AddPhotoFilter g_Language.TranslateMessage("Warming filter (%1)", "LBA"), RGB(250, 150, 0)
    AddPhotoFilter g_Language.TranslateMessage("Warming filter (%1)", "81"), RGB(235, 177, 19)
    AddPhotoFilter g_Language.TranslateMessage("Cooling filter (%1)", "80"), RGB(0, 109, 255)
    AddPhotoFilter g_Language.TranslateMessage("Cooling filter (%1)", "LBB"), RGB(0, 93, 255)
    AddPhotoFilter g_Language.TranslateMessage("Cooling filter (%1)", "82"), RGB(0, 181, 255)
    AddPhotoFilter g_Language.TranslateMessage("Red"), RGB(234, 26, 26)
    AddPhotoFilter g_Language.TranslateMessage("Orange"), RGB(243, 162, 23)
    AddPhotoFilter g_Language.TranslateMessage("Yellow"), RGB(249, 227, 28)
    AddPhotoFilter g_Language.TranslateMessage("Green"), RGB(25, 201, 25)
    AddPhotoFilter g_Language.TranslateMessage("Cyan"), RGB(29, 203, 234)
    AddPhotoFilter g_Language.TranslateMessage("Blue"), RGB(29, 53, 234)
    AddPhotoFilter g_Language.TranslateMessage("Violet"), RGB(155, 29, 234)
    AddPhotoFilter g_Language.TranslateMessage("Magenta"), RGB(227, 24, 227)
    AddPhotoFilter g_Language.TranslateMessage("Sepia"), RGB(172, 122, 51)
    AddPhotoFilter g_Language.TranslateMessage("Deep red"), RGB(255, 0, 0)
    AddPhotoFilter g_Language.TranslateMessage("Deep blue"), RGB(0, 34, 205)
    AddPhotoFilter g_Language.TranslateMessage("Deep emerald"), RGB(0, 140, 0)
    AddPhotoFilter g_Language.TranslateMessage("Deep yellow"), RGB(255, 213, 0)
    AddPhotoFilter g_Language.TranslateMessage("Underwater"), RGB(0, 193, 177)
    AddPhotoFilter g_Language.TranslateMessage("Custom"), RGB(127, 127, 127)
    
    'Add dummy entries to the owner-drawn listbox, so that it's initialized to the proper size and layout
    lstFilters.ListItemHeight = Interface.FixDPI(BLOCKHEIGHT)
    lstFilters.SetAutomaticRedraws False
    Dim i As Long
    For i = 0 To m_numOfFilters - 1
        lstFilters.AddItem vbNullString, i
    Next i
    lstFilters.ListIndex = 0
    lstFilters.SetAutomaticRedraws True, True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lstFilters_Click()
    If (Not m_SourceIsColor) Then
        m_SourceIsList = True
        csColor.Color = m_Filters(lstFilters.ListIndex).RGBColor
        m_SourceIsList = False
    End If
End Sub

Private Sub lstFilters_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    If (bufferDC = 0) Then Exit Sub
    
    'Retrieve the boundary region for this list entry
    Dim tmpRectF As RectF
    CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
    
    Dim offsetY As Single, offsetX As Single
    offsetX = tmpRectF.Left
    offsetY = tmpRectF.Top
    
    Dim linePadding As Long
    linePadding = Interface.FixDPI(2)
    
    Dim curBlockHeight As Long
    curBlockHeight = Interface.FixDPI(BLOCKHEIGHT)
    
    'pd2D is used for extra drawing capabilities
    Dim cPen As pd2DPen, cBrush As pd2DBrush, cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, False
        
    'Modify font colors to match the current selection state of this item
    If itemIsSelected Then
        m_TitleFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected, True, False, itemIsHovered)
    Else
        m_TitleFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, True, False, itemIsHovered)
    End If
    
    'Render a color box that represents the color of this filter
    Const COLOR_BOX_PADDING As Long = 3
    Dim colorRectF As RectF
    With colorRectF
        .Width = Interface.FixDPI(48)
        .Height = curBlockHeight - COLOR_BOX_PADDING * 2
        .Left = offsetX + Interface.FixDPI(4)
        .Top = offsetY + COLOR_BOX_PADDING
    End With
    
    If (itemIndex < m_numOfFilters - 1) Then
        Drawing2D.QuickCreateSolidBrush cBrush, m_Filters(itemIndex).RGBColor
        PD2D.FillRectangleF_FromRectF cSurface, cBrush, colorRectF
        Drawing2D.QuickCreateSolidPen cPen, 1!, g_Themer.GetGenericUIColor(UI_GrayDefault)
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, colorRectF
    End If
    
    'To minimize confusion, release the surface now; all subsequent rendering is text, which is easier
    ' to draw directly onto the target DC
    Set cSurface = Nothing
        
    'Render the Wratten ID and name fields
    Dim textLeft As Long
    textLeft = colorRectF.Left + colorRectF.Width + FixDPI(8)
    
    Dim titleString As String, titleHeight As Long
    titleString = m_Filters(itemIndex).Name
    m_TitleFont.AttachToDC bufferDC
    titleHeight = m_TitleFont.GetHeightOfString(titleString)
    m_TitleFont.FastRenderText textLeft, offsetY + (BLOCKHEIGHT - titleHeight) \ 2, titleString
    m_TitleFont.ReleaseFromDC
    
End Sub

Private Sub sltDensity_Change()
    UpdatePreview
End Sub

'Render a new preview
Private Sub UpdatePreview()
    
    If (lstFilters.ListIndex >= 0) Then
    
        'Sync the density slider's background gradient to match the currently selected color
        If (sltDensity.GradientColorRight <> m_Filters(lstFilters.ListIndex).RGBColor) Then sltDensity.GradientColorRight = m_Filters(lstFilters.ListIndex).RGBColor
        
        'Render the preview
        If cmdBar.PreviewsAllowed Then Me.ApplyPhotoFilter GetLocalParamString(), True, pdFxPreview
    
    End If
    
End Sub

'Cast an image with a new temperature value
' Input: desired temperature, whether to preserve luminance or not, and a blend ratio between 1 and 100
Public Sub ApplyPhotoFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Applying photo filter..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim filterColor As Long, filterDensity As Double, changeLuminance As Boolean
    
    With cParams
        filterColor = .GetLong("color")
        filterDensity = .GetDouble("density", 0#)
        changeLuminance = Not .GetBool("preserveluminance", True)
    End With
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left * 4
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * 4
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Cache some values to improve performance in the inner loop
    filterDensity = filterDensity * 0.01
    Dim invFilterDensity As Double
    invFilterDensity = 1# - filterDensity
    
    'Use LCMS to create an RGB to LAB transform (and vice-versa)
    Dim cRGB As pdLCMSProfile
    Set cRGB = New pdLCMSProfile
    cRGB.CreateSRGBProfile True
    
    Dim cLAB As pdLCMSProfile
    Set cLAB = New pdLCMSProfile
    cLAB.CreateLabProfile True
    
    Dim cTransformRGBtoLAB As pdLCMSTransform
    Set cTransformRGBtoLAB = New pdLCMSTransform
    cTransformRGBtoLAB.CreateTwoProfileTransform cRGB, cLAB, TYPE_BGRA_8, TYPE_ALab_8, INTENT_PERCEPTUAL
    
    Dim cTransformLABtoRGB As pdLCMSTransform
    Set cTransformLABtoRGB = New pdLCMSTransform
    cTransformLABtoRGB.CreateTwoProfileTransform cLAB, cRGB, TYPE_ALab_8, TYPE_BGRA_8, INTENT_PERCEPTUAL
    
    'Get LAB values of the target color
    Dim srcColorRGBA As RGBQuad
    srcColorRGBA.Blue = Colors.ExtractBlue(filterColor)
    srcColorRGBA.Green = Colors.ExtractGreen(filterColor)
    srcColorRGBA.Red = Colors.ExtractRed(filterColor)
    srcColorRGBA.Alpha = 255
    
    Dim srcColorLabA As RGBQuad
    cTransformRGBtoLAB.ApplyTransformToScanline VarPtr(srcColorRGBA), VarPtr(srcColorLabA), 1
    
    Dim srcL As Long, srcA As Long, srcB As Long
    srcB = srcColorLabA.Green
    srcA = srcColorLabA.Red
    srcL = srcColorLabA.Alpha
    
    'Resize the LAB array to the same size as a scanline of the source image;
    ' we'll translate RGB to LAB values a scanline at a time
    Dim srcPixelsLab() As Byte
    ReDim srcPixelsLab(0 To workingDIB.GetDIBStride - 1) As Byte
    
    Dim labL As Long, labA As Long, labB As Long
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        
        'Translate this scanline into LAB color
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
        cTransformRGBtoLAB.ApplyTransformToScanline VarPtr(imageData(0)), VarPtr(srcPixelsLab(0)), workingDIB.GetDIBWidth
        
    For x = initX To finalX Step 4
        
        'Get the source pixel color values (LAB color space)
        labB = srcPixelsLab(x + 1)
        labA = srcPixelsLab(x + 2)
        labL = srcPixelsLab(x + 3)
        
        'Perform the blend in LAB
        labB = Int(filterDensity * srcB) + Int(invFilterDensity * labB + 0.5)
        labA = Int(filterDensity * srcA) + Int(invFilterDensity * labA + 0.5)
        If changeLuminance Then labL = Int(filterDensity * srcL) + Int(invFilterDensity * labL + 0.5)
        
        'Copy the merged LAB results back into the dedicated LAB array
        srcPixelsLab(x + 1) = labB
        srcPixelsLab(x + 2) = labA
        srcPixelsLab(x + 3) = labL
        
    Next x
    
        'Translate all LAB results back into RGB in a single pass
        cTransformLABtoRGB.ApplyTransformToScanline VarPtr(srcPixelsLab(0)), VarPtr(imageData(0)), workingDIB.GetDIBWidth
        
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
        
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "color", csColor.Color
        .AddParam "density", sltDensity.Value
        .AddParam "preserveluminance", chkLuminance.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
