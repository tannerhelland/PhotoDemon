VERSION 5.00
Begin VB.Form FormPhotoFilters 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Photo filters"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   983
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdListBoxOD lstFilters 
      Height          =   4695
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8281
      Caption         =   "available filters"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   744
      Left            =   0
      TabIndex        =   0
      Top             =   5796
      Width           =   14748
      _ExtentX        =   26009
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
      Top             =   4860
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1270
      Caption         =   "density"
      Min             =   1
      Max             =   100
      SliderTrackStyle=   2
      Value           =   30
      GradientColorLeft=   16777215
      NotchPosition   =   2
      NotchValueCustom=   30
   End
End
Attribute VB_Name = "FormPhotoFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Photo Filter Application Tool
'Copyright 2013-2020 by Tanner Helland
'Created: 06/June/13
'Last updated: 02/August/16
'Last update: totally overhaul UI
'
'Advanced photo filter simulation tool.  A full discussion of photographic filters and how they work are available
' at this Wikipedia article: https://en.wikipedia.org/wiki/Photographic_filter
'
'This code is very similar to PhotoDemon's "Temperature" algorithm.  The main difference is the way the user
' selects a filter to apply.  The available list of filters is flexible, and I have gone to great lengths to find and
' implement a rough correlation for every traditional Wratten (Tiffen) photo filter.
'
'That said, I hope it is abundantly clear that these conversions are all very loose estimations.  Filters work by
' blocking specific wavelengths of light at the moment of photography, so it's impossible to perfectly replicate their
' behavior via code.  All we can do is approximate, so do not expect to get identical results between actual filters
' and post-production tools like PhotoDemon.
'
'Luminosity preservation is assumed.  I could provide a toggle for it, but I see no real benefit to unpreserved use
' of these tools.
'
'The list-box-style interface was custom built for this tool, and I derived it from similar code in the Metadata Browser
' and About dialog.  Please compile for best results; things like mousewheel support and mouseleave tracking require
' subclassing, so they may not behave as expected in the IDE.
'
'I used many resources while attempting to create a list of Wratten filters and their RGB equivalents.  In no particular
' order, thank you to:
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
    Id As String
    Name As String
    Description As String
    RGBColor As Long
End Type

'All current available filters
Private m_Filters() As WrattenFilter
Private m_numOfFilters As Long

'Height of each filter content block
Private Const BLOCKHEIGHT As Long = 53

'Two font objects; one for names and one for descriptions.  (Two are needed because they have different sizes and colors,
' and it is faster to cache these values rather than constantly recreating them on a single pdFont object.)
Private m_TitleFont As pdFont, m_DescriptionFont As pdFont

Private Sub cmdBar_AddCustomPresetData()
    cmdBar.AddPresetData "CurrentFilter", Trim$(Str$(lstFilters.ListIndex))
End Sub

Private Sub cmdBar_OKClick()
    Process "Photo filter", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    lstFilters.ListIndex = CLng(Trim$(cmdBar.RetrievePresetData("CurrentFilter")))
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

'Adding new filters is as simple as passing additional values through this sub
Private Sub AddWratten(ByVal wrattenID As String, ByVal filterColor As String, ByVal filterDescription As String, ByVal filterRGB As Long)
    
    With m_Filters(m_numOfFilters)
        .Id = wrattenID
        .Name = filterColor
        .Description = filterDescription
        .RGBColor = filterRGB
    End With
    m_numOfFilters = m_numOfFilters + 1
    ReDim Preserve m_Filters(0 To m_numOfFilters) As WrattenFilter
    
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.SetPreviewStatus False
    
    'Initialize a custom font object for names
    Set m_TitleFont = New pdFont
    m_TitleFont.SetFontBold True
    m_TitleFont.SetFontSize 12
    m_TitleFont.CreateFontObject
    m_TitleFont.SetTextAlignment vbLeftJustify
    
    '...and a second custom font object for descriptions
    Set m_DescriptionFont = New pdFont
    m_DescriptionFont.SetFontBold False
    m_DescriptionFont.SetFontSize 10
    m_DescriptionFont.CreateFontObject
    m_DescriptionFont.SetTextAlignment vbLeftJustify
    
    m_numOfFilters = 0
    ReDim m_Filters(0) As WrattenFilter
    
    'Add a comprehensive list of Wratten-type filters and their corresponding RGB values.  In the future, this may be moved to an external
    ' (or embedded) XML file.  That would certainly make editing the list easier, though I don't anticipate needing to change it much.
    AddWratten "1A", g_Language.TranslateMessage("skylight (pale pink)"), g_Language.TranslateMessage("reduce haze in landscape photography"), RGB(245, 236, 240)
    AddWratten "2A", g_Language.TranslateMessage("pale yellow"), g_Language.TranslateMessage("absorb UV radiation"), RGB(244, 243, 233)
    AddWratten "2B", g_Language.TranslateMessage("pale yellow"), g_Language.TranslateMessage("absorb UV radiation (slightly less than 2A)"), RGB(244, 245, 230)
    AddWratten "2E", g_Language.TranslateMessage("pale yellow"), g_Language.TranslateMessage("absorb UV radiation (slightly more than 2A)"), RGB(242, 254, 139)
    AddWratten "3", g_Language.TranslateMessage("light yellow"), g_Language.TranslateMessage("absorb excessive sky blue, make sky darker in black/white photos"), RGB(255, 250, 110)
    AddWratten "6", g_Language.TranslateMessage("light yellow"), g_Language.TranslateMessage("absorb excessive sky blue, emphasizing clouds"), RGB(253, 247, 3)
    AddWratten "8", g_Language.TranslateMessage("yellow"), g_Language.TranslateMessage("high blue absorption; correction for sky, cloud, and foliage"), RGB(247, 241, 0)
    AddWratten "9", g_Language.TranslateMessage("deep yellow"), g_Language.TranslateMessage("moderate contrast in black/white outdoor photography"), RGB(255, 228, 0)
    AddWratten "11", g_Language.TranslateMessage("yellow-green"), g_Language.TranslateMessage("correction for tungsten light"), RGB(75, 175, 65)
    AddWratten "12", g_Language.TranslateMessage("deep yellow"), g_Language.TranslateMessage("minus blue; reduce haze in aerial photos"), RGB(255, 220, 0)
    AddWratten "15", g_Language.TranslateMessage("deep yellow"), g_Language.TranslateMessage("darken sky in black/white outdoor photography"), RGB(240, 160, 50)
    AddWratten "16", g_Language.TranslateMessage("yellow-orange"), g_Language.TranslateMessage("stronger version of 15"), RGB(237, 140, 20)
    AddWratten "21", g_Language.TranslateMessage("orange"), g_Language.TranslateMessage("contrast filter for blue and blue-green absorption"), RGB(245, 100, 50)
    AddWratten "22", g_Language.TranslateMessage("deep orange"), g_Language.TranslateMessage("stronger version of 21"), RGB(247, 84, 33)
    AddWratten "23A", g_Language.TranslateMessage("light red"), g_Language.TranslateMessage("contrast effects, darken sky and water"), RGB(255, 117, 106)
    AddWratten "24", g_Language.TranslateMessage("red"), g_Language.TranslateMessage("red for two-color photography (daylight or tungsten)"), RGB(240, 0, 0)
    AddWratten "25", g_Language.TranslateMessage("red"), g_Language.TranslateMessage("tricolor red; contrast effects in outdoor scenes"), RGB(220, 0, 60)
    AddWratten "26", g_Language.TranslateMessage("red"), g_Language.TranslateMessage("stereo red; cuts haze, useful for storm or moonlight settings"), RGB(210, 0, 0)
    AddWratten "29", g_Language.TranslateMessage("deep red"), g_Language.TranslateMessage("color separation; extreme sky darkening in black/white photos"), RGB(115, 10, 25)
    AddWratten "32", g_Language.TranslateMessage("magenta"), g_Language.TranslateMessage("green absorption"), RGB(240, 0, 255)
    AddWratten "33", g_Language.TranslateMessage("strong green absorption"), g_Language.TranslateMessage("variant on 32"), RGB(154, 0, 78)
    AddWratten "34A", g_Language.TranslateMessage("violet"), g_Language.TranslateMessage("minus-green and plus-blue separation"), RGB(124, 40, 240)
    AddWratten "38A", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("red absorption; useful for contrast in microscopy"), RGB(1, 156, 210)
    AddWratten "44", g_Language.TranslateMessage("light blue-green"), g_Language.TranslateMessage("minus-red, two-color general viewing"), RGB(0, 136, 152)
    AddWratten "44A", g_Language.TranslateMessage("light blue-green"), g_Language.TranslateMessage("minus-red, variant on 44"), RGB(0, 175, 190)
    AddWratten "47", g_Language.TranslateMessage("blue tricolor"), g_Language.TranslateMessage("direct color separation; contrast effects in commercial photography"), RGB(43, 75, 220)
    AddWratten "47A", g_Language.TranslateMessage("light blue"), g_Language.TranslateMessage("enhance blue and purple objects; useful for fluorescent dyes"), RGB(0, 15, 150)
    AddWratten "47B", g_Language.TranslateMessage("deep blue tricolor"), g_Language.TranslateMessage("color separation; calibration using SMPTE color bars"), RGB(0, 0, 120)
    AddWratten "56", g_Language.TranslateMessage("very light green"), g_Language.TranslateMessage("darkens sky, improves flesh tones"), RGB(132, 206, 35)
    AddWratten "58", g_Language.TranslateMessage("green tricolor"), g_Language.TranslateMessage("used for color separation; improves definition of foliage"), RGB(40, 110, 5)
    AddWratten "61", g_Language.TranslateMessage("deep green tricolor"), g_Language.TranslateMessage("used for color separation, tungsten tricolor projection"), RGB(40, 70, 10)
    AddWratten "70", g_Language.TranslateMessage("dark red"), g_Language.TranslateMessage("infrared photography longpass filter blocking"), RGB(62, 0, 0)
    AddWratten "80A", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("cooling filter, 3200K to 5500K; converts indoor lighting to sunlight"), RGB(50, 100, 230)
    AddWratten "80B", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("variant of 80A, 3400K to 5500K"), RGB(70, 120, 230)
    AddWratten "80C", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("variant of 80A, 3800K to 5500K"), RGB(90, 140, 235)
    AddWratten "80D", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("variant of 80A, 4200K to 5500K"), RGB(110, 160, 240)
    AddWratten "81A", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter (lowers color temperature), 3400 K to 3200 K"), RGB(247, 240, 220)
    AddWratten "81B", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter; slightly stronger than 81A"), RGB(242, 232, 205)
    AddWratten "81C", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter; slightly stronger than 81B"), RGB(230, 220, 200)
    AddWratten "81D", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter; slightly stronger than 81C"), RGB(235, 220, 190)
    AddWratten "81EF", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter; slightly stronger than 81D"), RGB(215, 185, 150)
    AddWratten "82", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("cooling filter; raises color temperature 100K"), RGB(150, 205, 240)
    AddWratten "82A", g_Language.TranslateMessage("pale blue"), g_Language.TranslateMessage("cooling filter; opposite of 81A"), RGB(205, 225, 235)
    AddWratten "82B", g_Language.TranslateMessage("pale blue"), g_Language.TranslateMessage("cooling filter; opposite of 81B"), RGB(155, 190, 220)
    AddWratten "82C", g_Language.TranslateMessage("pale blue"), g_Language.TranslateMessage("cooling filter; opposite of 81C"), RGB(120, 155, 190)
    AddWratten "85", g_Language.TranslateMessage("amber"), g_Language.TranslateMessage("warming filter, 5500K to 3400K; converts sunlight to incandescent"), RGB(250, 155, 115)
    AddWratten "85B", g_Language.TranslateMessage("amber"), g_Language.TranslateMessage("warming filter, 5500K to 3200K; opposite of 80A"), RGB(250, 125, 95)
    AddWratten "85C", g_Language.TranslateMessage("amber"), g_Language.TranslateMessage("warming filter, 5500K to 3800K; opposite of 80C"), RGB(250, 155, 115)
    AddWratten "90", g_Language.TranslateMessage("dark gray amber"), g_Language.TranslateMessage("remove color before photographing; rarely used for actual photos"), RGB(100, 85, 20)
    AddWratten "96", g_Language.TranslateMessage("neutral gray"), g_Language.TranslateMessage("neutral density filter; equally blocks all light frequencies"), RGB(100, 100, 100)
    
    'Add dummy entries to the owner-drawn listbox, so that it's initialized to the proper size and layout
    lstFilters.ListItemHeight = FixDPI(BLOCKHEIGHT)
    lstFilters.SetAutomaticRedraws False
    Dim i As Long
    For i = 0 To m_numOfFilters - 1
        lstFilters.AddItem , i
    Next i
    lstFilters.SetAutomaticRedraws True, True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lstFilters_Click()
    UpdatePreview
End Sub

Private Sub lstFilters_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    If (bufferDC = 0) Then Exit Sub
    
    'Retrieve the boundary region for this list entry
    Dim tmpRectF As RectF
    CopyMemory ByVal VarPtr(tmpRectF), ByVal ptrToRectF, 16&
    
    Dim offsetY As Single, offsetX As Single
    offsetX = tmpRectF.Left
    offsetY = tmpRectF.Top
    
    Dim linePadding As Long
    linePadding = FixDPI(2)
    
    'pd2D is used for extra drawing capabilities
    Dim cPen As pd2DPen, cBrush As pd2DBrush, cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
        
    'Modify font colors to match the current selection state of this item
    If itemIsSelected Then
        m_TitleFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected, True, False, itemIsHovered)
        m_DescriptionFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected, True, False, itemIsHovered)
    Else
        m_TitleFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, True, False, itemIsHovered)
        m_DescriptionFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, True, False, itemIsHovered)
    End If
    
    'Render a color box that represents the color of this filter
    Dim colorRectF As RectF
    With colorRectF
        .Width = FixDPIFloat(64)
        .Height = FixDPIFloat(42)
        .Left = offsetX + FixDPIFloat(4)
        .Top = offsetY + ((FixDPIFloat(BLOCKHEIGHT) - .Height) / 2)
    End With
    
    Drawing2D.QuickCreateSolidBrush cBrush, m_Filters(itemIndex).RGBColor
    PD2D.FillRectangleF_FromRectF cSurface, cBrush, colorRectF
    
    Drawing2D.QuickCreateSolidPen cPen, 1#, g_Themer.GetGenericUIColor(UI_GrayDefault)
    PD2D.DrawRectangleF_FromRectF cSurface, cPen, colorRectF
    
    'To minimize confusion, release the surface now; all subsequent objects will draw directly to the target DC
    Set cSurface = Nothing
        
    'Render the Wratten ID and name fields
    Dim textLeft As Long
    textLeft = colorRectF.Left + colorRectF.Width + FixDPI(8)
    
    Dim titleString As String
    titleString = m_Filters(itemIndex).Id & " - " & m_Filters(itemIndex).Name
    m_TitleFont.AttachToDC bufferDC
    m_TitleFont.FastRenderText textLeft, offsetY + FixDPI(4), titleString
    
    'Calculate the drop-down for the description line
    Dim lineHeight As Single
    lineHeight = m_TitleFont.GetHeightOfString(titleString) + linePadding
    m_TitleFont.ReleaseFromDC
    
    'Below that, add the description text
    Dim descriptionString As String
    descriptionString = m_Filters(itemIndex).Description
    m_DescriptionFont.AttachToDC bufferDC
    m_DescriptionFont.FastRenderText textLeft, offsetY + FixDPI(4) + lineHeight, descriptionString
    m_DescriptionFont.ReleaseFromDC
    
End Sub

Private Sub sltDensity_Change()
    UpdatePreview
End Sub

'Render a new preview
Private Sub UpdatePreview()
    
    If (lstFilters.ListIndex >= 0) Then
    
        'Sync the density slider to match the currently selected color
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
    
    Dim filterColor As Long, filterDensity As Double, preserveLuminance As Boolean
    
    With cParams
        filterColor = .GetLong("color")
        filterDensity = .GetDouble("density", 0#)
        preserveLuminance = .GetBool("preserveluminance", True)
    End With
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
    Dim originalLuminance As Double
    Dim tmpR As Long, tmpG As Long, tmpB As Long
            
    'Extract the RGB values from the color we were passed
    tmpR = Colors.ExtractRed(filterColor)
    tmpG = Colors.ExtractGreen(filterColor)
    tmpB = Colors.ExtractBlue(filterColor)
            
    'Divide tempStrength by 100 to yield a value between 0 and 1
    filterDensity = filterDensity / 100
            
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
    
        'Get the source pixel color values
        b = imageData(xStride, y)
        g = imageData(xStride + 1, y)
        r = imageData(xStride + 2, y)
        
        'If luminance is being preserved, we need to determine the initial luminance value
        originalLuminance = (GetLuminance(r, g, b) / 255#)
        
        'Blend the original and new RGB values using the specified strength
        r = BlendColors(r, tmpR, filterDensity)
        g = BlendColors(g, tmpG, filterDensity)
        b = BlendColors(b, tmpB, filterDensity)
        
        'If the user wants us to preserve luminance, determine the hue and saturation of the new color, then replace the luminance
        ' value with the original
        If preserveLuminance Then
            ImpreciseRGBtoHSL r, g, b, h, s, l
            ImpreciseHSLtoRGB h, s, originalLuminance, r, g, b
        End If
        
        'Assign the new values to each color channel
        imageData(xStride, y) = b
        imageData(xStride + 1, y) = g
        imageData(xStride + 2, y) = r
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
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
        .AddParam "color", m_Filters(lstFilters.ListIndex).RGBColor
        .AddParam "density", sltDensity.Value
        .AddParam "preserveluminance", True
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
