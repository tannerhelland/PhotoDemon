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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   744
      Left            =   0
      TabIndex        =   0
      Top             =   5796
      Width           =   14748
      _ExtentX        =   26009
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      FillColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   6000
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   551
      TabIndex        =   5
      Top             =   480
      Width           =   8295
   End
   Begin VB.VScrollBar vsFilter 
      Height          =   4185
      LargeChange     =   32
      Left            =   14280
      TabIndex        =   4
      Top             =   480
      Width           =   330
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltDensity 
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
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "available filters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FormPhotoFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Photo Filter Application Tool
'Copyright 2013-2015 by Tanner Helland
'Created: 06/June/13
'Last updated: 18/June/14
'Last update: add arrow key nav support to the custom list
'
'Advanced photo filter simulation tool.  A full discussion of photographic filters and how they work are available
' at this Wikipedia article: http://en.wikipedia.org/wiki/Photographic_filter
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
'http://en.wikipedia.org/wiki/Wratten_85#Reference_table
'http://www.redisonellis.com/wratten.html
'http://www.vistaview360.com/cameras/filters_by_the_numbers.htm
'http://en.wikipedia.org/wiki/Photographic_filter
'http://photo.net/learn/optics/edscott/cf000020.htm
'http://www.olympusmicro.com/primer/photomicrography/bwfilters.html
'http://web.archive.org/web/20091028192325/http://www.geocities.com/cokinfiltersystem/color_corection.htm
'http://www.filmcentre.co.uk/faqs_filter.htm
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Specialized type for holding photographic filter information
Private Type wrattenFilter
    Id As String
    Name As String
    Description As String
    RGBColor As Long
End Type

'All current available filters
Dim fArray() As wrattenFilter
Dim numOfFilters As Long

'Height of each filter content block
Private Const BLOCKHEIGHT As Long = 53

'An outside class provides access to mousewheel events for scrolling the filter view
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1
Private WithEvents cKeyEvents As pdInputKeyboard
Attribute cKeyEvents.VB_VarHelpID = -1

'Extra variables for custom list rendering
Dim bufferDIB As pdDIB
Dim m_BufferWidth As Long, m_BufferHeight As Long

'Two font objects; one for names and one for descriptions.  (Two are needed because they have different sizes and colors,
' and it is faster to cache these values rather than constantly recreating them on a single pdFont object.)
Dim firstFont As pdFont, secondFont As pdFont

'A primary and secondary color for font rendering
Dim primaryColor As Long, secondaryColor As Long

'The currently selected and currently hovered filter entry
Dim curFilter As Long, curFilterHover As Long

'Redraw the current list of filters
Private Sub redrawFilterList()
        
    Dim scrollOffset As Long
    scrollOffset = vsFilter.Value
    
    bufferDIB.createBlank picBuffer.ScaleWidth, picBuffer.ScaleHeight
    
    Dim i As Long
    For i = 0 To numOfFilters - 1
        renderFilterBlock i, 0, FixDPI(i * BLOCKHEIGHT) - scrollOffset - FixDPI(2)
    Next i
    
    'Copy the buffer to the main form
    BitBlt picBuffer.hDC, 0, 0, m_BufferWidth, m_BufferHeight, bufferDIB.getDIBDC, 0, 0, vbSrcCopy
    picBuffer.Picture = picBuffer.Image
    picBuffer.Refresh
    
End Sub

'Render an individual "block" for a given filter (including name, description, color)
Private Sub renderFilterBlock(ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + FixDPI(BLOCKHEIGHT)) >= 0) And (offsetY <= m_BufferHeight) Then
    
        offsetY = offsetY + FixDPI(2)
        
        Dim linePadding As Long
        linePadding = FixDPI(2)
    
        Dim mHeight As Single
        Dim tmpRect As RECTL
        Dim hBrush As Long
        
        'If this filter has been selected, draw the background with the system's current selection color
        If blockIndex = curFilter Then
        
            SetRect tmpRect, offsetX, offsetY, m_BufferWidth, offsetY + FixDPI(BLOCKHEIGHT)
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FillRect bufferDIB.getDIBDC, tmpRect, hBrush
            DeleteObject hBrush
            
            'Also, color the fonts with the matching highlighted text color (otherwise they won't be readable)
            firstFont.SetFontColor ConvertSystemColor(vbHighlightText)
            secondFont.SetFontColor ConvertSystemColor(vbHighlightText)
        
        Else
            firstFont.SetFontColor primaryColor
            secondFont.SetFontColor secondaryColor
        End If
        
        'If the current filter is highlighted but not selected, simply render the border with a highlight
        If (blockIndex <> curFilter) And (blockIndex = curFilterHover) Then
            SetRect tmpRect, offsetX, offsetY, m_BufferWidth, offsetY + FixDPI(BLOCKHEIGHT)
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FrameRect bufferDIB.getDIBDC, tmpRect, hBrush
            DeleteObject hBrush
        End If
        
        Dim drawString As String
        drawString = fArray(blockIndex).Id & " - " & fArray(blockIndex).Name
        
        'Render a color box for the image
        Dim colorWidth As Long, colorHeight As Long
        colorWidth = FixDPI(64)
        colorHeight = FixDPI(42)
        SetRect tmpRect, offsetX + FixDPI(4), offsetY + ((FixDPI(BLOCKHEIGHT) - colorHeight) \ 2), offsetX + FixDPI(4) + colorWidth, offsetY + ((FixDPI(BLOCKHEIGHT) - colorHeight) \ 2) + colorHeight
        
        hBrush = CreateSolidBrush(fArray(blockIndex).RGBColor)
        FillRect bufferDIB.getDIBDC, tmpRect, hBrush
        DeleteObject hBrush
        hBrush = CreateSolidBrush(RGB(64, 64, 64))
        FrameRect bufferDIB.getDIBDC, tmpRect, hBrush
        DeleteObject hBrush
            
        'Render the Wratten ID and name fields
        firstFont.AttachToDC bufferDIB.getDIBDC
        firstFont.FastRenderText colorWidth + FixDPI(16) + offsetX, offsetY + FixDPI(4), drawString
        
        'Calculate the drop-down for the description line
        mHeight = firstFont.GetHeightOfString(drawString) + linePadding
        firstFont.ReleaseFromDC
        
        'Below that, add the description text
        drawString = fArray(blockIndex).Description
        
        secondFont.AttachToDC bufferDIB.getDIBDC
        secondFont.FastRenderText colorWidth + FixDPI(16) + offsetX, offsetY + FixDPI(4) + mHeight, drawString
        secondFont.ReleaseFromDC
        
    End If

End Sub

Private Sub cKeyEvents_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    'Up and down arrows navigate the list
    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Then
    
        If (vkCode = VK_UP) Then
            curFilter = curFilter - 1
            If curFilter < 0 Then curFilter = numOfFilters - 1
        End If
        
        If (vkCode = VK_DOWN) Then
            curFilter = curFilter + 1
            If curFilter >= numOfFilters Then curFilter = 0
        End If
        
        'Calculate a new vertical scroll position so that the selected filter appears on-screen
        Dim newScrollOffset As Long
        newScrollOffset = curFilter * FixDPI(BLOCKHEIGHT)
        If newScrollOffset > vsFilter.Max Then newScrollOffset = vsFilter.Max
        vsFilter.Value = newScrollOffset
        
        'Redraw the custom filter list
        redrawFilterList
        
    End If
    
    'Right and left arrows modify strength
    If (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then
        
        cmdBar.markPreviewStatus False
        If (vkCode = VK_RIGHT) Then sltDensity.Value = sltDensity.Value + 10
        If (vkCode = VK_LEFT) Then sltDensity.Value = sltDensity.Value - 10
        cmdBar.markPreviewStatus True
        
    End If
    
    updatePreview

End Sub

Private Sub cmdBar_AddCustomPresetData()
    cmdBar.addPresetData "CurrentFilter", Str(curFilter)
End Sub

Private Sub cmdBar_OKClick()
    Process "Photo filter", , buildParams(fArray(curFilter).RGBColor, sltDensity.Value, True), UNDO_LAYER
End Sub

Private Sub cmdBar_RandomizeClick()

    'This is sloppy, but effective.  The vertical scroll bar will be randomly set; we can thus fake a click
    ' somewhere inside the picture box to simulate selecting a random photo filter.
    cMouseEvents_MouseDownCustom pdLeftButton, 0, Rnd * picBuffer.ScaleWidth, Rnd * picBuffer.ScaleHeight

End Sub

Private Sub cmdBar_ReadCustomPresetData()
    curFilter = CLng(cmdBar.retrievePresetData("CurrentFilter"))
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    redrawFilterList
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'No filters are currently selected or hovered
    curFilter = 0
    curFilterHover = -1
    
    'Density is 30 by default
    sltDensity.Value = 30
    
    'Remove any active effect
    redrawFilterList
    
End Sub

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    curFilter = getFilterAtPosition(x, y)
    redrawFilterList
    updatePreview

End Sub

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEvents.setSystemCursor IDC_HAND
End Sub

'When the mouse leaves the filter box, remove any hovered entries and redraw
Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEvents.setSystemCursor IDC_DEFAULT
    curFilterHover = -1
    redrawFilterList
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    curFilterHover = getFilterAtPosition(x, y)
    redrawFilterList
    
End Sub

Private Sub cMouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

    'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
    If vsFilter.Visible Then
  
        If scrollAmount < 0 Then
            
            If vsFilter.Value + vsFilter.LargeChange > vsFilter.Max Then
                vsFilter.Value = vsFilter.Max
            Else
                vsFilter.Value = vsFilter.Value + vsFilter.LargeChange
            End If
            
            curFilterHover = getFilterAtPosition(x, y)
            redrawFilterList
        
        ElseIf scrollAmount > 0 Then
            
            If vsFilter.Value - vsFilter.LargeChange < vsFilter.Min Then
                vsFilter.Value = vsFilter.Min
            Else
                vsFilter.Value = vsFilter.Value - vsFilter.LargeChange
            End If
            
            curFilterHover = getFilterAtPosition(x, y)
            redrawFilterList
            
        End If
        
    End If

End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Display the previewed effect in the neighboring window, then render the list of available filters
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

'Adding new filters is as simple as passing additional values through this sub
Private Sub addWratten(ByVal wrattenID As String, ByVal filterColor As String, ByVal filterDescription As String, ByVal filterRGB As Long)
    
    With fArray(numOfFilters)
        .Id = wrattenID
        .Name = filterColor
        .Description = filterDescription
        .RGBColor = filterRGB
    End With
    numOfFilters = numOfFilters + 1
    ReDim Preserve fArray(0 To numOfFilters) As wrattenFilter
    
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.markPreviewStatus False

    'Enable mousewheel scrolling for the filter box
    Set cMouseEvents = New pdInputMouse
    cMouseEvents.addInputTracker picBuffer.hWnd, True, True, , True
    cMouseEvents.addInputTracker Me.hWnd
    cMouseEvents.setSystemCursor IDC_HAND
    
    'Track a few keypresses to make list navigation easier
    Set cKeyEvents = New pdInputKeyboard
    cKeyEvents.createKeyboardTracker "Photo Filters picBuffer", picBuffer.hWnd, VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN
    
    'Create a background buffer the same size as the buffer picture box
    Set bufferDIB = New pdDIB
    bufferDIB.createBlank picBuffer.ScaleWidth, picBuffer.ScaleHeight
    
    'Initialize a few other variables now (for performance reasons)
    m_BufferWidth = picBuffer.ScaleWidth
    m_BufferHeight = picBuffer.ScaleHeight
    
    'Initialize a custom font object for names
    primaryColor = RGB(64, 64, 64)
    Set firstFont = New pdFont
    firstFont.SetFontColor primaryColor
    firstFont.SetFontBold True
    firstFont.SetFontSize 12
    firstFont.CreateFontObject
    firstFont.SetTextAlignment vbLeftJustify
    
    '...and a second custom font object for descriptions
    secondaryColor = RGB(92, 92, 92)
    Set secondFont = New pdFont
    secondFont.SetFontColor secondaryColor
    secondFont.SetFontBold False
    secondFont.SetFontSize 10
    secondFont.CreateFontObject
    secondFont.SetTextAlignment vbLeftJustify
    
    numOfFilters = 0
    ReDim fArray(0) As wrattenFilter
    
    'Add a comprehensive list of Wratten-type filters and their corresponding RGB values.  In the future, this may be moved to an external
    ' (or embedded) XML file.  That would certainly make editing the list easier, though I don't anticipate needing to change it much.
    addWratten "1A", g_Language.TranslateMessage("skylight (pale pink)"), g_Language.TranslateMessage("reduce haze in landscape photography"), RGB(245, 236, 240)
    addWratten "2A", g_Language.TranslateMessage("pale yellow"), g_Language.TranslateMessage("absorb UV radiation"), RGB(244, 243, 233)
    addWratten "2B", g_Language.TranslateMessage("pale yellow"), g_Language.TranslateMessage("absorb UV radiation (slightly less than 2A)"), RGB(244, 245, 230)
    addWratten "2E", g_Language.TranslateMessage("pale yellow"), g_Language.TranslateMessage("absorb UV radiation (slightly more than 2A)"), RGB(242, 254, 139)
    addWratten "3", g_Language.TranslateMessage("light yellow"), g_Language.TranslateMessage("absorb excessive sky blue, make sky darker in black/white photos"), RGB(255, 250, 110)
    addWratten "6", g_Language.TranslateMessage("light yellow"), g_Language.TranslateMessage("absorb excessive sky blue, emphasizing clouds"), RGB(253, 247, 3)
    addWratten "8", g_Language.TranslateMessage("yellow"), g_Language.TranslateMessage("high blue absorption; correction for sky, cloud, and foliage"), RGB(247, 241, 0)
    addWratten "9", g_Language.TranslateMessage("deep yellow"), g_Language.TranslateMessage("moderate contrast in black/white outdoor photography"), RGB(255, 228, 0)
    addWratten "11", g_Language.TranslateMessage("yellow-green"), g_Language.TranslateMessage("correction for tungsten light"), RGB(75, 175, 65)
    addWratten "12", g_Language.TranslateMessage("deep yellow"), g_Language.TranslateMessage("minus blue; reduce haze in aerial photos"), RGB(255, 220, 0)
    addWratten "15", g_Language.TranslateMessage("deep yellow"), g_Language.TranslateMessage("darken sky in black/white outdoor photography"), RGB(240, 160, 50)
    addWratten "16", g_Language.TranslateMessage("yellow-orange"), g_Language.TranslateMessage("stronger version of 15"), RGB(237, 140, 20)
    addWratten "21", g_Language.TranslateMessage("orange"), g_Language.TranslateMessage("contrast filter for blue and blue-green absorption"), RGB(245, 100, 50)
    addWratten "22", g_Language.TranslateMessage("deep orange"), g_Language.TranslateMessage("stronger version of 21"), RGB(247, 84, 33)
    addWratten "23A", g_Language.TranslateMessage("light red"), g_Language.TranslateMessage("contrast effects, darken sky and water"), RGB(255, 117, 106)
    addWratten "24", g_Language.TranslateMessage("red"), g_Language.TranslateMessage("red for two-color photography (daylight or tungsten)"), RGB(240, 0, 0)
    addWratten "25", g_Language.TranslateMessage("red"), g_Language.TranslateMessage("tricolor red; contrast effects in outdoor scenes"), RGB(220, 0, 60)
    addWratten "26", g_Language.TranslateMessage("red"), g_Language.TranslateMessage("stereo red; cuts haze, useful for storm or moonlight settings"), RGB(210, 0, 0)
    addWratten "29", g_Language.TranslateMessage("deep red"), g_Language.TranslateMessage("color separation; extreme sky darkening in black/white photos"), RGB(115, 10, 25)
    addWratten "32", g_Language.TranslateMessage("magenta"), g_Language.TranslateMessage("green absorption"), RGB(240, 0, 255)
    addWratten "33", g_Language.TranslateMessage("strong green absorption"), g_Language.TranslateMessage("variant on 32"), RGB(154, 0, 78)
    addWratten "34A", g_Language.TranslateMessage("violet"), g_Language.TranslateMessage("minus-green and plus-blue separation"), RGB(124, 40, 240)
    addWratten "38A", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("red absorption; useful for contrast in microscopy"), RGB(1, 156, 210)
    addWratten "44", g_Language.TranslateMessage("light blue-green"), g_Language.TranslateMessage("minus-red, two-color general viewing"), RGB(0, 136, 152)
    addWratten "44A", g_Language.TranslateMessage("light blue-green"), g_Language.TranslateMessage("minus-red, variant on 44"), RGB(0, 175, 190)
    addWratten "47", g_Language.TranslateMessage("blue tricolor"), g_Language.TranslateMessage("direct color separation; contrast effects in commercial photography"), RGB(43, 75, 220)
    addWratten "47A", g_Language.TranslateMessage("light blue"), g_Language.TranslateMessage("enhance blue and purple objects; useful for fluorescent dyes"), RGB(0, 15, 150)
    addWratten "47B", g_Language.TranslateMessage("deep blue tricolor"), g_Language.TranslateMessage("color separation; calibration using SMPTE color bars"), RGB(0, 0, 120)
    addWratten "56", g_Language.TranslateMessage("very light green"), g_Language.TranslateMessage("darkens sky, improves flesh tones"), RGB(132, 206, 35)
    addWratten "58", g_Language.TranslateMessage("green tricolor"), g_Language.TranslateMessage("used for color separation; improves definition of foliage"), RGB(40, 110, 5)
    addWratten "61", g_Language.TranslateMessage("deep green tricolor"), g_Language.TranslateMessage("used for color separation, tungsten tricolor projection"), RGB(40, 70, 10)
    addWratten "70", g_Language.TranslateMessage("dark red"), g_Language.TranslateMessage("infrared photography longpass filter blocking"), RGB(62, 0, 0)
    addWratten "80A", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("cooling filter, 3200K to 5500K; converts indoor lighting to sunlight"), RGB(50, 100, 230)
    addWratten "80B", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("variant of 80A, 3400K to 5500K"), RGB(70, 120, 230)
    addWratten "80C", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("variant of 80A, 3800K to 5500K"), RGB(90, 140, 235)
    addWratten "80D", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("variant of 80A, 4200K to 5500K"), RGB(110, 160, 240)
    addWratten "81A", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter (lowers color temperature), 3400 K to 3200 K"), RGB(247, 240, 220)
    addWratten "81B", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter; slightly stronger than 81A"), RGB(242, 232, 205)
    addWratten "81C", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter; slightly stronger than 81B"), RGB(230, 220, 200)
    addWratten "81D", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter; slightly stronger than 81C"), RGB(235, 220, 190)
    addWratten "81EF", g_Language.TranslateMessage("pale orange"), g_Language.TranslateMessage("warming filter; slightly stronger than 81D"), RGB(215, 185, 150)
    addWratten "82", g_Language.TranslateMessage("blue"), g_Language.TranslateMessage("cooling filter; raises color temperature 100K"), RGB(150, 205, 240)
    addWratten "82A", g_Language.TranslateMessage("pale blue"), g_Language.TranslateMessage("cooling filter; opposite of 81A"), RGB(205, 225, 235)
    addWratten "82B", g_Language.TranslateMessage("pale blue"), g_Language.TranslateMessage("cooling filter; opposite of 81B"), RGB(155, 190, 220)
    addWratten "82C", g_Language.TranslateMessage("pale blue"), g_Language.TranslateMessage("cooling filter; opposite of 81C"), RGB(120, 155, 190)
    addWratten "85", g_Language.TranslateMessage("amber"), g_Language.TranslateMessage("warming filter, 5500K to 3400K; converts sunlight to incandescent"), RGB(250, 155, 115)
    addWratten "85B", g_Language.TranslateMessage("amber"), g_Language.TranslateMessage("warming filter, 5500K to 3200K; opposite of 80A"), RGB(250, 125, 95)
    addWratten "85C", g_Language.TranslateMessage("amber"), g_Language.TranslateMessage("warming filter, 5500K to 3800K; opposite of 80C"), RGB(250, 155, 115)
    addWratten "90", g_Language.TranslateMessage("dark gray amber"), g_Language.TranslateMessage("remove color before photographing; rarely used for actual photos"), RGB(100, 85, 20)
    addWratten "96", g_Language.TranslateMessage("neutral gray"), g_Language.TranslateMessage("neutral density filter; equally blocks all light frequencies"), RGB(100, 100, 100)
    
    'Determine if the vertical scrollbar needs to be visible or not
    Dim maxMDSize As Long
    maxMDSize = FixDPIFloat(BLOCKHEIGHT) * numOfFilters - 1
    
    vsFilter.Value = 0
    If maxMDSize < picBuffer.ScaleHeight Then
        vsFilter.Visible = False
    Else
        vsFilter.Visible = True
        vsFilter.Max = maxMDSize - picBuffer.ScaleHeight
    End If
    
    vsFilter.Height = picBuffer.Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
      
    'Unload the mouse tracker
    Set cMouseEvents = Nothing
    ReleaseFormTheming Me
        
End Sub

'Given mouse coordinates over the buffer picture box, return the filter at that location
Private Function getFilterAtPosition(ByVal x As Long, ByVal y As Long) As Long
    
    Dim vOffset As Long
    vOffset = vsFilter.Value
    
    getFilterAtPosition = (y + vOffset) \ FixDPI(BLOCKHEIGHT)
    
End Function

Private Sub sltDensity_Change()
    updatePreview
End Sub

'Render a new preview
Private Sub updatePreview()
    
    'Sync the density slider to match the currently selected color
    If sltDensity.GradientColorRight <> fArray(curFilter).RGBColor Then sltDensity.GradientColorRight = fArray(curFilter).RGBColor
    
    'Render the preview
    If cmdBar.previewsAllowed Then ApplyPhotoFilter fArray(curFilter).RGBColor, sltDensity.Value, True, True, fxPreview
    
End Sub

'Cast an image with a new temperature value
' Input: desired temperature, whether to preserve luminance or not, and a blend ratio between 1 and 100
Public Sub ApplyPhotoFilter(ByVal filterColor As Long, ByVal filterDensity As Double, Optional ByVal preserveLuminance As Boolean = True, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying photo filter..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
    Dim originalLuminance As Double
    Dim tmpR As Long, tmpG As Long, tmpB As Long
            
    'Extract the RGB values from the color we were passed
    tmpR = ExtractR(filterColor)
    tmpG = ExtractG(filterColor)
    tmpB = ExtractB(filterColor)
            
    'Divide tempStrength by 100 to yield a value between 0 and 1
    filterDensity = filterDensity / 100
            
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'If luminance is being preserved, we need to determine the initial luminance value
        originalLuminance = (getLuminance(r, g, b) / 255)
        
        'Blend the original and new RGB values using the specified strength
        r = BlendColors(r, tmpR, filterDensity)
        g = BlendColors(g, tmpG, filterDensity)
        b = BlendColors(b, tmpB, filterDensity)
        
        'If the user wants us to preserve luminance, determine the hue and saturation of the new color, then replace the luminance
        ' value with the original
        If preserveLuminance Then
            tRGBToHSL r, g, b, h, s, l
            tHSLToRGB h, s, originalLuminance, r, g, b
        End If
        
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub vsFilter_Change()
    redrawFilterList
End Sub

Private Sub vsFilter_Scroll()
    redrawFilterList
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

