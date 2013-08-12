VERSION 5.00
Begin VB.Form FormPhotoFilters 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Apply Photo Filter"
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
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
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
      TabIndex        =   8
      Top             =   480
      Width           =   8295
   End
   Begin VB.VScrollBar vsFilter 
      Height          =   4185
      LargeChange     =   32
      Left            =   14280
      TabIndex        =   7
      Top             =   480
      Width           =   330
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   11760
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   13230
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltDensity 
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   5040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   100
      Value           =   30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "available filters:"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   5760
      Width           =   14775
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "density:"
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
      Index           =   0
      Left            =   6810
      TabIndex        =   2
      Top             =   5130
      Width           =   840
   End
End
Attribute VB_Name = "FormPhotoFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Photo Filter Application Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 06/June/13
'Last updated: 07/June/13
'Last update: completed initial build
'
'Traditioanl photo filter simulation tool.  A full discussion of photographic filters and how they work are available
' at this Wikipedia article: http://en.wikipedia.org/wiki/Photographic_filter
'
'This code is very similar to PhotoDemon's "Temperature" algorithm.  The main difference is the way the user
' selects a filter to apply.  The available list of filters is flexible, and I have simply based it off Photoshop's
' photo filter list.
'
'I hope it is abundantly clear that these conversions are all very loose estimations.  Filters work by blocking
' specific wavelengths of light at the moment of photography, so it's impossible to perfectly replicate their behavior
' via code.  All we can do is approximate, so do not expect to get identical results between actual filters and
' post-production tools like PhotoDemon.
'
'Luminosity preservation is assumed.  I could provide a toggle for it, but I see no real benefit to unpreserved use
' of these tools.
'
'The list-box-style interface was custom built for this tool, and I derived it from similar code in the Metadata Browser
' and About dialog.  Please compile for best results; things like mousewheel support and mouseleave tracking require
' subclassing, so they are not enabled in the IDE.
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
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'API for drawing colored rectangles
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

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

'Subclass the window to enable mousewheel support for scrolling the filter view (compiled EXE only)
Dim m_Subclass As cSelfSubHookCallback

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Extra variables for custom list rendering
Dim bufferLayer As pdLayer
Dim m_BufferWidth As Long, m_BufferHeight As Long

'Two font objects; one for names and one for URLs.  (Two are needed because they have different sizes and colors.)
Dim firstFont As pdFont, secondFont As pdFont

'A primary and secondary color for font rendering
Dim primaryColor As Long, secondaryColor As Long

'The currently selected and currently hovered filter entry
Dim curFilter As Long, curFilterHover As Long

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'The scroll bar max and min values are used to check the temperature input for validity
    If sltDensity.IsValid Then
        Me.Visible = False
        Process "Photo filter", , buildParams(PicColor.backColor, sltDensity.Value, True)
        Unload Me
    End If
    
End Sub

'Redraw the current list of filters
Private Sub redrawFilterList()
        
    Dim scrollOffset As Long
    scrollOffset = vsFilter.Value
    
    bufferLayer.createBlank picBuffer.ScaleWidth, picBuffer.ScaleHeight
    
    Dim i As Long
    For i = 0 To numOfFilters - 1
        renderFilterBlock i, 0, i * BLOCKHEIGHT - scrollOffset - 2
    Next i
    
    'Copy the buffer to the main form
    BitBlt picBuffer.hDC, 0, 0, m_BufferWidth, m_BufferHeight, bufferLayer.getLayerDC, 0, 0, vbSrcCopy
    picBuffer.Picture = picBuffer.Image
    picBuffer.Refresh
    
End Sub

'Render an individual "block" filter content (name, description, color)
Private Sub renderFilterBlock(ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + BLOCKHEIGHT) > 0) And (offsetY < m_BufferHeight) Then
    
        offsetY = offsetY + 2
        
        Dim linePadding As Long
        linePadding = 2
    
        Dim mHeight As Single
        Dim tmpRect As RECT
        Dim hBrush As Long
        
        'If this filter has been selected, draw the background with the system's current selection color
        If blockIndex = curFilter Then
        
            SetRect tmpRect, offsetX, offsetY, m_BufferWidth, offsetY + BLOCKHEIGHT
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FillRect bufferLayer.getLayerDC, tmpRect, hBrush
            DeleteObject hBrush
            
            'Also, color the fonts with the matching highlighted text color (otherwise they won't be readable)
            firstFont.setFontColor ConvertSystemColor(vbHighlightText)
            secondFont.setFontColor ConvertSystemColor(vbHighlightText)
        
        Else
            firstFont.setFontColor primaryColor
            secondFont.setFontColor secondaryColor
        End If
        
        'If the current filter is highlighted but not selected, simply render the border with a highlight
        If (blockIndex <> curFilter) And (blockIndex = curFilterHover) Then
            SetRect tmpRect, offsetX, offsetY, m_BufferWidth, offsetY + BLOCKHEIGHT
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FrameRect bufferLayer.getLayerDC, tmpRect, hBrush
            DeleteObject hBrush
        End If
        
        Dim drawString As String
        drawString = fArray(blockIndex).Id & " - " & fArray(blockIndex).Name
        
        'Render a color box for the image
        Dim colorWidth As Long, colorHeight As Long
        colorWidth = 64
        colorHeight = 42
        SetRect tmpRect, offsetX + 4, offsetY + ((BLOCKHEIGHT - colorHeight) \ 2), offsetX + 4 + colorWidth, offsetY + ((BLOCKHEIGHT - colorHeight) \ 2) + colorHeight
        
        hBrush = CreateSolidBrush(fArray(blockIndex).RGBColor)
        FillRect bufferLayer.getLayerDC, tmpRect, hBrush
        DeleteObject hBrush
        hBrush = CreateSolidBrush(RGB(64, 64, 64))
        FrameRect bufferLayer.getLayerDC, tmpRect, hBrush
        DeleteObject hBrush
            
        'Render the Wratten ID and name fields
        firstFont.attachToDC bufferLayer.getLayerDC
        firstFont.fastRenderText colorWidth + 16 + offsetX, offsetY + 4, drawString
                
        'Below that, add the description text
        mHeight = firstFont.getHeightOfString(drawString) + linePadding
        drawString = fArray(blockIndex).Description
        
        secondFont.attachToDC bufferLayer.getLayerDC
        secondFont.fastRenderText colorWidth + 16 + offsetX, offsetY + 4 + mHeight, drawString
        
    End If

End Sub

'When the form is activated (e.g. made visible and receives focus),
Private Sub Form_Activate()
    
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
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    setHandCursorToHwnd picBuffer.hWnd
    
    'Determine if the vertical scrollbar needs to be visible or not
    Dim maxMDSize As Long
    maxMDSize = BLOCKHEIGHT * numOfFilters - 1
    
    vsFilter.Value = 0
    If maxMDSize < picBuffer.ScaleHeight Then
        vsFilter.Visible = False
    Else
        vsFilter.Visible = True
        vsFilter.Max = maxMDSize - picBuffer.ScaleHeight
    End If
    
    vsFilter.Height = picBuffer.Height
    
    'No filters are currently selected or hovered
    curFilter = -1
    curFilterHover = -1
    
    'Display the previewed effect in the neighboring window, then render the list of available filters
    ApplyPhotoFilter RGB(127, 127, 127), 0, True, True, fxPreview
    redrawFilterList
    
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

    'If the program is compiled, enable mousewheel scrolling for the filter box
    If g_IsProgramCompiled Then
    
        'Request mouse tracking for the buffer picture box
        requestMouseTracking picBuffer.hWnd
        
        'Add support for scrolling with the mouse wheel (e.g. initialize the relevant subclassing object)
        Set m_Subclass = New cSelfSubHookCallback
        
        'Add mousewheel messages to the subclassing handler (compiled only)
        If m_Subclass.ssc_Subclass(Me.hWnd, , 1, Me) Then m_Subclass.ssc_AddMsg Me.hWnd, MSG_BEFORE, WM_MOUSEWHEEL
        If m_Subclass.ssc_Subclass(picBuffer.hWnd, picBuffer.hWnd, 1, Me) Then
            m_Subclass.ssc_AddMsg picBuffer.hWnd, MSG_BEFORE, WM_MOUSEWHEEL
            m_Subclass.ssc_AddMsg picBuffer.hWnd, MSG_BEFORE, WM_MOUSELEAVE 'Mouse leaves the window (used to clear the actively hovered block)
        End If
        
    End If
    
    'Create a background buffer the same size as the buffer picture box
    Set bufferLayer = New pdLayer
    bufferLayer.createBlank picBuffer.ScaleWidth, picBuffer.ScaleHeight
    
    'Initialize a few other variables for speed reasons
    m_BufferWidth = picBuffer.ScaleWidth
    m_BufferHeight = picBuffer.ScaleHeight
    
    'Initialize a custom font objects for names
    primaryColor = RGB(64, 64, 64)
    Set firstFont = New pdFont
    firstFont.setFontColor primaryColor
    firstFont.setFontBold True
    firstFont.setFontSize 12
    firstFont.createFontObject
    firstFont.setTextAlignment vbLeftJustify
    
    '...and a second custom font object for URLs
    secondaryColor = RGB(92, 92, 92)
    Set secondFont = New pdFont
    secondFont.setFontColor secondaryColor
    secondFont.setFontBold False
    secondFont.setFontSize 10
    secondFont.createFontObject
    secondFont.setTextAlignment vbLeftJustify
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
      
    If g_IsProgramCompiled Then
    
        'Release the subclassing object responsible for mouse wheel support
        m_Subclass.ssc_Terminate
        Set m_Subclass = Nothing
        
        'Stop requesting mouse tracking
        requestMouseTracking picBuffer.hWnd, True
        
    End If
    
    ReleaseFormTheming Me
        
End Sub

Private Sub picBuffer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    curFilter = getFilterAtPosition(x, y)
    PicColor.backColor = fArray(curFilter).RGBColor
    redrawFilterList
    updatePreview
    
End Sub

Private Sub picBuffer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Ask Windows to track the mouse relative to this picture box
    If g_IsProgramCompiled Then requestMouseTracking picBuffer.hWnd
    
    curFilterHover = getFilterAtPosition(x, y)
    redrawFilterList
    
End Sub

'Given mouse coordinates over the buffer picture box, return the filter at that location
Private Function getFilterAtPosition(ByVal x As Long, ByVal y As Long) As Long
    
    Dim vOffset As Long
    vOffset = vsFilter.Value
    
    getFilterAtPosition = (y + vOffset) \ BLOCKHEIGHT
    
End Function

Private Sub sltDensity_Change()
    updatePreview
End Sub

'Render a new preview
Private Sub updatePreview()
    ApplyPhotoFilter PicColor.backColor, sltDensity.Value, True, True, fxPreview
End Sub

'Cast an image with a new temperature value
' Input: desired temperature, whether to preserve luminance or not, and a blend ratio between 1 and 100
Public Sub ApplyPhotoFilter(ByVal filterColor As Long, ByVal filterDensity As Double, Optional ByVal preserveLuminance As Boolean = True, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Applying photo filter..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
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

'This custom routine, combined with careful subclassing, allows us to handle mouse wheel events for this form.
Private Sub MouseWheel(ByVal MouseKeys As Long, ByVal mRotation As Long, ByVal xPos As Long, ByVal yPos As Long)
    
    'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
    If vsFilter.Visible Then
  
        If mRotation < 0 Then
            
            If vsFilter.Value + vsFilter.LargeChange > vsFilter.Max Then
                vsFilter.Value = vsFilter.Max
            Else
                vsFilter.Value = vsFilter.Value + vsFilter.LargeChange
            End If
            
            redrawFilterList
        
        ElseIf mRotation > 0 Then
            
            If vsFilter.Value - vsFilter.LargeChange < vsFilter.Min Then
                vsFilter.Value = vsFilter.Min
            Else
                vsFilter.Value = vsFilter.Value - vsFilter.LargeChange
            End If
            
            redrawFilterList
            
        End If
        
    End If
    
End Sub

'This routine MUST BE KEPT as the final routine for this form. Its ordinal position determines its ability to subclass properly.
' Subclassing is required to enable mousewheel support and other mouse events (e.g. the mouse leaving the window).
Private Sub myWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
        
    Dim MouseKeys As Long
    Dim mRotation As Long
    Dim xPos As Long
    Dim yPos As Long
    
    'Only handle scroll events if the message relates to this form
    Select Case uMsg
  
        Case WM_MOUSEWHEEL
    
            MouseKeys = wParam And 65535
            mRotation = wParam / 65536
            xPos = lParam And 65535
            yPos = lParam / 65536
            
            MouseWheel MouseKeys, mRotation, xPos, yPos
            
    End Select
    
    'If the mouse leaves the filter box, remove any hovered entry
    If (lParamUser = picBuffer.hWnd) And (uMsg = WM_MOUSELEAVE) Then
        curFilterHover = -1
        redrawFilterList
    End If
    
End Sub
