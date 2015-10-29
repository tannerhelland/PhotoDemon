VERSION 5.00
Begin VB.Form FormMetadata 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Browse image metadata"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12015
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
   ScaleHeight     =   523
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   13
      Top             =   7095
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.pdButton cmdTechnicalReport 
      Height          =   735
      Left            =   7440
      TabIndex        =   12
      Top             =   4380
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   1296
      Caption         =   "Generate full metadata report (HTML)..."
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   6615
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   11
      Top             =   1740
      Width           =   255
   End
   Begin PhotoDemon.buttonStrip btsGroup 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   1085
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00404040&
      Height          =   5115
      Left            =   120
      ScaleHeight     =   339
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   0
      Top             =   1740
      Width           =   6495
   End
   Begin PhotoDemon.buttonStrip btsTechnical 
      Height          =   495
      Index           =   0
      Left            =   7440
      TabIndex        =   6
      Top             =   2160
      Width           =   4410
      _ExtentX        =   20690
      _ExtentY        =   1085
   End
   Begin PhotoDemon.buttonStrip btsTechnical 
      Height          =   495
      Index           =   1
      Left            =   7440
      TabIndex        =   7
      Top             =   3240
      Width           =   4410
      _ExtentX        =   6720
      _ExtentY        =   873
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "advanced"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   5
      Left            =   7440
      TabIndex        =   10
      Top             =   3960
      Width           =   945
   End
   Begin VB.Label lblExifTool 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7320
      TabIndex        =   9
      Top             =   6120
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   476
      X2              =   476
      Y1              =   88
      Y2              =   464
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "metadata options"
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
      Index           =   4
      Left            =   7320
      TabIndex        =   8
      Top             =   1320
      Width           =   1830
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tag values"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   3
      Left            =   7440
      TabIndex        =   5
      Top             =   2820
      Width           =   1005
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tag names"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   2
      Left            =   7440
      TabIndex        =   4
      Top             =   1740
      Width           =   1050
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tags in this category"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2130
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "metadata groups in this image"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3225
   End
End
Attribute VB_Name = "FormMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Metadata Browser
'Copyright 2013-2015 by Tanner Helland
'Created: 27/May/13
'Last updated: 25/October/14
'Last update: clean up render code, improve mousewheel behavior, use button strip for some interface elements
'
'As of version 6.0, PhotoDemon now provides support for loading and saving image metadata.  What is metadata, you ask?
' See http://en.wikipedia.org/wiki/Metadata#Photographs for more details.
'
'This dialog interacts heavily with the pdMetadata class to present users with a relatively simple interface for
' perusing (and eventually, editing - didn't make the cut for 6.0 or 6.2 but maybe 6.4??) an image's metadata.
'
'Designing this dialog was quite difficult as it is impossible to predict what metadata types and entries might exist in
' an image file, so I've opted for the most flexible system I can.  No assumptions are made about present categories or
' tag counts, so any type or amount of metadata should theoretically be viewable.
'
'Categories are displayed on the left, and selecting a category repopulates the fields on the right.  Future updates
' could include the ability to add or remove individual tags...
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Type mdCategory
    Name As String
    Count As Long
End Type

Private mdCategories() As mdCategory
Private numOfCategories As Long
Private highestCategoryCount As Long

'This array will hold all tags currently in storage, but sorted into categories
Private allTags() As mdItem
Private curTagCount() As Long

'Height of each metadata content block
Private Const BLOCKHEIGHT As Long = 54

'An outside class provides access to mousewheel events for scrolling the filter view
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'The back buffer onto which the metadata list is rendered
Private m_BackBuffer As pdDIB

'Font objects for rendering
Private m_TitleFont As pdFont, m_DescriptionFont As pdFont

'Additional rendering values
Private m_SeparatorColor As OLE_COLOR

'API scrollbar allows for larger scroll values
Private WithEvents vsMetadata As pdScrollAPI
Attribute vsMetadata.VB_VarHelpID = -1

'When a new metadata category is selected, redraw all the metadata text currently on screen
Private Sub btsGroup_Click(ByVal buttonIndex As Long)
    
    Dim curCategory As Long
    curCategory = buttonIndex
    
    If mdCategories(curCategory).Count = 1 Then
        lblTitle(1).Caption = g_Language.TranslateMessage("1 tag in this category:")
    Else
        lblTitle(1).Caption = g_Language.TranslateMessage("%1 tags in this category:", mdCategories(curCategory).Count)
    End If
    
    'First, determine if the vertical scrollbar needs to be visible or not
    Dim maxMDSize As Long
    maxMDSize = FixDPIFloat(BLOCKHEIGHT) * mdCategories(curCategory).Count
    
    vsMetadata.Value = 0
    If maxMDSize < picBuffer.Height Then
        picScroll.Visible = False
    Else
        picScroll.Visible = True
        vsMetadata.Max = maxMDSize - picBuffer.Height
    End If
    
    redrawMetadataList
        
End Sub

Private Sub btsGroup_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    cMouseEvents_MouseWheelVertical Button, Shift, x, y, scrollAmount
End Sub

Private Sub btsTechnical_Click(Index As Integer, ByVal buttonIndex As Long)
    redrawMetadataList
End Sub

Private Sub cmdTechnicalReport_Click()
    Plugin_ExifTool_Interface.createTechnicalMetadataReport pdImages(g_CurrentImage)
End Sub

Private Sub cMouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

    'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
    If picScroll.Visible Then
  
        If scrollAmount < 0 Then
            
            If vsMetadata.Value + vsMetadata.LargeChange > vsMetadata.Max Then
                vsMetadata.Value = vsMetadata.Max
            Else
                vsMetadata.Value = vsMetadata.Value + vsMetadata.LargeChange
            End If
            
            redrawMetadataList
        
        ElseIf scrollAmount > 0 Then
            
            If vsMetadata.Value - vsMetadata.LargeChange < vsMetadata.Min Then
                vsMetadata.Value = vsMetadata.Min
            Else
                vsMetadata.Value = vsMetadata.Value - vsMetadata.LargeChange
            End If
            
            redrawMetadataList
            
        End If
        
    End If

End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
End Sub

'LOAD dialog
Private Sub Form_Load()
    
    'Create an API scroll bar for the main metadata window
    Set vsMetadata = New pdScrollAPI
    vsMetadata.initializeScrollBarWindow picScroll.hWnd, False, 0, 100, 0, 1, 32
    
    'Note that this form will be interacting heavily with the current image's metadata container.
    
    'Enable mousewheel scrolling for the metadata box
    Set cMouseEvents = New pdInputMouse
    cMouseEvents.addInputTracker picBuffer.hWnd, True, , , True
    cMouseEvents.addInputTracker Me.hWnd
    cMouseEvents.setSystemCursor IDC_ARROW
    
    'Prepare all rendering objects
    Set m_BackBuffer = New pdDIB
    m_BackBuffer.createBlank picBuffer.ScaleWidth, picBuffer.ScaleHeight, 24
    
    m_SeparatorColor = vbActiveTitleBar
    
    Set m_TitleFont = New pdFont
    m_TitleFont.SetFontColor RGB(64, 64, 64)
    m_TitleFont.SetFontBold True
    m_TitleFont.SetFontSize 10
    m_TitleFont.CreateFontObject
    m_TitleFont.SetTextAlignment vbLeftJustify
    
    Set m_DescriptionFont = New pdFont
    m_DescriptionFont.SetFontColor RGB(92, 92, 92)
    m_DescriptionFont.SetFontBold False
    m_DescriptionFont.SetFontSize 10
    m_DescriptionFont.CreateFontObject
    m_DescriptionFont.SetTextAlignment vbLeftJustify
    
    'Make the invisible buffer's font match the rest of PD
    picBuffer.fontName = g_InterfaceFont
        
    'Initialize the category array
    ReDim mdCategories(0) As mdCategory
    numOfCategories = 0
    
    'Start by tallying up information on the various metadata types within this image
    Dim chkGroup As String
    Dim curMetadata As mdItem
    Dim categoryFound As Boolean
    
    Dim i As Long, j As Long
    For i = 0 To pdImages(g_CurrentImage).imgMetadata.getMetadataCount - 1
    
        categoryFound = False
    
        'Retrieve the next metadata entry
        curMetadata = pdImages(g_CurrentImage).imgMetadata.getMetadataEntry(i)
        chkGroup = curMetadata.Group
        
        'Search the current list of known categories for this metadata object's category
        For j = 0 To numOfCategories
            If StrComp(mdCategories(j).Name, chkGroup) = 0 Then
                categoryFound = True
                mdCategories(j).Count = mdCategories(j).Count + 1
                Exit For
            End If
        Next j
        
        'If no matching category was found, create a new category entry
        If Not categoryFound Then
            mdCategories(numOfCategories).Name = chkGroup
            mdCategories(numOfCategories).Count = 1
            numOfCategories = numOfCategories + 1
            ReDim Preserve mdCategories(0 To numOfCategories) As mdCategory
        End If
    
    Next i
    
    'We can now populate the left-side list box with the categories we found.  While doing this, find
    ' the category with the highest tag count.
    highestCategoryCount = 0
    
    'Prior to adding category names, set a relevant button strip font according to the number of metadata groups.
    ' If an image has a ton of groups (10+ is not unheard of), reduce font size.
    If numOfCategories > 5 Then
        btsGroup.FontSize = 10
    Else
        btsGroup.FontSize = 12
    End If
    
    For i = 0 To numOfCategories - 1
        btsGroup.AddItem mdCategories(i).Name, i
        If mdCategories(i).Count > highestCategoryCount Then highestCategoryCount = mdCategories(i).Count
    Next i
    
    'We can now build a 2D array that contains all tags, sorted by category.  Why not do this above?  Because
    ' it's computationally expensive to constantly redim arrays in VB, and this technique allows us to redim
    ' the main tag array only once, after all values have been tallied.
    ReDim allTags(0 To numOfCategories - 1, 0 To highestCategoryCount - 1) As mdItem
    ReDim curTagCount(0 To numOfCategories - 1) As Long
    
    For i = 0 To pdImages(g_CurrentImage).imgMetadata.getMetadataCount - 1
        
        'As above, retrieve the next metadata entry
        curMetadata = pdImages(g_CurrentImage).imgMetadata.getMetadataEntry(i)
        chkGroup = curMetadata.Group
        
        'Find the matching group in the Group array, then insert this tag into place
        For j = 0 To numOfCategories - 1
            If StrComp(mdCategories(j).Name, chkGroup) = 0 Then
            
                allTags(j, curTagCount(j)) = curMetadata
                curTagCount(j) = curTagCount(j) + 1
                Exit For
                
            End If
        Next j
        
    Next i
    
    'Populate the simple/technical switches at the bottom
    btsTechnical(0).AddItem "simple", 0
    btsTechnical(0).AddItem "technical", 1
    btsTechnical(0).ListIndex = 0
    
    btsTechnical(1).AddItem "simple", 0
    btsTechnical(1).AddItem "technical", 1
    btsTechnical(1).ListIndex = 0
    
    'Select the first group by default
    btsGroup.ListIndex = 0
    btsGroup_Click 0
    
    'Technical metadata reports are only available for images that exist on disk
    If Len(pdImages(g_CurrentImage).locationOnDisk) <> 0 Then
        cmdTechnicalReport.Visible = True
    Else
        cmdTechnicalReport.Visible = False
    End If
    
    'Give ExifTool credit for its amazing work!
    lblExifTool.Caption = g_Language.TranslateMessage("All metadata information is supplied by the ExifTool plugin.  You can learn more about ExifTool at http://www.sno.phy.queensu.ca/~phil/exiftool/")
    
End Sub

'UNLOAD form
Private Sub Form_Unload(Cancel As Integer)
    
    'Unload the mouse tracker
    Set cMouseEvents = Nothing
    ReleaseFormTheming Me

End Sub

'Redraw the full metadata list
Private Sub redrawMetadataList()

    Dim curCategory As Long
    curCategory = btsGroup.ListIndex
        
    Dim scrollOffset As Long
    scrollOffset = vsMetadata.Value
    
    'Clear the back buffer
    GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, vbWhite, 255
    
    'Render each block in turn
    Dim i As Long
    For i = 0 To mdCategories(curCategory).Count - 1
        renderMDBlock curCategory, i, FixDPI(8), FixDPI(i * BLOCKHEIGHT) - scrollOffset - FixDPI(2)
    Next i
    
    'Copy the buffer to the target picture box
    BitBlt picBuffer.hDC, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
    picBuffer.Picture = picBuffer.Image
    
End Sub

'Render the given metadata index onto the background picture box at the specified offset
Private Sub renderMDBlock(ByVal blockCategory As Long, ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + BLOCKHEIGHT) > 0) And (offsetY < m_BackBuffer.getDIBHeight) Then
        
        offsetY = offsetY + FixDPI(9)
        
        Dim thisTag As mdItem
        thisTag = allTags(blockCategory, blockIndex)
        
        Dim linePadding As Long
        linePadding = FixDPI(4)
    
        Dim mHeight As Single
        Dim drawString As String, numericalPrefix As String
        
        numericalPrefix = CStr(blockIndex + 1) & " - "
        
        If (btsTechnical(0).ListIndex = 0) Then
            drawString = thisTag.Description
        Else
            drawString = thisTag.FullGroupAndName
        End If
        
        'Notify the user of text we were unable to manually convert
        If thisTag.isValueBase64 Then
            drawString = drawString & " " & g_Language.TranslateMessage("(encoding unknown)")
        End If
    
        'Start with the simplest field: the tag title (readable form)
        m_TitleFont.AttachToDC m_BackBuffer.getDIBDC
        m_TitleFont.FastRenderText offsetX + 0, offsetY + 0, numericalPrefix & drawString
                
        'Below the tag title, add the human-friendly description
        mHeight = m_TitleFont.GetHeightOfString(drawString) + linePadding
        
        If (btsTechnical(1).ListIndex = 0) Then
            drawString = thisTag.Value
        Else
            If Len(thisTag.ActualValue) <> 0 Then
                drawString = thisTag.ActualValue
            Else
                drawString = thisTag.Value
            End If
        End If
        
        m_TitleFont.ReleaseFromDC
        m_DescriptionFont.AttachToDC m_BackBuffer.getDIBDC
        m_DescriptionFont.FastRenderTextWithClipping offsetX + m_TitleFont.GetWidthOfString(numericalPrefix), offsetY + mHeight, m_BackBuffer.getDIBWidth - offsetX - FixDPI(17), m_DescriptionFont.GetHeightOfString(drawString), drawString
        m_DescriptionFont.ReleaseFromDC
        
        'Draw a divider line near the bottom of the metadata block
        Dim lineY As Long
        'If blockIndex < mdCategories(blockCategory).Count - 1 Then
            lineY = offsetY + FixDPI(BLOCKHEIGHT - 7)
            GDI_Plus.GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, FixDPI(4), lineY, m_BackBuffer.getDIBWidth - FixDPI(4), lineY, m_SeparatorColor
        'End If
        
    End If

End Sub

'When the scrollbar is moved, redraw the metadata list
Private Sub vsMetadata_Scroll()
    redrawMetadataList
End Sub
