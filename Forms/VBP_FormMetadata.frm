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
   Begin PhotoDemon.buttonStrip btsGroup 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   1085
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
   Begin VB.VScrollBar vsMetadata 
      Height          =   5115
      LargeChange     =   32
      Left            =   6615
      TabIndex        =   4
      Top             =   1740
      Width           =   330
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
      TabIndex        =   3
      Top             =   1740
      Width           =   6495
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10500
      TabIndex        =   1
      Top             =   7230
      Width           =   1365
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   7230
      Width           =   1365
   End
   Begin PhotoDemon.buttonStrip btsTechnical 
      Height          =   495
      Index           =   0
      Left            =   7440
      TabIndex        =   10
      Top             =   2160
      Width           =   3810
      _ExtentX        =   20690
      _ExtentY        =   1085
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
   Begin PhotoDemon.buttonStrip btsTechnical 
      Height          =   495
      Index           =   1
      Left            =   7440
      TabIndex        =   11
      Top             =   3240
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   873
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
      Caption         =   "metadata options:"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   1920
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tag values:"
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
      TabIndex        =   9
      Top             =   2820
      Width           =   1080
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tag names:"
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
      TabIndex        =   8
      Top             =   1740
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tags in this category:"
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
      TabIndex        =   7
      Top             =   1320
      Width           =   2220
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "metadata groups in this image:"
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
      TabIndex        =   6
      Top             =   120
      Width           =   3315
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   7080
      Width           =   12255
   End
End
Attribute VB_Name = "FormMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Metadata Browser
'Copyright ©2013-2014 by Tanner Helland
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

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

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
    maxMDSize = fixDPIFloat(BLOCKHEIGHT) * mdCategories(curCategory).Count
    
    vsMetadata.Value = 0
    If maxMDSize < picBuffer.Height Then
        vsMetadata.Visible = False
    Else
        vsMetadata.Visible = True
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

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    Unload Me
End Sub

Private Sub cMouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

    'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
    If vsMetadata.Visible Then
  
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
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

'LOAD dialog
Private Sub Form_Load()

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
    m_TitleFont.setFontColor RGB(64, 64, 64)
    m_TitleFont.setFontBold True
    m_TitleFont.setFontSize 10
    m_TitleFont.createFontObject
    m_TitleFont.setTextAlignment vbLeftJustify
    
    Set m_DescriptionFont = New pdFont
    m_DescriptionFont.setFontColor RGB(92, 92, 92)
    m_DescriptionFont.setFontBold False
    m_DescriptionFont.setFontSize 10
    m_DescriptionFont.createFontObject
    m_DescriptionFont.setTextAlignment vbLeftJustify
    
    'Make the invisible buffer's font match the rest of PD
    picBuffer.FontName = g_InterfaceFont
        
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
        btsGroup.Font.Size = 10
    Else
        btsGroup.Font.Size = 12
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
        renderMDBlock curCategory, i, fixDPI(8), fixDPI(i * BLOCKHEIGHT) - scrollOffset - fixDPI(2)
    Next i
    
    'Copy the buffer to the target picture box
    BitBlt picBuffer.hDC, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
    picBuffer.Picture = picBuffer.Image
    
End Sub

'Render the given metadata index onto the background picture box at the specified offset
Private Sub renderMDBlock(ByVal blockCategory As Long, ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + BLOCKHEIGHT) > 0) And (offsetY < m_BackBuffer.getDIBHeight) Then
        
        offsetY = offsetY + fixDPI(9)
        
        Dim thisTag As mdItem
        thisTag = allTags(blockCategory, blockIndex)
        
        Dim linePadding As Long
        linePadding = fixDPI(4)
    
        Dim mHeight As Single
        Dim drawString As String, numericalPrefix As String
        
        numericalPrefix = CStr(blockIndex + 1) & " - "
        
        If (btsTechnical(0).ListIndex = 0) Then
            drawString = thisTag.Description
        Else
            drawString = thisTag.FullGroupAndName
        End If
    
        'Start with the simplest field: the tag title (readable form)
        m_TitleFont.attachToDC m_BackBuffer.getDIBDC
        m_TitleFont.fastRenderText offsetX + 0, offsetY + 0, numericalPrefix & drawString
                
        'Below the tag title, add the human-friendly description
        mHeight = m_TitleFont.getHeightOfString(drawString) + linePadding
        
        If (btsTechnical(1).ListIndex = 0) Then
            drawString = thisTag.Value
        Else
            If Len(thisTag.ActualValue) > 0 Then
                drawString = thisTag.ActualValue
            Else
                drawString = thisTag.Value
            End If
        End If
        
        m_DescriptionFont.attachToDC m_BackBuffer.getDIBDC
        m_DescriptionFont.fastRenderTextWithClipping offsetX + m_TitleFont.getWidthOfString(numericalPrefix), offsetY + mHeight, m_BackBuffer.getDIBWidth - offsetX - fixDPI(17), m_DescriptionFont.getHeightOfString(drawString), drawString
        
        'Draw a divider line near the bottom of the metadata block
        Dim lineY As Long
        If blockIndex < mdCategories(blockCategory).Count - 1 Then
            lineY = offsetY + fixDPI(BLOCKHEIGHT - 7)
            GDI_Plus.GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, fixDPI(4), lineY, m_BackBuffer.getDIBWidth - fixDPI(4), lineY, m_SeparatorColor
        End If
        
    End If

End Sub

'When the scrollbar is moved, redraw the metadata list
Private Sub vsMetadata_Change()
    redrawMetadataList
End Sub

Private Sub vsMetadata_Scroll()
    redrawMetadataList
End Sub
