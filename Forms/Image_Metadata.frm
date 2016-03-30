VERSION 5.00
Begin VB.Form FormMetadata 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Browse image metadata"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14070
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
   ScaleHeight     =   623
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   938
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdHyperlink hypExiftool 
      Height          =   255
      Left            =   240
      Top             =   8160
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "click to visit the ExifTool homepage"
      FontSize        =   9
      URL             =   "http://www.sno.phy.queensu.ca/~phil/exiftool/"
   End
   Begin PhotoDemon.pdButtonStripVertical btsEditPanel 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3201
      Caption         =   "metadata tools"
   End
   Begin PhotoDemon.pdListBox lstGroup 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _ExtentX        =   5953
      _ExtentY        =   9975
      Caption         =   "metadata groups in this image"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   8595
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdLabel lblExifTool 
      Height          =   855
      Left            =   240
      Top             =   7200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1508
      Alignment       =   2
      Caption         =   ""
      FontSize        =   9
      ForeColor       =   -2147483640
      Layout          =   1
   End
   Begin PhotoDemon.pdListBoxOD lstMetadata 
      Height          =   4455
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   10095
      _ExtentX        =   12091
      _ExtentY        =   9975
      Caption         =   "tags in this category"
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   3855
      Index           =   0
      Left            =   3840
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   673
      TabIndex        =   4
      Top             =   4680
      Width           =   10095
      Begin PhotoDemon.pdCheckBox chkFriendlyValue 
         Height          =   345
         Left            =   4080
         TabIndex        =   10
         Top             =   1560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   609
         Caption         =   "use formal values only"
         Value           =   0
      End
      Begin PhotoDemon.pdTextBox txtValue 
         Height          =   1140
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2011
         Multiline       =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblTagName 
         Height          =   300
         Left            =   480
         Top             =   360
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         Caption         =   ""
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   300
         Index           =   0
         Left            =   240
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         Caption         =   "tag name"
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTagID 
         Height          =   300
         Left            =   480
         Top             =   1200
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         Caption         =   ""
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   300
         Index           =   1
         Left            =   240
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         Caption         =   "tag ID (optional)"
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTable 
         Height          =   300
         Left            =   480
         Top             =   2040
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         Caption         =   ""
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   300
         Index           =   2
         Left            =   240
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         Caption         =   "parent table"
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblGroup 
         Height          =   300
         Left            =   480
         Top             =   2880
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         Caption         =   ""
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   300
         Index           =   3
         Left            =   240
         Top             =   2520
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         Caption         =   "subgroup"
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   300
         Index           =   5
         Left            =   3960
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         Caption         =   "tag value"
         FontSize        =   12
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   3855
      Index           =   1
      Left            =   3840
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   673
      TabIndex        =   5
      Top             =   4680
      Width           =   10095
      Begin PhotoDemon.pdButtonStrip btsTechnical 
         Height          =   975
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4530
         _ExtentX        =   7779
         _ExtentY        =   1720
         Caption         =   "tag names"
      End
      Begin PhotoDemon.pdButtonStrip btsTechnical 
         Height          =   975
         Index           =   1
         Left            =   5040
         TabIndex        =   7
         Top             =   360
         Width           =   4530
         _ExtentX        =   7779
         _ExtentY        =   1720
         Caption         =   "tag values"
      End
      Begin PhotoDemon.pdButton cmdTechnicalReport 
         Height          =   555
         Left            =   360
         TabIndex        =   8
         Top             =   1950
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   979
         Caption         =   "Generate full metadata report (HTML)..."
      End
      Begin PhotoDemon.pdLabel lblTechnicalReport 
         Height          =   270
         Left            =   240
         Top             =   1560
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   476
         Caption         =   "advanced tools"
         FontSize        =   11
         ForeColor       =   4210752
      End
   End
End
Attribute VB_Name = "FormMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Metadata Browser
'Copyright 2013-2016 by Tanner Helland
'Created: 27/May/13
'Last updated: 27/March/16
'Last update: start overhaul required for metadata editing support
'
'As of version 6.0, PhotoDemon now provides support for loading and saving image metadata.  What is metadata, you ask?
' See http://en.wikipedia.org/wiki/Metadata#Photographs for more details.
'
'This dialog interacts heavily with the pdMetadata class to present users with a relatively simple interface for
' perusing (and eventually, editing) an image's metadata.
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
Private allTags() As PDMetadataItem
Private curTagCount() As Long

'Height of each metadata content block
Private Const BLOCKHEIGHT As Long = 54

'Font objects for rendering
Private m_TitleFont As pdFont, m_DescriptionFont As pdFont

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDMETADATA_COLOR_LIST
    [_First] = 0
    PDMD_TitleSelected = 0
    PDMD_TitleUnselected = 1
    PDMD_DescriptionSelected = 2
    PDMD_DescriptionUnselected = 3
    [_Last] = 3
    [_Count] = 4
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Private Sub btsEditPanel_Click(ByVal buttonIndex As Long)
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).Visible = CBool(i = buttonIndex)
    Next i
End Sub

Private Sub btsTechnical_Click(Index As Integer, ByVal buttonIndex As Long)
    Dim vScrollValue As Long, lstListIndex As Long
    vScrollValue = lstMetadata.GetScrollValue
    lstListIndex = lstMetadata.ListIndex
    UpdateMetadataList
    lstMetadata.SetScrollValue vScrollValue
    lstMetadata.ListIndex = lstListIndex
End Sub

Private Sub chkFriendlyValue_Click()
    UpdateTagView
End Sub

Private Sub cmdTechnicalReport_Click()
    ExifTool.CreateTechnicalMetadataReport pdImages(g_CurrentImage)
End Sub

Private Sub Form_Load()
    
    lstMetadata.ListItemHeight = FixDPI(BLOCKHEIGHT)
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDMETADATA_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDMetadataList", colorCount
    UpdateColorList
    
    Set m_TitleFont = New pdFont
    m_TitleFont.SetFontBold True
    m_TitleFont.SetFontSize 10
    m_TitleFont.CreateFontObject
    m_TitleFont.SetTextAlignment vbLeftJustify
    
    Set m_DescriptionFont = New pdFont
    m_DescriptionFont.SetFontBold False
    m_DescriptionFont.SetFontSize 10
    m_DescriptionFont.CreateFontObject
    m_DescriptionFont.SetTextAlignment vbLeftJustify
        
    'Initialize the category array
    ReDim mdCategories(0) As mdCategory
    numOfCategories = 0
    
    'Start by tallying up information on the various metadata types within this image
    Dim chkGroup As String
    Dim curMetadata As PDMetadataItem
    Dim categoryFound As Boolean
    
    Dim i As Long, j As Long
    For i = 0 To pdImages(g_CurrentImage).imgMetadata.GetMetadataCount - 1
    
        categoryFound = False
    
        'Retrieve the next metadata entry
        curMetadata = pdImages(g_CurrentImage).imgMetadata.GetMetadataEntry(i)
        chkGroup = curMetadata.TagGroupFriendly
        
        If (Not curMetadata.InternalUseOnly) Then
        
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
            
        End If
    
    Next i
    
    'We can now populate the left-side list box with the categories we found.  While doing this, find
    ' the category with the highest tag count.
    highestCategoryCount = 0
    
    For i = 0 To numOfCategories - 1
        lstGroup.AddItem mdCategories(i).Name, i
        If mdCategories(i).Count > highestCategoryCount Then highestCategoryCount = mdCategories(i).Count
    Next i
    
    'We can now build a 2D array that contains all tags, sorted by category.  Why not do this above?  Because
    ' it's computationally expensive to constantly redim arrays in VB, and this technique allows us to redim
    ' the main tag array only once, after all values have been tallied.
    ReDim allTags(0 To numOfCategories - 1, 0 To highestCategoryCount - 1) As PDMetadataItem
    ReDim curTagCount(0 To numOfCategories - 1) As Long
    
    For i = 0 To pdImages(g_CurrentImage).imgMetadata.GetMetadataCount - 1
        
        'As above, retrieve the next metadata entry
        curMetadata = pdImages(g_CurrentImage).imgMetadata.GetMetadataEntry(i)
        chkGroup = curMetadata.TagGroupFriendly
        
        'Find the matching group in the Group array, then insert this tag into place
        For j = 0 To numOfCategories - 1
            If StrComp(mdCategories(j).Name, chkGroup) = 0 Then
            
                allTags(j, curTagCount(j)) = curMetadata
                curTagCount(j) = curTagCount(j) + 1
                Exit For
                
            End If
        Next j
        
    Next i
    
    lstGroup.Caption = g_Language.TranslateMessage("%1 groups in this image:", numOfCategories)
    
    'Populate the simple/technical switches at the bottom
    btsTechnical(0).AddItem "simple", 0
    btsTechnical(0).AddItem "technical", 1
    btsTechnical(0).ListIndex = 0
    
    btsTechnical(1).AddItem "simple", 0
    btsTechnical(1).AddItem "technical", 1
    btsTechnical(1).ListIndex = 0
    
    'Select the first group by default
    lstGroup.ListIndex = 0
    
    'Prep any other interface components
    btsEditPanel.AddItem "tag editor", 0
    btsEditPanel.AddItem "tag settings", 1
    btsEditPanel.ListIndex = 0
    
    'Technical metadata reports are only available for images that actually exist on disk (vs clipboard or scanned images)
    If Len(pdImages(g_CurrentImage).imgStorage.GetEntry_String("CurrentLocationOnDisk")) <> 0 Then
        lblTechnicalReport.Visible = True
        cmdTechnicalReport.Visible = True
    Else
        lblTechnicalReport.Visible = False
        cmdTechnicalReport.Visible = False
    End If
    
    'Give ExifTool credit for its amazing work!
    lblExifTool.Caption = g_Language.TranslateMessage("Metadata support is provided by the 3rd-party ExifTool library.  ExifTool is Copyright 2016 by Phil Harvey.")
    
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Before the metadata list box does any painting, we need to retrieve relevant colors from PD's primary theming class.
' Note that this step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDMD_TitleSelected, "TitleSelected", IDE_GRAY
        .LoadThemeColor PDMD_TitleUnselected, "TitleUnselected", IDE_GRAY
        .LoadThemeColor PDMD_DescriptionSelected, "TitleSelected", IDE_GRAY
        .LoadThemeColor PDMD_DescriptionUnselected, "TitleUnselected", IDE_GRAY
    End With
End Sub

'Fill the metadata list with all entries from the current category
Private Sub UpdateMetadataList()
    
    Dim curCategory As Long
    curCategory = lstGroup.ListIndex
    
    lstMetadata.SetAutomaticRedraws False
    lstMetadata.Clear
    
    Dim i As Long
    For i = 0 To mdCategories(curCategory).Count - 1
        lstMetadata.AddItem , i
    Next i
    
    lstMetadata.SetAutomaticRedraws True, True
    
End Sub

Private Sub lstGroup_Click()
    
    Dim curCategory As Long
    curCategory = lstGroup.ListIndex
    
    If mdCategories(curCategory).Count = 1 Then
        lstMetadata.Caption = g_Language.TranslateMessage("1 tag in this category:")
    Else
        lstMetadata.Caption = g_Language.TranslateMessage("%1 tags in this category:", mdCategories(curCategory).Count)
    End If
    
    'Update the metadata list to reflect the new category
    UpdateMetadataList
    
End Sub

Private Sub lstMetadata_Click()
    UpdateTagView
End Sub

Private Sub lstMetadata_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    'Calculate colors
    Dim titleColor As Long, descriptionColor As Long
    If itemIsSelected Then
        titleColor = m_Colors.RetrieveColor(PDMD_TitleSelected, lstMetadata.Enabled, , itemIsHovered)
        descriptionColor = m_Colors.RetrieveColor(PDMD_DescriptionSelected, lstMetadata.Enabled, , itemIsHovered)
    Else
        titleColor = m_Colors.RetrieveColor(PDMD_TitleUnselected, lstMetadata.Enabled, , itemIsHovered)
        descriptionColor = m_Colors.RetrieveColor(PDMD_DescriptionUnselected, lstMetadata.Enabled, , itemIsHovered)
    End If
    
    Dim blockCategory As Long
    blockCategory = lstGroup.ListIndex
    
    Dim tmpRectF As RECTF
    CopyMemory ByVal VarPtr(tmpRectF), ByVal ptrToRectF, 16&
    
    Dim offsetY As Single, offsetX As Single
    offsetX = tmpRectF.Left + FixDPI(8)
    offsetY = tmpRectF.Top + FixDPI(7)
    
    Dim thisTag As PDMetadataItem
    thisTag = allTags(blockCategory, itemIndex)
    
    Dim linePadding As Long
    linePadding = FixDPI(4)
    
    Dim mHeight As Single
    Dim drawString As String, numericalPrefix As String
    
    numericalPrefix = CStr(itemIndex + 1) & " - "
        
    If (btsTechnical(0).ListIndex = 0) Then
        drawString = thisTag.TagNameFriendly
    Else
        drawString = thisTag.TagGroupAndName
    End If
        
    'Notify the user of text we were unable to convert to a human-readable value
    If thisTag.IsTagBinary Then
        drawString = drawString & " " & g_Language.TranslateMessage("(encoding unknown)")
    End If
    
    'Start with the simplest field: the tag title (readable form)
    m_TitleFont.AttachToDC bufferDC
    m_TitleFont.SetFontColor titleColor
    m_TitleFont.FastRenderText offsetX + 0, offsetY + 0, numericalPrefix & drawString
                
    'Below the tag title, add the human-friendly description
    mHeight = m_TitleFont.GetHeightOfString(drawString) + linePadding
    m_TitleFont.ReleaseFromDC
    
    If (btsTechnical(1).ListIndex = 0) Then
        drawString = thisTag.TagValueFriendly
    Else
        drawString = thisTag.TagValue
    End If
    
    m_DescriptionFont.AttachToDC bufferDC
    m_DescriptionFont.SetFontColor descriptionColor
    m_DescriptionFont.FastRenderTextWithClipping offsetX + m_TitleFont.GetWidthOfString(numericalPrefix), offsetY + mHeight, (tmpRectF.Left + tmpRectF.Width) - offsetX - FixDPI(17), m_DescriptionFont.GetHeightOfString(drawString), drawString
    m_DescriptionFont.ReleaseFromDC
    
End Sub

Private Sub UpdateTagView()

    Dim curGroup As Long, curTag As Long
    curGroup = lstGroup.ListIndex
    curTag = lstMetadata.ListIndex
    
    If (curTag >= 0) Then
    
        With allTags(curGroup, curTag)
            Me.lblGroup.Caption = .TagGroup
            Me.lblTable.Caption = .TagTable
            Me.lblTagID.Caption = .TagID
            Me.lblTagName.Caption = .TagNameFriendly
        
            If CBool(chkFriendlyValue.Value) Then
                Me.txtValue.Text = .TagValue
            Else
                Me.txtValue.Text = .TagValueFriendly
            End If
        
        End With
        
    End If
    
End Sub

