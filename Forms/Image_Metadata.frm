VERSION 5.00
Begin VB.Form FormMetadata 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Browse image metadata"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14190
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
   ScaleHeight     =   539
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   946
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdHyperlink hypExiftool 
      Height          =   255
      Left            =   120
      Top             =   6960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "visit the ExifTool homepage"
      FontSize        =   9
      URL             =   "http://www.sno.phy.queensu.ca/~phil/exiftool/"
   End
   Begin PhotoDemon.pdButtonStrip btsEditPanel 
      Height          =   975
      Left            =   8040
      TabIndex        =   3
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      Caption         =   "tools"
   End
   Begin PhotoDemon.pdListBox lstGroup 
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   11245
      Caption         =   "metadata groups in this image"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   7335
      Width           =   14190
      _ExtentX        =   25030
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdLabel lblExifTool 
      Height          =   255
      Left            =   120
      Top             =   6600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   ""
      FontSize        =   9
      ForeColor       =   -2147483640
      Layout          =   1
   End
   Begin PhotoDemon.pdListBoxOD lstMetadata 
      Height          =   6375
      Left            =   3765
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   11245
      Caption         =   "tags in this category"
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   6015
      Index           =   0
      Left            =   8040
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   406
      TabIndex        =   4
      Top             =   1200
      Width           =   6090
      Begin PhotoDemon.pdListBox lstValue 
         Height          =   3000
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5292
      End
      Begin PhotoDemon.pdLabel lblValue 
         Height          =   3000
         Left            =   195
         Top             =   390
         Visible         =   0   'False
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   5292
         Caption         =   ""
         Layout          =   1
      End
      Begin PhotoDemon.pdTextBox txtValue 
         Height          =   3000
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5292
         Multiline       =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblTagName 
         Height          =   300
         Left            =   120
         Top             =   0
         Width           =   5850
         _ExtentX        =   6085
         _ExtentY        =   529
         Caption         =   ""
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTagType 
         Height          =   300
         Left            =   120
         Top             =   3480
         Width           =   5895
         _ExtentX        =   5741
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin PhotoDemon.pdLabel lblGroupDescription 
         Height          =   1095
         Left            =   120
         Top             =   4560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1931
         Alignment       =   2
         FontItalic      =   -1  'True
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblWarning 
         Height          =   540
         Left            =   120
         Top             =   3840
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   953
         Caption         =   ""
         Layout          =   1
         UseCustomForeColor=   -1  'True
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   6015
      Index           =   1
      Left            =   8040
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   406
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   6090
      Begin PhotoDemon.pdButtonStrip btsTechnical 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   1720
         Caption         =   "tag names"
      End
      Begin PhotoDemon.pdButtonStrip btsTechnical 
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   1720
         Caption         =   "tag values"
      End
      Begin PhotoDemon.pdButton cmdTechnicalReport 
         Height          =   555
         Left            =   240
         TabIndex        =   8
         Top             =   2550
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   979
         Caption         =   "Generate full metadata report (HTML)..."
      End
      Begin PhotoDemon.pdLabel lblTechnicalReport 
         Height          =   270
         Left            =   120
         Top             =   2160
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   476
         Caption         =   "advanced"
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

Private Type MDCategory
    Name As String
    Count As Long
    LastListIndex As Long
End Type

Private m_MDCategories() As MDCategory
Private m_NumOfCategories As Long
Private m_LargestCategoryCount As Long

'This array holds all tags currently in storage, but sorted into a 2D array by category
Private m_AllTags() As PDMetadataItem

'Height of each metadata content block
Private Const BLOCKHEIGHT As Long = 46

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
    PDMD_TagIsWritable = 4
    PDMD_TagIsNotWritable = 5
    PDMD_TagIsUnsafe = 6
    PDMD_TextTagEditError = 7
    [_Last] = 7
    [_Count] = 8
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
    ReDim m_MDCategories(0 To 3) As MDCategory
    m_NumOfCategories = 0
    
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
            For j = 0 To m_NumOfCategories - 1
                If StrComp(m_MDCategories(j).Name, chkGroup, vbBinaryCompare) = 0 Then
                    categoryFound = True
                    m_MDCategories(j).Count = m_MDCategories(j).Count + 1
                    Exit For
                End If
            Next j
            
            'If no matching category was found, create a new category entry
            If (Not categoryFound) Then
                If (m_NumOfCategories) > UBound(m_MDCategories) Then ReDim Preserve m_MDCategories(0 To m_NumOfCategories * 2 - 1) As MDCategory
                m_MDCategories(m_NumOfCategories).Name = chkGroup
                m_MDCategories(m_NumOfCategories).Count = 1
                m_NumOfCategories = m_NumOfCategories + 1
            End If
            
        End If
    
    Next i
    
    'With all categories now detected, we want to sort the list
    SortCategoryList
    
    'We can now populate the left-side list box with the categories we found.  While doing this, find
    ' the category with the highest tag count.
    m_LargestCategoryCount = 0
    
    lstGroup.SetAutomaticRedraws False
    For i = 0 To m_NumOfCategories - 1
        lstGroup.AddItem m_MDCategories(i).Name, i, CBool(StrComp(LCase$(m_MDCategories(i).Name), "inferred", vbBinaryCompare) = 0)
        If m_MDCategories(i).Count > m_LargestCategoryCount Then m_LargestCategoryCount = m_MDCategories(i).Count
    Next i
    lstGroup.SetAutomaticRedraws True, True
    
    'We can now build a 2D array that contains all tags, sorted by category.  Why not do this above?  Because
    ' it's computationally expensive to constantly redim arrays in VB, and this technique allows us to redim
    ' the main tag array only once, after all values have been tallied.
    ReDim m_AllTags(0 To m_NumOfCategories - 1, 0 To m_LargestCategoryCount - 1) As PDMetadataItem
    
    Dim curTagCount() As Long
    ReDim curTagCount(0 To m_NumOfCategories - 1) As Long
    
    For i = 0 To pdImages(g_CurrentImage).imgMetadata.GetMetadataCount - 1
        
        'As above, retrieve the next metadata entry
        curMetadata = pdImages(g_CurrentImage).imgMetadata.GetMetadataEntry(i)
        chkGroup = curMetadata.TagGroupFriendly
        
        'By default, PD only grabs as much metadata information as it needs to successfully write the metadata out to file.
        ' Editing requires additional tag data.  Populate that now, by synchronizing each tag against its ExifTool
        ' database entry.
        ExifTool.FillTagFromDatabase curMetadata
        
        'Find the matching group in the Group array, then insert this tag into place
        For j = 0 To m_NumOfCategories - 1
            If StrComp(m_MDCategories(j).Name, chkGroup) = 0 Then
                m_AllTags(j, curTagCount(j)) = curMetadata
                curTagCount(j) = curTagCount(j) + 1
                Exit For
            End If
        Next j
        
    Next i
    
    lstGroup.Caption = g_Language.TranslateMessage("%1 groups in this image:", m_NumOfCategories)
    
    'Populate the simple/technical switches at the bottom
    btsTechnical(0).AddItem "simple", 0
    btsTechnical(0).AddItem "technical", 1
    btsTechnical(0).ListIndex = 0
    
    btsTechnical(1).AddItem "simple", 0
    btsTechnical(1).AddItem "technical", 1
    btsTechnical(1).ListIndex = 0
    
    'Select the first group by default
    lstGroup.ListIndex = 0
    If (lstMetadata.ListCount > 0) Then lstMetadata.ListIndex = 0
    
    'Prep any other interface components
    btsEditPanel.AddItem "edit", 0
    btsEditPanel.AddItem "settings", 1
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
    lblExifTool.Caption = g_Language.TranslateMessage("Metadata support is made possible by the ExifTool library.")
    
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub SortCategoryList()
    
    Dim i As Long, j As Long
    
    'We first want to sort the group names alphabetically.  The easiest way to do this is with pdStringStack.
    Dim cNames As pdStringStack
    Set cNames = New pdStringStack
    For i = 0 To m_NumOfCategories - 1
        cNames.AddString m_MDCategories(i).Name
    Next i
    cNames.SortAlphabetically
    
    'We now want to do something weird.  The "System", "File", and "Composite" categories should always come first.
    ' These categories are largely un-editable, and they are persistent across image formats.
    For i = 0 To m_NumOfCategories - 1
        If StrComp(LCase$(cNames.GetString(i)), "system", vbBinaryCompare) = 0 Then
            cNames.MoveStringToNewPosition i, 0
            Exit For
        End If
    Next i
    
    For i = 0 To m_NumOfCategories - 1
        If StrComp(LCase$(cNames.GetString(i)), "file", vbBinaryCompare) = 0 Then
            cNames.MoveStringToNewPosition i, 1
            Exit For
        End If
    Next i
    
    For i = 0 To m_NumOfCategories - 1
        If StrComp(LCase$(cNames.GetString(i)), "inferred", vbBinaryCompare) = 0 Then
            cNames.MoveStringToNewPosition i, 2
            Exit For
        End If
    Next i
    
    'We now want to sort the main category list to match this order.
    Dim tmpCat As MDCategory
    For i = 0 To m_NumOfCategories - 1
        For j = i To m_NumOfCategories - 1
            If (StrComp(cNames.GetString(i), m_MDCategories(j).Name, vbBinaryCompare) = 0) And (i <> j) Then
                tmpCat = m_MDCategories(i)
                m_MDCategories(i) = m_MDCategories(j)
                m_MDCategories(j) = tmpCat
                Exit For
            End If
        Next j
    Next i
    
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
        .LoadThemeColor PDMD_TagIsWritable, "TagIsWritable", RGB(0, 255, 0)
        .LoadThemeColor PDMD_TagIsNotWritable, "TagIsNotWritable", RGB(255, 0, 0)
        .LoadThemeColor PDMD_TagIsUnsafe, "TagIsUnsafe", RGB(0, 255, 255)
        .LoadThemeColor PDMD_TextTagEditError, "TextTagEditError", RGB(255, 0, 0)
    End With
End Sub

'Fill the metadata list with all entries from the current category
Private Sub UpdateMetadataList()
    
    Dim curCategory As Long
    curCategory = lstGroup.ListIndex
    
    lstMetadata.SetAutomaticRedraws False
    lstMetadata.Clear
    
    Dim i As Long
    For i = 0 To m_MDCategories(curCategory).Count - 1
        lstMetadata.AddItem , i
    Next i
    
    lstMetadata.SetAutomaticRedraws True, True
    
End Sub

Private Sub lstGroup_Click()
    
    Dim curCategory As Long
    curCategory = lstGroup.ListIndex
    
    If m_MDCategories(curCategory).Count = 1 Then
        lstMetadata.Caption = g_Language.TranslateMessage("1 tag in this category")
    Else
        lstMetadata.Caption = g_Language.TranslateMessage("%1 tags in this category", m_MDCategories(curCategory).Count)
    End If
    
    'Some categories display a "helper" description.  Check for that now.
    Dim catName As String, groupDescription As String
    catName = LCase$(m_MDCategories(curCategory).Name)
    If StrComp(catName, "system", vbBinaryCompare) = 0 Then
        groupDescription = g_Language.TranslateMessage("Note: ""System"" tags are provided by the operating system.  They are not embedded as traditional metadata.")
    ElseIf StrComp(catName, "file", vbBinaryCompare) = 0 Then
        groupDescription = g_Language.TranslateMessage("Note: ""File"" tags are required by this image format.  They are not embedded as traditional metadata.")
    ElseIf StrComp(catName, "inferred", vbBinaryCompare) = 0 Then
        groupDescription = g_Language.TranslateMessage("Note: ""Inferred"" tags are hypothetical values inferred from other metadata.  They are not embedded as traditional metadata.")
    End If
    
    'If a helper description exists, show/hide the description label to match
    If (Len(groupDescription) = 0) Then
        lblGroupDescription.Visible = False
    Else
        lblGroupDescription.Caption = groupDescription
        lblGroupDescription.Visible = True
    End If
    
    'Update the metadata list to reflect the new category
    UpdateMetadataList
    
    'We remember the last ListIndex for each category.  With the listbox successfully filled, set the new index now
    lstMetadata.ListIndex = m_MDCategories(curCategory).LastListIndex
    
End Sub

Private Sub lstMetadata_Click()
    m_MDCategories(lstGroup.ListIndex).LastListIndex = lstMetadata.ListIndex
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
    offsetY = tmpRectF.Top + FixDPI(4)
    
    Dim thisTag As PDMetadataItem
    thisTag = m_AllTags(blockCategory, itemIndex)
    
    Dim linePadding As Long
    linePadding = FixDPI(3)
    
    'Note that we deliberately maintain the numerical prefix as a separate entity; we need its size (in pixels) to calculate
    ' proper padding for the description line of text.
    Dim numericalPrefix As String
    numericalPrefix = CStr(itemIndex + 1) & "  "
    
    Dim drawString As String
    If (btsTechnical(0).ListIndex = 0) Then
        drawString = thisTag.TagNameFriendly
    Else
        drawString = thisTag.TagGroupAndName
    End If
    
    'Before rendering the title, we render a colored bar to indicate the write-ability of this tag
    Dim tagColor As Long
    If thisTag.DB_IsWritable Then
        If (thisTag.DBF_IsUnsafe Or thisTag.DBF_IsProtected Or thisTag.DBF_IsMandatory) Then
            tagColor = m_Colors.RetrieveColor(PDMD_TagIsUnsafe, Me.Enabled)
        Else
            tagColor = m_Colors.RetrieveColor(PDMD_TagIsWritable, Me.Enabled)
        End If
    Else
        tagColor = m_Colors.RetrieveColor(PDMD_TagIsNotWritable, Me.Enabled)
    End If
    
    Dim spaceWidth As Single
    spaceWidth = m_TitleFont.GetWidthOfString(" ")
    With tmpRectF
        GDI_Plus.GDIPlusFillRectToDC bufferDC, .Left, .Top, (offsetX - .Left) + m_TitleFont.GetWidthOfString(CStr(itemIndex + 1)) + spaceWidth + 1, .Height + 1, tagColor
    End With
    
    'Start with the simplest field: the tag title (readable form)
    m_TitleFont.AttachToDC bufferDC
    m_TitleFont.SetFontColor titleColor
    m_TitleFont.FastRenderText offsetX + 0, offsetY + 0, numericalPrefix & drawString
                
    'Below the tag title, add the human-friendly description
    Dim mHeight As Single
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
    
        With m_AllTags(curGroup, curTag)
            lblTagName.Caption = .TagNameFriendly
            
            'Editable values are presented using an edit-friendly control (text box, dropdown, etc).  Non-editable values
            ' only get a label.
            If .DB_IsWritable Then
                lblValue.Visible = False
                
                'Values that are part of a hardcoded list are available via dropdown, but *only* if they consist of a single entry.
                ' (Some list values, like JPEG component configuration, are hard-coded list values x4.  This is very difficult
                '  to handle programmatically, so we default to text entry in those cases.)
                If .DB_HardcodedList And (.DB_TypeCount < 2) Then
                    
                    lstValue.SetAutomaticRedraws False
                    lstValue.Clear
                    
                    Dim newListIndex As Long, listIndexFound As Boolean
                    newListIndex = -1
                    listIndexFound = False
                    
                    Dim strID As String, strValue As String
                    
                    Dim i As Long
                    For i = 0 To .DB_StackValues.GetNumOfStrings - 1
                        
                        strID = .DB_StackIDs.GetString(i)
                        strValue = .DB_StackValues.GetString(i)
                        
                        'If this entry's ID and VALUE are identical (as they are for many XMP tags, as XMP deals solely in text),
                        ' it's redundant to list both.  For EXIF values, however, the ID may be numerical while the VALUE is
                        ' a nice, human-readable chunk of text, so display both.
                        If StrComp(strID, strValue, vbBinaryCompare) = 0 Then
                            lstValue.AddItem strValue
                        Else
                            lstValue.AddItem "(" & strID & ") " & strValue
                        End If
                        
                        If StrComp(.TagValue, .DB_StackIDs.GetString(i), vbBinaryCompare) = 0 Then
                            newListIndex = i
                            listIndexFound = True
                        
                        'As a failsafe, also compare the "print-friendly" version of the current value
                        Else
                            If StrComp(.TagValueFriendly, .DB_StackValues.GetString(i), vbBinaryCompare) = 0 Then
                                newListIndex = i
                                listIndexFound = True
                            End If
                        End If
                        
                    Next i
                    
                    'If a matching list index is *not* found, add the current value (whatever it is) to the list in the
                    ' final position.
                    If (Not listIndexFound) Then
                        lstValue.AddItem "(" & .TagValue & ") " & .TagValueFriendly
                        lstValue.ListIndex = (lstValue.ListCount - 1)
                    Else
                        lstValue.ListIndex = newListIndex
                    End If
                    
                    lstValue.SetAutomaticRedraws True, True
                    lstValue.Visible = True
                    txtValue.Visible = False
                
                'Any other values (text and numeric entry, among others) are handled via text box
                Else
                    lstValue.Visible = False
                    txtValue.Visible = True
                    If (btsTechnical(1).ListIndex = 0) Then txtValue.Text = .TagValueFriendly Else txtValue.Text = .TagValue
                End If
                
            Else
                lstValue.Visible = False
                txtValue.Visible = False
                lblValue.Visible = True
                If (btsTechnical(1).ListIndex = 0) Then lblValue.Caption = .TagValueFriendly Else lblValue.Caption = .TagValue
            End If
            
            'The title caption changes depending on the data type, but *only* if the tag is writable!
            If .DB_IsWritable Then
                If .DB_HardcodedList Then
                    lblTagType.Visible = False
                Else
                    
                    lblTagType.UseCustomForeColor = False
                    
                    Dim strTagRestrictions As String
                    strTagRestrictions = ConvertDataTypeToString(m_AllTags(curGroup, curTag))
                    
                    'We only list restrictions if necessary.  Generic "text" tags are treated as though they have no restrictions.
                    If (Len(strTagRestrictions) <> 0) And (StrComp(strTagRestrictions, "text", vbBinaryCompare) <> 0) Then
                        lblTagType.Caption = g_Language.TranslateMessage("tag restrictions: ") & strTagRestrictions
                        lblTagType.Visible = True
                    Else
                        lblTagType.Visible = False
                    End If
                    
                End If
                
                'Protected tags can technically be edited, but there may be unforeseen consequences.  Let the user know.
                If (.DBF_IsUnsafe Or .DBF_IsProtected Or .DBF_IsMandatory) Then
                    lblWarning.UseCustomForeColor = True
                    lblWarning.ForeColor = m_Colors.RetrieveColor(PDMD_TextTagEditError)
                    lblWarning.Caption = g_Language.TranslateMessage("NOTE: this is a protected tag.  You can edit it, but PhotoDemon may overwrite your changes to produce a valid image file.")
                    
                    If lblTagType.Visible Then
                        lblWarning.SetTop lblTagType.GetTop + lblTagType.GetHeight + FixDPI(8)
                    Else
                        lblWarning.SetTop txtValue.GetTop + txtValue.GetHeight + FixDPI(8)
                    End If
                    
                    lblWarning.Visible = True
                Else
                    lblWarning.Visible = False
                End If
                
            Else
                lblTagType.ForeColor = m_Colors.RetrieveColor(PDMD_TextTagEditError)
                lblTagType.UseCustomForeColor = True
                lblTagType.Caption = g_Language.TranslateMessage("NOTE: this tag cannot be edited")
                lblTagType.Visible = True
                lblWarning.Visible = False
            End If
            
        End With
        
    End If
    
End Sub

Private Function ConvertDataTypeToString(ByRef srcMetadata As PDMetadataItem) As String
    
    Dim strResult As String
    
    Dim countPresent As Boolean, countValue As Long
    countPresent = (srcMetadata.DB_TypeCount <> 0)
    countValue = srcMetadata.DB_TypeCount
    If countValue < 2 Then countValue = 1
    
    Select Case srcMetadata.DB_DataTypeStrict
    
        Case MD_int8s
            strResult = g_Language.TranslateMessage("integers only [-127 to 127]")
            If countPresent Then strResult = CStr(countValue) & " x " & strResult
        Case MD_int8u
            strResult = g_Language.TranslateMessage("integers only [0 to 255]")
        Case MD_int16s
            strResult = g_Language.TranslateMessage("integers only [-32,768 to 32,767]")
        Case MD_int16u
            strResult = g_Language.TranslateMessage("integers only [0 to 65,535]")
        Case MD_int32s
            strResult = g_Language.TranslateMessage("integers only")
        Case MD_int32u
            strResult = g_Language.TranslateMessage("integers >= 0")
        Case MD_int64s
            strResult = g_Language.TranslateMessage("integers only")
        Case MD_int64u
            strResult = g_Language.TranslateMessage("integers >= 0")
        Case MD_rational32s
            strResult = g_Language.TranslateMessage("numbers only")
        Case MD_rational32u
            strResult = g_Language.TranslateMessage("numbers >= 0")
        Case MD_rational64s
            strResult = g_Language.TranslateMessage("numbers only")
        Case MD_rational64u
            strResult = g_Language.TranslateMessage("numbers >= 0")
        Case MD_fixed16s
            strResult = g_Language.TranslateMessage("numbers only")
        Case MD_fixed16u
            strResult = g_Language.TranslateMessage("numbers >= 0")
        Case MD_fixed32s
            strResult = g_Language.TranslateMessage("numbers only")
        Case MD_fixed32u
            strResult = g_Language.TranslateMessage("numbers >= 0")
        Case MD_float
            strResult = g_Language.TranslateMessage("numbers only")
        Case MD_double
            strResult = g_Language.TranslateMessage("numbers only")
        Case MD_extended
            strResult = g_Language.TranslateMessage("numbers only")
        Case MD_ifd
            strResult = g_Language.TranslateMessage("must be a valid file position marker")
        Case MD_ifd64
            strResult = g_Language.TranslateMessage("must be a valid file position marker")
        Case MD_string
            strResult = g_Language.TranslateMessage("text")
        Case MD_undef
            Debug.Print "The selected tag actually has an ""undefined"" data format, but PD displays ""text"" as a convenience."
            strResult = g_Language.TranslateMessage("text")
        Case MD_binary
            strResult = g_Language.TranslateMessage("must be valid binary data")
        Case MD_integerstring
            strResult = g_Language.TranslateMessage("list of digits")
        Case MD_floatstring
            strResult = g_Language.TranslateMessage("any real number")
        Case MD_rationalstring
            strResult = g_Language.TranslateMessage("any real number")
        Case MD_datestring
            strResult = g_Language.TranslateMessage("date (YYYY:mm:dd HH:MM:SS[.ss][+/-HH:MM])")
        Case MD_booleanstring
            strResult = g_Language.TranslateMessage("true or false")
        Case MD_digits
            strResult = g_Language.TranslateMessage("list of digits")
    
    End Select

    'Some tags will specify a count, e.g. "string [64]" or "integer [4]" - with the last being common for RGBA entries.
    ' We'll append such a count to the type description, for convenience.
    If countPresent Then
        
        Select Case srcMetadata.DB_DataTypeStrict
    
            Case MD_int8s, MD_int8u, MD_int16s, MD_int16u
                strResult = CStr(countValue) & " x " & strResult
            Case MD_int32s, MD_int32u, MD_int64s, MD_int64u
                strResult = CStr(countValue) & " x " & strResult
            Case MD_rational32s, MD_rational32u, MD_rational64s, MD_rational64u
                strResult = CStr(countValue) & " x " & strResult
            Case MD_fixed16s, MD_fixed16u, MD_fixed32s, MD_fixed32u
                strResult = CStr(countValue) & " x " & strResult
            Case MD_float, MD_double, MD_extended
                strResult = CStr(countValue) & " x " & strResult
            Case MD_ifd, MD_ifd64
                strResult = CStr(countValue) & " x " & strResult
            Case MD_string
                strResult = strResult & " [" & g_Language.TranslateMessage("%1 characters max", CStr(countValue)) & "]"
            Case MD_undef, MD_binary
                strResult = strResult & " [" & g_Language.TranslateMessage("%1 bytes max", CStr(countValue)) & "]"
            Case MD_integerstring, MD_digits
                strResult = strResult & " [" & g_Language.TranslateMessage("%1 numbers max", CStr(countValue)) & "]"
            Case MD_floatstring, MD_rationalstring
                strResult = CStr(countValue) & " x " & strResult
            Case MD_datestring
                strResult = CStr(countValue) & " x " & strResult
            Case MD_booleanstring
                strResult = CStr(countValue) & " x " & strResult
            
        End Select
        
    End If
    
    If Len(strResult) <> 0 Then ConvertDataTypeToString = strResult
    
End Function
