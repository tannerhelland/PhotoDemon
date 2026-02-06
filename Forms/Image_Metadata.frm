VERSION 5.00
Begin VB.Form FormMetadata 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit metadata"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14190
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
   ScaleHeight     =   539
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   946
   ShowInTaskbar   =   0   'False
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
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3360
      _ExtentX        =   5927
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
   Begin PhotoDemon.pdListBoxOD lstMetadata 
      Height          =   7095
      Left            =   3525
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   11245
      Caption         =   "tags in this category"
   End
   Begin PhotoDemon.pdLabel lblGroupDescription 
      Height          =   495
      Left            =   240
      Top             =   6120
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   1931
      Caption         =   ""
      FontItalic      =   -1  'True
      FontSize        =   9
      Layout          =   3
   End
   Begin PhotoDemon.pdButtonToolbox btnGroupOptions 
      Height          =   630
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   6600
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox btnGroupOptions 
      Height          =   630
      Index           =   1
      Left            =   900
      TabIndex        =   14
      Top             =   6600
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      AutoToggle      =   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox btnGroupOptions 
      Height          =   630
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Top             =   6600
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      AutoToggle      =   -1  'True
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6015
      Index           =   1
      Left            =   8040
      Top             =   1200
      Visible         =   0   'False
      Width           =   6090
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonStrip btsTechnical 
         Height          =   975
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   0
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   1720
         Caption         =   "tag names"
      End
      Begin PhotoDemon.pdButtonStrip btsTechnical 
         Height          =   975
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   1720
         Caption         =   "tag values"
      End
      Begin PhotoDemon.pdButton cmdTechnicalReport 
         Height          =   555
         Left            =   420
         TabIndex        =   8
         Top             =   3240
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   979
         Caption         =   "Generate full metadata report (HTML)..."
      End
      Begin PhotoDemon.pdLabel lblTechnicalReport 
         Height          =   270
         Left            =   240
         Top             =   2160
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   476
         Caption         =   "advanced"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdHyperlink hypExiftool 
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   5700
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "visit the ExifTool homepage"
         FontSize        =   9
         URL             =   "http://www.sno.phy.queensu.ca/~phil/exiftool/"
      End
      Begin PhotoDemon.pdLabel lblExifTool 
         Height          =   255
         Left            =   120
         Top             =   5370
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   ""
         FontSize        =   9
         ForeColor       =   -2147483640
         Layout          =   1
      End
      Begin PhotoDemon.pdButton cmdMarkPrivateTags 
         Height          =   555
         Left            =   420
         TabIndex        =   15
         Top             =   2640
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   979
         Caption         =   "Remove tags that might contain personal information"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6015
      Index           =   0
      Left            =   8040
      Top             =   1200
      Width           =   6090
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonToolbox btnTagOptions 
         Height          =   630
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   4560
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
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
         Layout          =   3
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
         Layout          =   3
      End
      Begin PhotoDemon.pdLabel lblWarning 
         Height          =   540
         Left            =   150
         Top             =   3840
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   953
         Caption         =   ""
         Layout          =   1
         UseCustomForeColor=   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnTagOptions 
         Height          =   630
         Index           =   1
         Left            =   780
         TabIndex        =   12
         Top             =   4560
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1111
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox btnTagOptions 
         Height          =   630
         Index           =   2
         Left            =   1440
         TabIndex        =   5
         Top             =   4560
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1111
         AutoToggle      =   -1  'True
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
'Copyright 2013-2026 by Tanner Helland
'Created: 27/May/13
'Last updated: 23/January/21
'Last update: fix display of some esoteric list-type values
'
'As of version 6.0, PhotoDemon now provides support for loading and saving image metadata.  What is metadata, you ask?
' See https://en.wikipedia.org/wiki/Metadata#Photographs for more details.
'
'This dialog interacts heavily with the pdMetadata class to present users with a relatively simple interface for
' perusing (and eventually, editing) an image's metadata.
'
'Designing this dialog was quite difficult as it is impossible to predict what metadata types and entries might exist in
' an image file, so I've opted for the most flexible system I can.  No assumptions are made about present categories or
' tag counts, so any type or amount of metadata should theoretically be viewable.
'
'Categories are displayed on the left, and selecting a category repopulates the fields on the right.  Future updates
' could include the ability to add individual tags, and ideally, a simplified interface for common tags...
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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

'The IDs of the last group and metadata item selected.  We need to track these so we can "apply" changes whenever the
' current item "loses focus".  (I use quotation marks because we use custom focus events in place of the traditional ones.)
Private m_GroupIndex As Long, m_TagIndex As Long

'Focus changes trigger various responses; we ignore them until the dialog is fully visible
Private m_DialogFinishedLoading As Boolean

'Height of each metadata content block
Private Const BLOCKHEIGHT As Long = 46

'Font objects for rendering
Private m_TitleFont As pdFont, m_DescriptionFont As pdFont

'Tag buttons.  These may not all be available on all tags (e.g. some tags cannot be edited, so they don't get a "reset" button)
Private Enum PDMETADATA_TAG_BUTTONS
    MDTB_Remove = 0
    MDTB_Copy = 1
    MDTB_Reset = 2
End Enum

#If False Then
    Private Const MDTB_Remove = 0, MDTB_Copy = 1, MDTB_Reset = 2
#End If

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
    PDMD_TagBackgroundDeleted = 8
    PDMD_TagBackgroundEdited = 9
    [_Last] = 9
    [_Count] = 10
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Private Sub btnGroupOptions_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    Dim curCategory As Long
    curCategory = lstGroup.ListIndex
    
    Dim i As Long
    
    Select Case Index
    
        'Delete entire group
        Case MDTB_Remove
            Dim removalState As Boolean
            removalState = btnGroupOptions(MDTB_Remove).Value
            For i = 0 To m_MDCategories(curCategory).Count - 1
                m_AllTags(curCategory, i).TagMarkedForRemoval = removalState
            Next i
        
        'Copy entire group to clipboard
        Case MDTB_Copy
            
            Dim cString As pdString
            Set cString = New pdString
            Const COLON_SPACE As String = ": "
            
            For i = 0 To m_MDCategories(curCategory).Count - 1
                cString.Append m_AllTags(curCategory, i).TagNameFriendly
                cString.Append COLON_SPACE
                cString.AppendLine m_AllTags(curCategory, i).TagValueFriendly
            Next i
            
            g_Clipboard.ClipboardCopy_Text cString.ToString()
            
        'Reset entire group
        Case MDTB_Reset
            For i = 0 To m_MDCategories(curCategory).Count - 1
                m_AllTags(curCategory, i).UserIDNew = vbNullString
                m_AllTags(curCategory, i).UserValueNew = vbNullString
                m_AllTags(curCategory, i).UserModifiedAllSessions = False
                m_AllTags(curCategory, i).UserModifiedThisSession = False
            Next i
    
    End Select
    
    lstMetadata.SetAutomaticRedraws True, True
    UpdateGroupButtonEnablement
    UpdateTagView

End Sub

Private Sub btnTagOptions_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    Select Case Index
    
        Case MDTB_Remove
        
        Case MDTB_Copy
        
            If (m_GroupIndex >= 0) And (m_TagIndex >= 0) Then
                g_Clipboard.ClipboardCopy_Text m_AllTags(m_GroupIndex, m_TagIndex).TagNameFriendly & ": " & m_AllTags(m_GroupIndex, m_TagIndex).TagValueFriendly
            End If
            
        Case MDTB_Reset
                
            If (m_GroupIndex >= 0) And (m_TagIndex >= 0) Then
            
                m_AllTags(m_GroupIndex, m_TagIndex).UserIDNew = vbNullString
                m_AllTags(m_GroupIndex, m_TagIndex).UserValueNew = vbNullString
                m_AllTags(m_GroupIndex, m_TagIndex).UserModifiedAllSessions = False
                m_AllTags(m_GroupIndex, m_TagIndex).UserModifiedThisSession = False
            
                Dim backupTagIndex As Long
                backupTagIndex = m_TagIndex
                
                UpdateTagView
                UpdateMetadataList
                lstMetadata.ListIndex = backupTagIndex
            
            End If
    
    End Select

End Sub

Private Sub btsEditPanel_Click(ByVal buttonIndex As Long)
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).Visible = (i = buttonIndex)
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

Private Sub cmdBarMini_OKClick()
    
    'Before doing anything else, trigger a lost-focus check to make sure any edits on the current tag are preserved.
    TagLostFocus False
    
    'When OK is clicked, we need to relay any changed metadata entries back to the parent metadata collection.
    ' (Hypothetically, unchanged entries could be entirely ignored by this step, but PD copies them back to preserve
    ' any semantic data we filled via the primary ExifTool database.  This saves a bit of work if the metadata editor
    ' is invoked again on this image.)
    Dim i As Long, j As Long, k As Long
    Dim curMetadata As PDMetadataItem, targetMetadata As PDMetadataItem
    
    'The local metadata collection is organized as a 2D array where each row has varying length.
    For i = 0 To m_NumOfCategories - 1
        For j = 0 To m_MDCategories(i).Count - 1
            
            curMetadata = m_AllTags(i, j)
            
            'Find the matching tag entry in the parent image's metadata collection
            For k = 0 To PDImages.GetActiveImage.ImgMetadata.GetMetadataCount - 1
                targetMetadata = PDImages.GetActiveImage.ImgMetadata.GetMetadataEntry(k)
                If Strings.StringsEqual(m_MDCategories(i).Name, targetMetadata.TagGroupFriendly, False) Then
                    If Strings.StringsEqual(curMetadata.TagNameFriendly, targetMetadata.TagNameFriendly, False) Then
                        PDImages.GetActiveImage.ImgMetadata.SetMetadataEntryByIndex k, curMetadata
                        Exit For
                    End If
                End If
            Next k
        
        Next j
    Next i
    
    'With all metadata updated, notify the central processor that an Undo update is required
    Process "Edit metadata", False, , UNDO_ImageHeader
    
End Sub

Private Sub cmdMarkPrivateTags_Click()

    'The local metadata collection is organized as a 2D array where each row has varying length.
    Dim i As Long, j As Long
    For i = 0 To m_NumOfCategories - 1
        For j = 0 To m_MDCategories(i).Count - 1
            m_AllTags(i, j).TagMarkedForRemoval = ExifTool.DoesTagHavePrivacyConcerns(m_AllTags(i, j))
        Next j
    Next i
    
    lstMetadata.SetAutomaticRedraws True, True
    UpdateGroupButtonEnablement
    UpdateTagView

End Sub

Private Sub cmdTechnicalReport_Click()
    ExifTool.CreateTechnicalMetadataReport PDImages.GetActiveImage()
End Sub

Private Sub Form_Activate()
    m_DialogFinishedLoading = True
End Sub

Private Sub Form_Load()
    
    lstMetadata.ListItemHeight = Interface.FixDPI(BLOCKHEIGHT)
    
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
    For i = 0 To PDImages.GetActiveImage.ImgMetadata.GetMetadataCount - 1
    
        categoryFound = False
    
        'Retrieve the next metadata entry
        curMetadata = PDImages.GetActiveImage.ImgMetadata.GetMetadataEntry(i)
        chkGroup = curMetadata.TagGroupFriendly
        
        If (Not curMetadata.InternalUseOnly) Then
        
            'Search the current list of known categories for this metadata object's category
            For j = 0 To m_NumOfCategories - 1
                If Strings.StringsEqual(m_MDCategories(j).Name, chkGroup, False) Then
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
        lstGroup.AddItem m_MDCategories(i).Name, i, Strings.StringsEqual(m_MDCategories(i).Name, "inferred", True)
        If (m_MDCategories(i).Count > m_LargestCategoryCount) Then m_LargestCategoryCount = m_MDCategories(i).Count
    Next i
    lstGroup.SetAutomaticRedraws True, True
    
    'We can now build a 2D array that contains all tags, sorted by category.  Why not do this above?  Because
    ' it's computationally expensive to constantly redim arrays in VB, and this technique allows us to redim
    ' the main tag array only once, after all values have been tallied.
    ReDim m_AllTags(0 To m_NumOfCategories - 1, 0 To m_LargestCategoryCount - 1) As PDMetadataItem
    
    Dim curTagCount() As Long
    ReDim curTagCount(0 To m_NumOfCategories - 1) As Long
    
    For i = 0 To PDImages.GetActiveImage.ImgMetadata.GetMetadataCount - 1
        
        'As above, retrieve the next metadata entry, and this time, reset any per-session trackers
        curMetadata.UserModifiedThisSession = False
        curMetadata = PDImages.GetActiveImage.ImgMetadata.GetMetadataEntry(i)
        chkGroup = curMetadata.TagGroupFriendly
        
        'By default, PD only grabs as much metadata information as it needs to successfully write the metadata out to file.
        ' Editing requires additional tag data.  Populate that now, by synchronizing each tag against its ExifTool
        ' database entry.
        ExifTool.FillTagFromDatabase curMetadata
        
        'Find the matching group in the Group array, then insert this tag into place
        For j = 0 To m_NumOfCategories - 1
            If Strings.StringsEqual(m_MDCategories(j).Name, chkGroup, False) Then
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
    lstGroup.ListIndex = 0: m_GroupIndex = 0
    If (lstMetadata.ListCount > 0) Then
        lstMetadata.ListIndex = 0
        m_TagIndex = 0
    End If
    
    'Prep any other interface components
    btsEditPanel.AddItem "edit tags", 0
    btsEditPanel.AddItem "editor options", 1
    btsEditPanel.ListIndex = 0
    
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(32)
    
    btnTagOptions(MDTB_Remove).AssignImage "generic_trash", , buttonSize, buttonSize
    btnTagOptions(MDTB_Remove).AssignTooltip "Mark this tag for removal"
    btnTagOptions(MDTB_Copy).AssignImage "edit_copy", , buttonSize, buttonSize
    btnTagOptions(MDTB_Copy).AssignTooltip "Copy this tag to the clipboard"
    btnTagOptions(MDTB_Reset).AssignImage "generic_reset", , buttonSize, buttonSize
    btnTagOptions(MDTB_Reset).AssignTooltip "Reset tag to its original value"
    
    btnGroupOptions(MDTB_Remove).AssignImage "generic_trash", , buttonSize, buttonSize
    btnGroupOptions(MDTB_Remove).AssignTooltip "Mark this entire group for removal"
    btnGroupOptions(MDTB_Copy).AssignImage "edit_copy", , buttonSize, buttonSize
    btnGroupOptions(MDTB_Copy).AssignTooltip "Copy all tags in this group to the clipboard"
    btnGroupOptions(MDTB_Reset).AssignImage "generic_reset", , buttonSize, buttonSize
    btnGroupOptions(MDTB_Reset).AssignTooltip "Reset entire group to its original values"
    
    'Technical metadata reports are only available for images that actually exist on disk (vs clipboard or scanned images)
    If (LenB(PDImages.GetActiveImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk")) <> 0) Then
        lblTechnicalReport.Visible = True
        cmdTechnicalReport.Visible = True
    Else
        lblTechnicalReport.Visible = False
        cmdTechnicalReport.Visible = False
    End If
    
    'Give ExifTool credit for its amazing work!
    lblExifTool.Caption = g_Language.TranslateMessage("Metadata support is provided by the open-source ExifTool library.")
    
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
    
    'We now want to do something weird.  Certain hard-coded, non-editable categories should always come first.  Specifically:
    ' "System" / "File" / "ICC Profile" / "Inferred"
    ' These categories tend to be persistent across image formats, and their behavior is controlled by PhotoDemon.
    Dim targetPosition As Long
    targetPosition = 0
    
    For i = 0 To m_NumOfCategories - 1
        If Strings.StringsEqual(cNames.GetString(i), "system", True) Then
            cNames.MoveStringToNewPosition i, targetPosition
            targetPosition = targetPosition + 1
            Exit For
        End If
    Next i
    
    For i = 0 To m_NumOfCategories - 1
        If Strings.StringsEqual(cNames.GetString(i), "file", True) Then
            cNames.MoveStringToNewPosition i, targetPosition
            targetPosition = targetPosition + 1
            Exit For
        End If
    Next i
    
    For i = 0 To m_NumOfCategories - 1
        If Strings.StringsEqual(cNames.GetString(i), "icc profile", True) Then
            cNames.MoveStringToNewPosition i, targetPosition
            targetPosition = targetPosition + 1
            Exit For
        End If
    Next i
    
    For i = 0 To m_NumOfCategories - 1
        If Strings.StringsEqual(cNames.GetString(i), "inferred", True) Then
            cNames.MoveStringToNewPosition i, targetPosition
            targetPosition = targetPosition + 1
            Exit For
        End If
    Next i
    
    'We now want to sort the main category list to match this order.
    Dim tmpCat As MDCategory
    For i = 0 To m_NumOfCategories - 1
        For j = i To m_NumOfCategories - 1
            If Strings.StringsEqual(cNames.GetString(i), m_MDCategories(j).Name) And (i <> j) Then
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
        .LoadThemeColor PDMD_TagBackgroundDeleted, "TagBackgroundDeleted", RGB(255, 200, 200)
        .LoadThemeColor PDMD_TagBackgroundEdited, "TagBackgroundEdited", RGB(200, 255, 200)
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
    
    UpdateGroupButtonEnablement
    
End Sub

Private Sub UpdateGroupButtonEnablement()
    
    Dim curCategory As Long
    curCategory = lstGroup.ListIndex
    
    Dim atLeastOneTagEdited As Boolean: atLeastOneTagEdited = False
    Dim allTagsMarkedForRemoval As Boolean: allTagsMarkedForRemoval = True
    
    Dim i As Long
    For i = 0 To m_MDCategories(curCategory).Count - 1
        If (Not m_AllTags(curCategory, i).TagMarkedForRemoval) Then allTagsMarkedForRemoval = False
        If m_AllTags(curCategory, i).UserModifiedAllSessions Then atLeastOneTagEdited = True
    Next i
    
    'En/disable the group operation buttons according to the current group criteria
    btnGroupOptions(MDTB_Remove).Value = allTagsMarkedForRemoval
    btnGroupOptions(MDTB_Reset).Enabled = atLeastOneTagEdited
    
End Sub

Private Sub lstGroup_Click()
    
    'Before doing anything, store any changes to the current tag
    TagLostFocus
    
    Dim curCategory As Long
    curCategory = lstGroup.ListIndex
    
    If m_MDCategories(curCategory).Count > 0 Then
        If m_MDCategories(curCategory).Count = 1 Then
            lstMetadata.Caption = g_Language.TranslateMessage("1 tag in this category")
        Else
            lstMetadata.Caption = g_Language.TranslateMessage("%1 tags in this category", m_MDCategories(curCategory).Count)
        End If
    Else
        lstMetadata.Caption = g_Language.TranslateMessage("no tags in this category")
    End If
    
    'Beneath the group box, disply group-level buttons (delete all, reset all, etc)
    Dim topOfDescription As Long
    topOfDescription = (lstMetadata.GetTop + lstMetadata.GetHeight) - btnGroupOptions(0).GetHeight
    
    Dim i As Long
    For i = btnGroupOptions.lBound To btnGroupOptions.UBound
        btnGroupOptions(i).SetTop topOfDescription
    Next i
    
    topOfDescription = topOfDescription - FixDPI(4)
    
    'Some categories display a "helper" description
    Dim catName As String, groupDescription As String
    catName = m_MDCategories(curCategory).Name
    If Strings.StringsEqual(catName, "system", True) Then
        groupDescription = g_Language.TranslateMessage("""System"" tags are provided by the operating system.  They are not embedded as traditional metadata.")
    ElseIf Strings.StringsEqual(catName, "file", True) Then
        groupDescription = g_Language.TranslateMessage("""File"" tags are required by this image format.  They are not embedded as traditional metadata.")
    ElseIf Strings.StringsEqual(catName, "icc profile", True) Then
        groupDescription = g_Language.TranslateMessage("ICC profiles are handled automatically by PhotoDemon.  They are not embedded as traditional metadata.")
    ElseIf Strings.StringsEqual(catName, "inferred", True) Then
        groupDescription = g_Language.TranslateMessage("""Inferred"" tags are hypothetical values inferred from other metadata.  They are not embedded as traditional metadata.")
    End If
    
    'If a helper description exists, show/hide the description label to match
    If (LenB(groupDescription) = 0) Then
        lblGroupDescription.Visible = False
    Else
        lblGroupDescription.Caption = groupDescription
        lblGroupDescription.Top = topOfDescription - lblGroupDescription.GetHeight
        lblGroupDescription.Visible = True
        topOfDescription = topOfDescription - (lblGroupDescription.GetHeight + FixDPI(4))
    End If
    
    'With all UI elements beneath the group box now displayed correctly, set the final listbox height
    lstGroup.SetHeight (topOfDescription - lstGroup.GetTop)
    
    'Update the metadata list to reflect the new category
    UpdateMetadataList
    
    'We remember the last ListIndex for each category.  With the listbox successfully filled, set the new index now
    lstMetadata.ListIndex = m_MDCategories(curCategory).LastListIndex
    
    'We also remember the last group and tax index, so we can apply "lost focus" changes to edited tags
    m_GroupIndex = lstGroup.ListIndex
    m_TagIndex = lstMetadata.ListIndex
    
End Sub

Private Sub lstMetadata_Click()
        
    'Before doing anything, store any changes to the current tag
    TagLostFocus
        
    m_MDCategories(lstGroup.ListIndex).LastListIndex = lstMetadata.ListIndex
    UpdateTagView
    
    'We also remember the last group and tax index, so we can apply "lost focus" changes to edited tags
    m_GroupIndex = lstGroup.ListIndex
    m_TagIndex = lstMetadata.ListIndex
    
End Sub

Private Sub lstMetadata_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    If (bufferDC = 0) Then Exit Sub
    
    'Calculate text colors (which vary depending on selection state)
    Dim titleColor As Long, descriptionColor As Long
    If itemIsSelected Then
        titleColor = m_Colors.RetrieveColor(PDMD_TitleSelected, lstMetadata.Enabled, , itemIsHovered)
        descriptionColor = m_Colors.RetrieveColor(PDMD_DescriptionSelected, lstMetadata.Enabled, , itemIsHovered)
    Else
        titleColor = m_Colors.RetrieveColor(PDMD_TitleUnselected, lstMetadata.Enabled, , itemIsHovered)
        descriptionColor = m_Colors.RetrieveColor(PDMD_DescriptionUnselected, lstMetadata.Enabled, , itemIsHovered)
    End If
    
    'Prep various default rendering values (including retrieval of the boundary rect from the list box manager)
    Dim blockCategory As Long
    blockCategory = lstGroup.ListIndex
    
    Dim tmpRectF As RectF
    CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
    
    Dim offsetY As Single, offsetX As Single
    offsetX = tmpRectF.Left + Interface.FixDPI(8)
    offsetY = tmpRectF.Top + Interface.FixDPI(4)
    
    Dim thisTag As PDMetadataItem
    thisTag = m_AllTags(blockCategory, itemIndex)
    
    Dim linePadding As Long
    linePadding = Interface.FixDPI(3)
    
    'pd2D handles rendering duties
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundDC bufferDC
    cSurface.SetSurfaceAntialiasing P2_AA_None
    
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    
    'If the user has modified this tag, but the tag is *not* currently selected, we paint it with a different background color.
    If (Not itemIsSelected) Then
        
        'Tags marked for removal get a pink background
        If thisTag.TagMarkedForRemoval Then
            With tmpRectF
                cBrush.SetBrushColor m_Colors.RetrieveColor(PDMD_TagBackgroundDeleted, Me.Enabled)
                PD2D.FillRectangleI cSurface, cBrush, .Left, .Top, .Width, .Height + 1!
            End With
            
        'If the user has supplied their own value for this tag, we use a green background
        ElseIf (thisTag.UserModifiedAllSessions) Then
            If Strings.StringsNotEqual(thisTag.UserValueNew, thisTag.TagValueFriendly, True) Then
                With tmpRectF
                    cBrush.SetBrushColor m_Colors.RetrieveColor(PDMD_TagBackgroundEdited, Me.Enabled)
                    PD2D.FillRectangleI cSurface, cBrush, .Left, .Top, .Width, .Height + 1!
                End With
            End If
        End If
        
    End If
    
    'Note that we deliberately maintain the numerical prefix as a separate entity; we need its size (in pixels) to calculate
    ' proper padding for the description line of text, and width for the colored bar that indicates the tag's writability.
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
    cBrush.SetBrushColor tagColor
    
    With tmpRectF
        PD2D.FillRectangleI cSurface, cBrush, .Left, .Top, (offsetX - .Left) + m_TitleFont.GetWidthOfString(CStr(itemIndex + 1)) + spaceWidth + 1!, .Height + 1!
    End With
    
    '/end GDI+ rendering.  Font duties will be handled by GDI (it's faster + higher-quality)
    Set cSurface = Nothing
    
    'Start with the simplest field: the tag title (readable form)
    m_TitleFont.AttachToDC bufferDC
    m_TitleFont.SetFontColor titleColor
    m_TitleFont.FastRenderText offsetX + 0, offsetY + 0, numericalPrefix & drawString
                
    'Below the tag title, add the human-friendly description
    Dim mHeight As Single
    mHeight = m_TitleFont.GetHeightOfString(drawString) + linePadding
    m_TitleFont.ReleaseFromDC
    
    With thisTag
        
        'Trim long strings to prevent issues with DrawString
        Const MAX_DRAW_CHAR_LENGTH As Long = 64
        If .UserModifiedAllSessions Then
            If (Len(.UserValueNew) < MAX_DRAW_CHAR_LENGTH) Then drawString = .UserValueNew Else drawString = Left$(.UserValueNew, MAX_DRAW_CHAR_LENGTH)
        Else
            If (btsTechnical(1).ListIndex = 0) Then
                If (Len(.TagValueFriendly) < MAX_DRAW_CHAR_LENGTH) Then drawString = .TagValueFriendly Else drawString = Left$(.TagValueFriendly, MAX_DRAW_CHAR_LENGTH)
            Else
                If (Len(.TagValueFriendly) < MAX_DRAW_CHAR_LENGTH) Then drawString = .TagValue Else drawString = Left$(.TagValue, MAX_DRAW_CHAR_LENGTH)
            End If
        End If
    
        'List-type tags use a special delimiter (;;;).  Change this to commas to make the list a bit prettier
        If .UserModifiedAllSessions Then
            drawString = Replace$(drawString, vbCrLf, ", ", , , vbBinaryCompare)
        Else
            drawString = Replace$(drawString, ";;;", ", ", , , vbBinaryCompare)
        End If
        
    End With
    
    m_DescriptionFont.AttachToDC bufferDC
    m_DescriptionFont.SetFontColor descriptionColor
    m_DescriptionFont.FastRenderTextWithClipping offsetX + m_TitleFont.GetWidthOfString(numericalPrefix), offsetY + mHeight, (tmpRectF.Left + tmpRectF.Width) - offsetX - FixDPI(17), m_DescriptionFont.GetHeightOfString(drawString), drawString
    m_DescriptionFont.ReleaseFromDC
    
End Sub

'When a new tag is selected from the metadata list, this sub updates the right-side panel to match
Private Sub UpdateTagView()

    Dim curGroup As Long, curTag As Long
    curGroup = lstGroup.ListIndex
    curTag = lstMetadata.ListIndex
    
    If (curTag >= 0) Then
        
        'The positioning of certain controls is reflowed depending on the contents of the tag.  We track the vertical
        ' position of reflow as we go, as some elements are hidden/shown dynamically, and this affects the positioning
        ' of subsequent elements.
        Dim reflowTop As Long
        
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
                    
                    Dim targetID As String, targetValue As String
                    If .UserModifiedAllSessions Then
                        targetID = LCase$(.UserIDNew)
                        targetValue = LCase$(.UserValueNew)
                    Else
                        targetID = LCase$(.TagValue)
                        targetValue = LCase$(.TagValueFriendly)
                    End If
                    
                    Dim strID As String, strValue As String, strValueModified As String
                    
                    Dim i As Long
                    For i = 0 To .DB_StackValues.GetNumOfStrings - 1
                        
                        strID = .DB_StackIDs.GetString(i)
                        strValue = .DB_StackValues.GetString(i)
                        
                        'If this entry's ID and VALUE are identical (as they are for many XMP tags, as XMP deals solely in text),
                        ' it's redundant to list both.  For EXIF values, however, the ID may be numerical while the VALUE is
                        ' a nice, human-readable chunk of text, so display both.
                        If Strings.StringsEqual(strID, strValue, False) Then
                            lstValue.AddItem strValue
                        Else
                            strValueModified = "(" & strID & ") " & strValue
                            lstValue.AddItem strValueModified
                        End If
                        
                        'Next, we want to figure out what .ListIndex value to set.  Typically, this is the tag's initial value,
                        ' but if the user has edited this value (during this session or a previous one), we need to use that
                        ' .ListIndex instead.  This is determined prior to running the loop (e.g. see above).
                        If Strings.StringsEqual(targetID, .DB_StackIDs.GetString(i), True) Then
                            newListIndex = i
                            listIndexFound = True
                        
                        'As a failsafe, also compare the "print-friendly" version of the current value
                        Else
                            If Strings.StringsEqual(targetValue, .DB_StackValues.GetString(i), True) Then
                                newListIndex = i
                                listIndexFound = True
                            End If
                        End If
                        
                    Next i
                    
                    'If a matching list index is *not* found, add the current value (whatever it is) to the list in the
                    ' final position.
                    If (Not listIndexFound) Then
                        strValueModified = "(" & .TagValue & ") " & .TagValueFriendly
                        lstValue.AddItem strValueModified
                        lstValue.ListIndex = (lstValue.ListCount - 1)
                    Else
                        lstValue.ListIndex = newListIndex
                    End If
                    
                    lstValue.SetAutomaticRedraws True, True
                    lstValue.Visible = True
                    txtValue.Visible = False
                    reflowTop = lstValue.GetTop + lstValue.GetHeight
                
                'Any other values (text and numeric entry, among others) are handled via text box
                Else
                    lstValue.Visible = False
                    txtValue.Visible = True
                    If .UserModifiedAllSessions Then
                        txtValue.Text = .UserValueNew
                    Else
                        
                        'Before displaying, replace the default separator (;;;) with newlines
                        If (btsTechnical(1).ListIndex = 0) Then
                            txtValue.Text = Replace$(.TagValueFriendly, ";;;", vbCrLf, , , vbBinaryCompare)
                        Else
                            txtValue.Text = Replace$(.TagValue, ";;;", vbCrLf, , , vbBinaryCompare)
                        End If
                        
                    End If
                    reflowTop = txtValue.GetTop + txtValue.GetHeight
                End If
                
            Else
                lstValue.Visible = False
                txtValue.Visible = False
                lblValue.Visible = True
                
                'We still need to check for list-type values. (If found, we will replace PD's default
                ' custom separator (;;;) with newlines.)
                If (btsTechnical(1).ListIndex = 0) Then
                    lblValue.Caption = Replace$(.TagValueFriendly, ";;;", vbCrLf, , , vbBinaryCompare)
                Else
                    lblValue.Caption = Replace$(.TagValue, ";;;", vbCrLf, , , vbBinaryCompare)
                End If
                
                reflowTop = lblValue.GetTop + lblValue.GetHeight
            End If
            
            'The reflow position variable now points at the bottom of the relevant edit control (edit box, list, label, etc).
            ' Add some padding before proceeding.
            reflowTop = reflowTop + Interface.FixDPI(6)
            
            'After determining which edit control to use, we now need to determine the visibility and positioning of various
            ' warning labels, type descriptors, and other per-tag information.  As before, most of this information is
            ' contingent on the tag being writable.
            If .DB_IsWritable Then
                
                'Hard-coded lists do not display tag formatting requirements (formatting is handled silently, based on
                ' the user's list selection).
                If .DB_HardcodedList Then
                    lblTagType.Visible = False
                Else
                    
                    lblTagType.UseCustomForeColor = False
                    
                    Dim strTagRestrictions As String
                    strTagRestrictions = ConvertDataTypeToString(m_AllTags(curGroup, curTag))
                    
                    'We only list restrictions if necessary.  Generic "text" tags are treated as if they have no restrictions.
                    If (LenB(strTagRestrictions) <> 0) And Strings.StringsNotEqual(strTagRestrictions, "text", False) Then
                        lblTagType.Caption = g_Language.TranslateMessage("tag restrictions: ") & strTagRestrictions
                        lblTagType.SetTop reflowTop
                        lblTagType.Visible = True
                        
                        'Any time an element is made visible, we add its height to the running reflow offset, so subsequent
                        ' elements can be positioned correctly.
                        reflowTop = reflowTop + lblTagType.GetHeight + Interface.FixDPI(6)
                        
                    Else
                        lblTagType.Visible = False
                    End If
                    
                End If
                
                'Protected tags can technically be edited, but there may be unforeseen consequences.  Let the user know.
                If (.DBF_IsUnsafe Or .DBF_IsProtected Or .DBF_IsMandatory) Then
                    lblWarning.ForeColor = m_Colors.RetrieveColor(PDMD_TextTagEditError)
                    lblWarning.UseCustomForeColor = True
                    lblWarning.Caption = g_Language.TranslateMessage("This is a protected tag.  Edits are allowed, but they may be overwritten to produce a valid image file.")
                    lblWarning.SetTop reflowTop
                    lblWarning.Visible = True
                    
                    'Any time an element is made visible, we add its height to the running reflow offset, so subsequent
                    ' elements can be positioned correctly.
                    reflowTop = reflowTop + lblWarning.GetHeight + Interface.FixDPI(6)
                    
                Else
                    lblWarning.Visible = False
                End If
                
            Else
                
                lblTagType.Visible = False
                
                'Non-writable tags get a warning about non-editability
                lblWarning.ForeColor = m_Colors.RetrieveColor(PDMD_TextTagEditError)
                lblWarning.UseCustomForeColor = True
                lblWarning.Caption = g_Language.TranslateMessage("This tag is restricted.  It cannot be edited, and removal requests will be ignored if they result in an invalid file.")
                lblWarning.SetTop reflowTop
                lblWarning.Visible = True
                
                'Any time an element is made visible, we add its height to the running reflow offset, so subsequent
                ' elements can be positioned correctly.
                reflowTop = reflowTop + lblWarning.GetHeight + FixDPI(6)
                
            End If
            
            'All tags receive "remove this tag" and "copy this tag" options
            btnTagOptions(MDTB_Remove).SetTop reflowTop
            btnTagOptions(MDTB_Remove).Value = .TagMarkedForRemoval
            btnTagOptions(MDTB_Copy).SetTop reflowTop
            
            'Editable tags also receive a "reset this tag" button
            If .DB_IsWritable Then
                btnTagOptions(MDTB_Reset).Enabled = .UserModifiedAllSessions
                btnTagOptions(MDTB_Reset).SetTop reflowTop
                btnTagOptions(MDTB_Reset).Visible = True
            Else
                btnTagOptions(MDTB_Reset).Visible = False
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
    
    Dim isList As Boolean
    isList = srcMetadata.DBF_IsBag Or srcMetadata.DBF_IsList Or srcMetadata.DBF_IsSequence
    
    Select Case srcMetadata.DB_DataTypeStrict
    
        Case MD_int8s
            strResult = g_Language.TranslateMessage("integers only [%1 to %2]", -127, 127)
            If countPresent Then strResult = CStr(countValue) & " x " & strResult
        Case MD_int8u
            strResult = g_Language.TranslateMessage("integers only [%1 to %2]", 0, 255)
        Case MD_int16s
            strResult = g_Language.TranslateMessage("integers only [%1 to %2]", -32768, 32767)
        Case MD_int16u
            strResult = g_Language.TranslateMessage("integers only [%1 to %2]", 0, 65535)
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
            strResult = g_Language.TranslateMessage("file position marker")
        Case MD_ifd64
            strResult = g_Language.TranslateMessage("file position marker")
        Case MD_string
            strResult = g_Language.TranslateMessage("text")
        Case MD_undef
            'Debug.Print "The selected tag actually has an ""undefined"" data format, but PD displays ""text"" as a convenience."
            strResult = g_Language.TranslateMessage("text")
        Case MD_binary
            strResult = g_Language.TranslateMessage("binary data")
        Case MD_integerstring
            strResult = g_Language.TranslateMessage("digits")
        Case MD_floatstring
            strResult = g_Language.TranslateMessage("numbers only")
        Case MD_rationalstring
            strResult = g_Language.TranslateMessage("numbers only")
        Case MD_datestring
            strResult = g_Language.TranslateMessage("dates only (YYYY:mm:dd HH:MM:SS[.ss][+/-HH:MM])")
        Case MD_booleanstring
            strResult = g_Language.TranslateMessage("true or false only")
        Case MD_digits
            strResult = g_Language.TranslateMessage("digits only")
    
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
    
    If isList Then strResult = g_Language.TranslateMessage("list of %1, one entry per line", strResult)
    
    If (LenB(strResult) <> 0) Then ConvertDataTypeToString = strResult
    
End Function

'When the current tag loses focus, call this sub to update the tag's information against any user-applied edits.
Private Sub TagLostFocus(Optional ByVal redrawListToMatch As Boolean = True)
        
    If (Not m_DialogFinishedLoading) Then Exit Sub
        
    If (m_GroupIndex >= 0) And (m_TagIndex >= 0) Then
        
        Dim tagWasEdited As Boolean: tagWasEdited = False
        Dim tagStateChangedOther As Boolean: tagStateChangedOther = False
        
        With m_AllTags(m_GroupIndex, m_TagIndex)
            
            'Tag removal is handled specially.  (Specifically, note that it is unrelated to the .UserModified trackers;
            ' this is important because PhotoDemon itself may mark tags for removal, independent of the user.)
            If (.TagMarkedForRemoval) <> btnTagOptions(MDTB_Remove).Value Then
                .TagMarkedForRemoval = btnTagOptions(MDTB_Remove).Value
                tagStateChangedOther = True
            End If
            
            'There's no point checking for value changes if a tag is un-editable
            If .DB_IsWritable Then
            
                'Detecting a value change varies by edit type.  Text box entries can be compared pretty easily (just a StrComp),
                ' while list-box changes are a bit more convoluted.
                
                'Values that are part of a hardcoded list are available via dropdown, and the dropdown contains both the ID
                ' and the friendly text, crammed into one.  As such, we can't just compare listbox text.
                If .DB_HardcodedList And (.DB_TypeCount < 2) Then
                    
                    'Start by detecting which ListIndex corresponds to the tag's original value.
                    Dim foundDefaultListIndex As Boolean: foundDefaultListIndex = False
                    
                    Dim i As Long
                    For i = 0 To .DB_StackValues.GetNumOfStrings - 1
                        
                        If Strings.StringsEqual(.TagValue, .DB_StackIDs.GetString(i), True) Then
                            foundDefaultListIndex = True

                        'As a failsafe, also compare the "print-friendly" version of the current value
                        Else
                            If Strings.StringsEqual(.TagValueFriendly, .DB_StackValues.GetString(i), True) Then
                                foundDefaultListIndex = True
                            End If
                        End If
                        
                        If foundDefaultListIndex Then
                        
                            '"i" is the .ListIndex of the tag's original value.
                            If (lstValue.ListIndex <> i) Then
                                
                                'This tag has been edited
                                tagWasEdited = True
                                .UserIDNew = .DB_StackIDs.GetString(lstValue.ListIndex)
                                .UserValueNew = .DB_StackValues.GetString(lstValue.ListIndex)
                                
                            End If
                            
                            Exit For
                        
                        End If

                    Next i
                
                'Any other values (text and numeric entry, among others) are handled via text box
                Else
                    
                    Dim testString As String
                    
                    'List-type tags require a special check, as we will have forcibly converted them
                    ' from a special delimiter state (;;;) to newlines, for the user's convenience
                    If (btsTechnical(1).ListIndex = 0) Then
                        testString = Replace$(.TagValueFriendly, ";;;", vbCrLf, , , vbBinaryCompare)
                    Else
                        testString = Replace$(.TagValue, ";;;", vbCrLf, , , vbBinaryCompare)
                    End If
                    
                    If Strings.StringsNotEqual(txtValue.Text, testString, True) Then
                        
                        'This string is different from the one we placed inside.  Mark the tag as edited, and store the
                        ' user-supplied value.
                        tagWasEdited = True
                        .UserValueNew = txtValue.Text
                    
                    End If
                    
                End If
            '/End "is tag writable?"
            End If
            
            'We track two "was this tag edited?" values: one for this session (which is reset every time the metadata editor
            ' is loaded), and one for ALL sessions (once it becomes TRUE, it stays TRUE until PD exits).
            If tagWasEdited Then
                .UserModifiedThisSession = True
                .UserModifiedAllSessions = True
            End If
            
            'State changes require us to repaint the metadata list box, as the changes should be immediately reflected.
            If (tagWasEdited Or tagStateChangedOther) Then
                If redrawListToMatch Then lstMetadata.SetAutomaticRedraws True, True
                UpdateGroupButtonEnablement
            End If
            
        End With
    
    End If
    
End Sub

Private Sub lstValue_Click()
    btnTagOptions(MDTB_Reset).Enabled = True
End Sub

Private Sub txtValue_Change()
    btnTagOptions(MDTB_Reset).Enabled = True
End Sub
