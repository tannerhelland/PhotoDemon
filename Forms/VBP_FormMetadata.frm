VERSION 5.00
Begin VB.Form FormMetadata 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Browse image metadata"
   ClientHeight    =   7140
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
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.smartCheckBox chkFriendlyNames 
      Height          =   360
      Left            =   3240
      TabIndex        =   6
      Top             =   5760
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   635
      Caption         =   "use readable names"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.VScrollBar vsMetadata 
      Height          =   5340
      LargeChange     =   32
      Left            =   11430
      TabIndex        =   5
      Top             =   240
      Width           =   330
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   3360
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   537
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.ListBox lstMetadata 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   5340
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10380
      TabIndex        =   1
      Top             =   6510
      Width           =   1365
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   8880
      TabIndex        =   0
      Top             =   6510
      Width           =   1365
   End
   Begin PhotoDemon.smartCheckBox chkFriendlyValues 
      Height          =   360
      Left            =   7320
      TabIndex        =   7
      Top             =   5760
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   635
      Caption         =   "use readable values"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -120
      TabIndex        =   2
      Top             =   6360
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
'Last updated: 31/May/14
'Last update: convert mousewheel interactions to pdInput
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
Private Const BLOCKHEIGHT As Long = 64

'An outside class provides access to mousewheel events for scrolling the filter view
Private WithEvents cMouseEvents As pdInput
Attribute cMouseEvents.VB_VarHelpID = -1

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Because we technically have two scrollable controls on this page, we must differentiate between the two when a mousewheel event is fired.
' By default, Windows will not pass scrollwheel events to the parent form (per my testing?) when a list box on that form has focus.  If
' the left list box is clicked, this value will be set to TRUE.  If the mouse leaves the list box, we automatically set this to FALSE.
Private mouseOverMetadataCategoryBox As Boolean

Private Sub chkFriendlyNames_Click()
    redrawMetadataList
End Sub

Private Sub chkFriendlyValues_Click()
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
    
    'Realign the bottom check boxes
    chkFriendlyValues.Left = chkFriendlyNames.Left + chkFriendlyNames.Width + fixDPI(24)
    
End Sub

'LOAD dialog
Private Sub Form_Load()

    'Note that this form will be interacting heavily with the current image's metadata container.
    
    'Enable mousewheel scrolling for the metadata box
    Set cMouseEvents = New pdInput
    cMouseEvents.addInputTracker picBuffer.hWnd, True, , , True
    cMouseEvents.addInputTracker Me.hWnd
    cMouseEvents.setSystemCursor IDC_ARROW
        
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
    lstMetadata.Clear
    
    For i = 0 To numOfCategories - 1
        lstMetadata.AddItem mdCategories(i).Name & " (" & mdCategories(i).Count & ")", i
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
    
    'Set the height of the picture box buffer to be the same as the category list box
    picBuffer.Height = lstMetadata.Height
    
    lstMetadata.ListIndex = 0
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'When the mouse enters the form area, set focus to the vertical scroll bar, so we can automatically mousewheel it (instead of the
    ' left listbox)
    If mouseOverMetadataCategoryBox Then
        mouseOverMetadataCategoryBox = False
        If vsMetadata.Visible Then vsMetadata.SetFocus
    End If
    
End Sub

'UNLOAD form
Private Sub Form_Unload(Cancel As Integer)
    
    'Unload the mouse tracker
    Set cMouseEvents = Nothing
    ReleaseFormTheming Me

End Sub

'When a new metadata category is selected, redraw all the metadata text currently on screen
Private Sub lstMetadata_Click()
    
    Dim curCategory As Long
    curCategory = lstMetadata.ListIndex
        
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

'Redraw the full metadata list
Private Sub redrawMetadataList()

    Dim curCategory As Long
    curCategory = lstMetadata.ListIndex

    picBuffer.Picture = LoadPicture("")
        
    Dim scrollOffset As Long
    scrollOffset = vsMetadata.Value
        
    Dim i As Long
    For i = 0 To mdCategories(curCategory).Count - 1
        renderMDBlock curCategory, i, fixDPI(8), fixDPI(i * BLOCKHEIGHT) - scrollOffset - fixDPI(2)
    Next i
    
    'Copy the buffer to the main form
    picBuffer.Picture = picBuffer.Image
    Me.PaintPicture picBuffer.Picture, lstMetadata.Width + lstMetadata.Left * 2, lstMetadata.Top, picBuffer.ScaleWidth, picBuffer.ScaleHeight, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight

End Sub

'Render the given metadata index onto the background picture box at the specified offset
Private Sub renderMDBlock(ByVal blockCategory As Long, ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + BLOCKHEIGHT) > 0) And (offsetY < picBuffer.Height) Then

        Dim thisTag As mdItem
        thisTag = allTags(blockCategory, blockIndex)
    
        Dim primaryColor As Long, secondaryColor As Long, tertiaryColor As Long
        primaryColor = RGB(64, 64, 64)
        secondaryColor = RGB(92, 92, 92)
        tertiaryColor = vbActiveTitleBar
    
        Dim linePadding As Long
        linePadding = fixDPI(4)
    
        Dim mHeight As Single
        Dim drawString As String
        
        If CBool(chkFriendlyNames.Value) Then
            drawString = thisTag.Description
        Else
            drawString = thisTag.FullGroupAndName
        End If
    
        'Start with the simplest field: the tag title (readable form)
        drawTextOnObject picBuffer, Str(blockIndex + 1) & " - " & drawString, offsetX + 0, offsetY + 0, 12, primaryColor, True, False
                
        'Below the tag title, add the human-friendly description
        mHeight = picBuffer.TextHeight(drawString) + linePadding
        
        If CBool(chkFriendlyValues.Value) Then
            drawString = thisTag.Value
        Else
            If Len(thisTag.ActualValue) > 0 Then
                drawString = thisTag.ActualValue
            Else
                drawString = thisTag.Value
            End If
        End If
        
        drawTextOnObject picBuffer, drawString, offsetX + fixDPI(4), offsetY + mHeight, 11, secondaryColor, False
        
        'Draw a divider line near the bottom of the metadata block
        Dim lineY As Long
        If blockIndex < mdCategories(blockCategory).Count - 1 Then
            lineY = offsetY + fixDPI(BLOCKHEIGHT - 8)
            picBuffer.Line (fixDPI(4), lineY)-(picBuffer.ScaleWidth - fixDPI(8), lineY), tertiaryColor
        End If
        
    End If

End Sub

'To solve the problem of "two mouse-scrollable lists on one dialog), we track the mouse's location and handle scroll events accordingly.
Private Sub lstMetadata_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseOverMetadataCategoryBox = True
End Sub

'When the scrollbar is moved, redraw the metadata list
Private Sub vsMetadata_Change()
    redrawMetadataList
End Sub

Private Sub vsMetadata_Scroll()
    redrawMetadataList
End Sub
