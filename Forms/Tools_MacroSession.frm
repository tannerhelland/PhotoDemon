VERSION 5.00
Begin VB.Form FormMacroSession 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Create macro from session history"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11790
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
   ScaleHeight     =   503
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   786
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdLabel lblSummary 
      Height          =   495
      Left            =   360
      Top             =   6000
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2143
      Alignment       =   2
      Caption         =   ""
      FontBold        =   -1  'True
      FontSize        =   11
   End
   Begin PhotoDemon.pdListBoxOD lstStart 
      Height          =   5055
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8916
      Caption         =   "first action"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   1296
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdListBoxOD lstEnd 
      Height          =   5055
      Left            =   6120
      TabIndex        =   2
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8916
      Caption         =   "last action"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   255
      Index           =   0
      Left            =   360
      Top             =   5400
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      Caption         =   "* current image state"
      FontItalic      =   -1  'True
   End
End
Attribute VB_Name = "FormMacroSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Create macro from session history tool
'Copyright 2018-2026 by Tanner Helland
'Created: 18/June/18
'Last updated: 21/June/18
'Last update: wrap up initial build
'
'In https://github.com/tannerhelland/PhotoDemon/issues/265, jpbro provided a great suggestion - that for many users,
' it would be easier to create macros retroactively, from some set of steps they have already completed.  His idea
' was to use the Undo History as a starting point for a "create macro from history" tool, and sure enough, that's
' exactly what this dialog does.
'
'The main UI code is lifted almost entirely from the Undo History window, with a few minor modifications to make
' its unique rules clearer (e.g. graying out actions in the "final action" list that occurred before the currently
' selected action in the "first action" list).
'
'Macros saved from this tool will be byte-for-byte identical to ones created via standard recording, because the
' Macro engine itself is still used for writing the macro files.  (We just pass it an "artificial" list of
' recorded actions, created from the selected history points.)
'
'For additional discussion and testing, please refer to https://github.com/tannerhelland/PhotoDemon/issues/265.
' Thank you again to jpbro for his suggestion and subsequent testing.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This array contains the contents of the current Undo stack, as copied from the pdUndo class.  We use it to
' generate the list of operations applied during this session.
Private m_undoEntries() As PD_UndoEntry

'Total number of Undo entries, and index of the current Undo entry (e.g. the current image state in the
' undo/redo chain).
Private m_numOfUndos As Long, m_curUndoIndex As Long

'Height of each content block in the custom-drawn listboxes
Private Const BLOCKHEIGHT As Long = 58

'Two font objects; one for action names and one for action descriptions.  (Two are needed because they have
' different sizes and colors, and it is faster to cache the associated fonts rather than constantly recreating
' them through a single pdFont object.)
Private m_TitleFont As pdFont, m_DescriptionFont As pdFont

'The size used for rendering action thumbnail images
Private Const UNDO_THUMB_SMALL As Long = 48

'Certain types of actions (e.g. selections, operations on layers) may need additional information to be helpful.
' This helper function provides generic helper text for a given action type.
Private Function GetStringForUndoType(ByVal typeOfUndo As PD_UndoType, Optional ByVal layerID As Long = 0) As String

    Dim newText As String
    
    Select Case typeOfUndo
    
        Case UNDO_Everything
            newText = vbNullString
            
        Case UNDO_Image, UNDO_Image_VectorSafe, UNDO_ImageHeader
            newText = vbNullString
            
        Case UNDO_Layer, UNDO_Layer_VectorSafe, UNDO_LayerHeader
            If Not (PDImages.GetActiveImage.GetLayerByID(layerID) Is Nothing) Then
                newText = g_Language.TranslateMessage("layer: %1", PDImages.GetActiveImage.GetLayerByID(layerID).GetLayerName())
            Else
                newText = vbNullString
            End If
        
        Case UNDO_Selection
            newText = g_Language.TranslateMessage("selection shape shown")
        
    End Select
    
    GetStringForUndoType = newText

End Function

Private Sub cmdBar_OKClick()
    
    'Make sure the selected positions are valid.
    If (lstEnd.ListIndex >= lstStart.ListIndex) Then
        
        'Perform a second check to make sure the user hasn't just selected the "original image" index
        If (lstEnd.ListIndex > 0) Then
        
            'The macro module handles the actual dialog raising for us
            Dim dstFile As String
            If Macros.DisplayMacroSaveDialog(dstFile) Then
            
                'Copy the relevant processor data into a dedicated PD_ProcessCall array, which we will then pass
                ' to the macro module for further handling.
                Dim finalMacro() As PD_ProcessCall, numMacros As Long
                numMacros = (lstEnd.ListIndex - lstStart.ListIndex) + 1
                
                ReDim finalMacro(0 To numMacros - 1) As PD_ProcessCall
                
                Dim i As Long
                For i = 0 To numMacros - 1
                    finalMacro(i) = m_undoEntries(lstStart.ListIndex + i).srcProcCall
                Next i
                
                If ExportProcCallsToMacroFile(dstFile, finalMacro, 0, numMacros - 1) Then
                    Message "Macro saved successfully."
                    'Unload will occur automatically, c/o the command bar user control that raised this event
                End If
                
            End If
            
        Else
            PDMsgBox "WARNING: this is not a valid macro.  You must include at least one action that modifies the image.", vbExclamation Or vbOKOnly Or vbApplicationModal, "Invalid macro"
            cmdBar.DoNotUnloadForm
        End If
        
    Else
        PDMsgBox "WARNING: this is not a valid macro.  The last action cannot occur after the first action.", vbExclamation Or vbOKOnly Or vbApplicationModal, "Invalid macro"
        cmdBar.DoNotUnloadForm
    End If
    
End Sub

Private Sub Form_Load()
    
    'Initialize a custom font object for action names
    Set m_TitleFont = New pdFont
    m_TitleFont.SetFontBold True
    m_TitleFont.SetFontSize 12
    m_TitleFont.CreateFontObject
    m_TitleFont.SetTextAlignment vbLeftJustify
    
    '...and a second custom font object for action descriptions
    Set m_DescriptionFont = New pdFont
    m_DescriptionFont.SetFontBold False
    m_DescriptionFont.SetFontSize 10
    m_DescriptionFont.CreateFontObject
    m_DescriptionFont.SetTextAlignment vbLeftJustify
    
    'Retrieve a copy of all session Undo data from the current image's undo manager
    PDImages.GetActiveImage.UndoManager.CopyUndoStack m_numOfUndos, m_curUndoIndex, m_undoEntries
    
    'Populate the owner-drawn listboxes with copies of the retrieved action lists (including thumbnails)
    Dim i As Long
    
    lstStart.ListItemHeight = Interface.FixDPI(BLOCKHEIGHT)
    lstStart.SetAutomaticRedraws False
    For i = 0 To m_numOfUndos - 1
        lstStart.AddItem , i
    Next i
    lstStart.SetAutomaticRedraws True, True
    If (lstStart.ListCount > 1) Then lstStart.ListIndex = 1 Else lstStart.ListIndex = 0
    
    lstEnd.ListItemHeight = Interface.FixDPI(BLOCKHEIGHT)
    lstEnd.SetAutomaticRedraws False
    For i = 0 To m_numOfUndos - 1
        lstEnd.AddItem , i
    Next i
    lstEnd.SetAutomaticRedraws True, True
    lstEnd.ListIndex = lstEnd.ListCount - 1
    
    'Update the summary report of the default macro
    UpdateSummary
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lstEnd_Click()
    UpdateSummary
End Sub

Private Sub lstEnd_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)

    If (bufferDC = 0) Then Exit Sub
    
    'Retrieve the boundary region for this list entry
    Dim tmpRectF As RectF
    CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
    
    Dim offsetY As Single, offsetX As Single
    offsetX = tmpRectF.Left
    offsetY = tmpRectF.Top + Interface.FixDPI(2)
        
    Dim linePadding As Long
    linePadding = Interface.FixDPI(2)
    
    Dim mHeight As Single
        
    'If this filter has been selected, draw the background with the system's current selection color
    If itemIsSelected Then
        m_TitleFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected)
        m_DescriptionFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected)
    Else
        Dim itemIsValid As Boolean
        itemIsValid = (itemIndex >= lstStart.ListIndex)
        m_TitleFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, itemIsValid, , itemIsHovered)
        m_DescriptionFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, itemIsValid, , itemIsHovered)
    End If
    
    'Prepare a title string (with an asterisk added to the "current" image state title)
    Dim drawString As String
    If (itemIndex + 1) = m_curUndoIndex Then drawString = "* "
    drawString = drawString & CStr(itemIndex + 1) & " - " & g_Language.TranslateMessage(m_undoEntries(itemIndex).srcProcCall.pcID)
    
    'Render the thumbnail for this entry, and note that the thumbnail is *not* guaranteed to be square.
    Dim thumbNewWidth As Long, thumbNewHeight As Long, thumbMax As Long
    thumbMax = Interface.FixDPI(UNDO_THUMB_SMALL)
    PDMath.ConvertAspectRatio m_undoEntries(itemIndex).thumbnailLarge.GetDIBWidth, m_undoEntries(itemIndex).thumbnailLarge.GetDIBHeight, thumbMax, thumbMax, thumbNewWidth, thumbNewHeight
    GDI_Plus.GDIPlus_StretchBlt Nothing, offsetX + Interface.FixDPI(4) + (thumbMax - thumbNewWidth) \ 2, offsetY + (Interface.FixDPI(BLOCKHEIGHT) - thumbMax) \ 2 + (thumbMax - thumbNewHeight) \ 2, thumbNewWidth, thumbNewHeight, m_undoEntries(itemIndex).thumbnailLarge, 0, 0, m_undoEntries(itemIndex).thumbnailLarge.GetDIBWidth, m_undoEntries(itemIndex).thumbnailLarge.GetDIBHeight, , , bufferDC
    
    'Figure out how much space the thumbnail has taken; we'll shift text to the left of this
    Dim thumbWidth As Long
    thumbWidth = offsetX + Interface.FixDPI(4) + Interface.FixDPI(UNDO_THUMB_SMALL)
    
    'Render the title text
    If (LenB(drawString) <> 0) Then
        m_TitleFont.AttachToDC bufferDC
        m_TitleFont.FastRenderText thumbWidth + Interface.FixDPI(16) + offsetX, offsetY + Interface.FixDPI(4), drawString
        m_TitleFont.ReleaseFromDC
    End If
            
    'Below that, add the description text (if any)
    drawString = GetStringForUndoType(m_undoEntries(itemIndex).srcProcCall.pcUndoType, m_undoEntries(itemIndex).undoLayerID)
    
    If (LenB(drawString) <> 0) Then
        mHeight = m_TitleFont.GetHeightOfString(drawString) + linePadding
        m_DescriptionFont.AttachToDC bufferDC
        m_DescriptionFont.FastRenderText thumbWidth + Interface.FixDPI(16) + offsetX, offsetY + Interface.FixDPI(4) + mHeight, drawString
        m_DescriptionFont.ReleaseFromDC
    End If
    
End Sub

Private Sub lstStart_Click()
    
    'Whenever the starting macro changes, we want to redraw the final macro list so we can
    ' "dim" invalid final macro choices.
    lstEnd.SetAutomaticRedraws False
    
    'Try to prevent the user from creating invalid final macros.
    If (lstEnd.ListIndex < lstStart.ListIndex) Then lstEnd.ListIndex = lstStart.ListIndex
    
    lstEnd.SetAutomaticRedraws True, True
    
    UpdateSummary
    
End Sub

Private Sub lstStart_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    If (bufferDC = 0) Then Exit Sub
    
    'Retrieve the boundary region for this list entry
    Dim tmpRectF As RectF
    CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
    
    Dim offsetY As Single, offsetX As Single
    offsetX = tmpRectF.Left
    offsetY = tmpRectF.Top + Interface.FixDPI(2)
        
    Dim linePadding As Long
    linePadding = Interface.FixDPI(2)
    
    Dim mHeight As Single
        
    'If this filter has been selected, draw the background with the system's current selection color
    If itemIsSelected Then
        m_TitleFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected)
        m_DescriptionFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected)
    Else
        m_TitleFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, , , itemIsHovered)
        m_DescriptionFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, , , itemIsHovered)
    End If
    
    'Prepare a title string (with an asterisk added to the "current" image state title)
    Dim drawString As String
    If (itemIndex + 1) = m_curUndoIndex Then drawString = "* "
    drawString = drawString & CStr(itemIndex + 1) & " - " & g_Language.TranslateMessage(m_undoEntries(itemIndex).srcProcCall.pcID)
    
    'Render the thumbnail for this entry, and note that the thumbnail is *not* guaranteed to be square.
    Dim thumbNewWidth As Long, thumbNewHeight As Long, thumbMax As Long
    thumbMax = Interface.FixDPI(UNDO_THUMB_SMALL)
    PDMath.ConvertAspectRatio m_undoEntries(itemIndex).thumbnailLarge.GetDIBWidth, m_undoEntries(itemIndex).thumbnailLarge.GetDIBHeight, thumbMax, thumbMax, thumbNewWidth, thumbNewHeight
    GDI_Plus.GDIPlus_StretchBlt Nothing, offsetX + Interface.FixDPI(4) + (thumbMax - thumbNewWidth) \ 2, offsetY + (Interface.FixDPI(BLOCKHEIGHT) - thumbMax) \ 2 + (thumbMax - thumbNewHeight) \ 2, thumbNewWidth, thumbNewHeight, m_undoEntries(itemIndex).thumbnailLarge, 0, 0, m_undoEntries(itemIndex).thumbnailLarge.GetDIBWidth, m_undoEntries(itemIndex).thumbnailLarge.GetDIBHeight, , , bufferDC
    
    'Figure out how much space the thumbnail has taken; we'll shift text to the left of this
    Dim thumbWidth As Long
    thumbWidth = offsetX + Interface.FixDPI(4) + Interface.FixDPI(UNDO_THUMB_SMALL)
    
    'Render the title text
    If (LenB(drawString) <> 0) Then
        m_TitleFont.AttachToDC bufferDC
        m_TitleFont.FastRenderText thumbWidth + Interface.FixDPI(16) + offsetX, offsetY + Interface.FixDPI(4), drawString
        m_TitleFont.ReleaseFromDC
    End If
    
    'Below that, add the description text (if any)
    drawString = GetStringForUndoType(m_undoEntries(itemIndex).srcProcCall.pcUndoType, m_undoEntries(itemIndex).undoLayerID)
    
    If (LenB(drawString) <> 0) Then
        mHeight = m_TitleFont.GetHeightOfString(drawString) + linePadding
        m_DescriptionFont.AttachToDC bufferDC
        m_DescriptionFont.FastRenderText thumbWidth + Interface.FixDPI(16) + offsetX, offsetY + Interface.FixDPI(4) + mHeight, drawString
        m_DescriptionFont.ReleaseFromDC
    End If
        
End Sub

Private Sub UpdateSummary()
    If (lstEnd.ListIndex >= lstStart.ListIndex) Then
        If (lstEnd.ListIndex > 0) Then
            lblSummary.Caption = g_Language.TranslateMessage("Your final macro will contain %1 action(s).", (lstEnd.ListIndex - lstStart.ListIndex) + 1)
        Else
            lblSummary.Caption = g_Language.TranslateMessage("WARNING: this is not a valid macro.  You must include at least one action that modifies the image.")
        End If
    Else
        lblSummary.Caption = g_Language.TranslateMessage("WARNING: this is not a valid macro.  The last action cannot occur after the first action.")
    End If
End Sub
