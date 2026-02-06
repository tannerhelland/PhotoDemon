VERSION 5.00
Begin VB.Form dialog_UnsavedChanges 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Unsaved changes"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10110
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
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
   ScaleHeight     =   308
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   674
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdDropDown cboImageNames 
      Height          =   795
      Left            =   105
      TabIndex        =   4
      Top             =   3810
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1402
      Caption         =   "images with unsaved changes:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdPictureBox picWarning 
      Height          =   855
      Left            =   3915
      Top             =   240
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   840
      Left            =   4920
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1482
      Caption         =   ""
      Layout          =   1
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   0
      Left            =   3960
      TabIndex        =   0
      Top             =   1260
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   1296
      Caption         =   "Save the image before closing it"
   End
   Begin PhotoDemon.pdCheckBox chkRepeat 
      Height          =   330
      Left            =   3945
      TabIndex        =   3
      Top             =   4215
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   582
      Caption         =   "Repeat this action for all unsaved images (%1 in total)"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   3495
      Left            =   120
      Top             =   120
      Width           =   3495
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   1
      Left            =   3960
      TabIndex        =   1
      Top             =   2070
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   1296
      Caption         =   "Do not save the image (discard all changes)"
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   2880
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   1296
      Caption         =   "Cancel, and return to editing"
   End
End
Attribute VB_Name = "dialog_UnsavedChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Unsaved Changes Dialog
'Copyright 2011-2026 by Tanner Helland
'Created: 13/November/12
'Last updated: 01/December/12
'Last update: removed the DrawSystemIcon sub; now it can be found in the "Drawing" module
'
'Custom dialog box for warning the user that they are about to close an image with unsaved changes.
'
'This form was built after much usability testing.  There are many bad ways to design a save prompt,
' and only a few good ones.  I felt that descriptive icons were necessary to help the user quickly
' determine what choice to make.  A preview of the image in question is also displayed, to make it
' absolutely certain that the user is not confused about which image they're dealing with.  (This is
' important for photos from a digital camera, which often have names like "1004701.jpg".) Very
' descriptive tooltip text has also been added, and I genuinely believe that this is one of the best
' unsaved changes dialogs available.
'
'Finally, note that this prompt can be turned off completely from the Edit -> Preferences menu.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

'The ID number of the primary image being closed.  (This is generally the currently active image.)
Private m_imageBeingClosed As Long

'If the user is closing multiple images, these values will contain additional information.
Private m_numOfUnsavedImages As Long

'If the user is closing multiple images, this stack will contain all image indices
Private m_unsavedImageIDs As pdStack

'The user input from the dialog
Private m_userAnswer As VbMsgBoxResult

'Theme-specific icons are fully supported
Private m_warningDIB As pdDIB

'Current preview image (thumbnail)
Private m_PreviewDIB As pdDIB

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_userAnswer
End Property

'The ShowDialog routine presents the dialog to the user.  Other passed variables help us create a more detailed description
' of how many images are unsaved, and what the repercussions are for exiting now.
Public Sub ShowDialog(ByVal srcImageID As Long, ByVal numOfUnsavedImages As Long, ByRef unsavedImageIDs As pdStack)
    
    'Before displaying the "do you want to save this image?" dialog, bring the image in question to the foreground.
    If FormMain.Enabled Then CanvasManager.ActivatePDImage srcImageID, "unsaved changes dialog", True
    
    m_imageBeingClosed = srcImageID
    m_numOfUnsavedImages = numOfUnsavedImages
    If (Not unsavedImageIDs Is Nothing) Then
        Set m_unsavedImageIDs = New pdStack
        m_unsavedImageIDs.CloneStack unsavedImageIDs
    End If
    
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    Dim buttonIconSize As Long
    buttonIconSize = Interface.FixDPI(26)
    cmdAnswer(0).AssignImage "file_save", , buttonIconSize, buttonIconSize
    cmdAnswer(1).AssignImage "file_close", , buttonIconSize, buttonIconSize, g_Themer.GetGenericUIColor(UI_ErrorRed)
    cmdAnswer(2).AssignImage "edit_undo", , buttonIconSize, buttonIconSize
        
    'If the image has been saved before, update the tooltip text on the "Save" button accordingly
    If (LenB(PDImages.GetImageByID(m_imageBeingClosed).ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) <> 0) Then
        cmdAnswer(0).AssignTooltip "NOTE: if you click 'Save', PhotoDemon will save this image using its current file name." & vbCrLf & vbCrLf & "If you want to save it with a different file name, please select 'Cancel', then use the File -> Save As menu item."
    Else
        cmdAnswer(0).AssignTooltip "Because this image has not been saved before, you will be prompted to provide a file name for it."
    End If
    
    cmdAnswer(1).AssignTooltip "If you do not save this image, any changes you have made will be permanently lost."
    cmdAnswer(2).AssignTooltip "Canceling will return you to the main PhotoDemon window."
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_userAnswer = vbCancel
    
    'Adjust the save message to match this image's name
    Dim imageName As String
    imageName = PDImages.GetImageByID(m_imageBeingClosed).ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    If (LenB(Trim$(imageName)) = 0) Then imageName = g_Language.TranslateMessage("This image")
    lblWarning.Caption = g_Language.TranslateMessage("%1 has unsaved changes.  What would you like to do?", """" & imageName & """")
    lblWarning.RequestRefresh
    
    'Make some measurements of the form size.  We need these if we choose to display the check box at the bottom of the form
    If (Not g_WindowManager Is Nothing) Then
    
        Dim vDifference As Long, wRect As winRect, wRectClient As winRect
        g_WindowManager.GetWindowRect_API Me.hWnd, wRect
        g_WindowManager.GetClientWinRect Me.hWnd, wRectClient
        vDifference = (wRect.y2 - wRect.y1) - (wRectClient.y2 - wRectClient.y1)
        
        'If there are multiple unsaved images, give the user a prompt to apply this action to all of them.
        ' (If there are not multiple unsaved images, hide that section from view.)
        If (m_numOfUnsavedImages <= 1) Then
            chkRepeat.Visible = False
            g_WindowManager.SetSizeByHWnd Me.hWnd, wRect.x2 - wRect.x1, vDifference + picPreview.GetHeight + (picPreview.GetTop * 2)
        Else
            
            chkRepeat.Visible = True
            
            'If multiple images are being unloaded, and their IDs were successfully passed, load their names into the dropdown
            ' box now.
            If (Not m_unsavedImageIDs Is Nothing) Then
                
                cboImageNames.Visible = True
                cboImageNames.SetAutomaticRedraws False
                
                Dim numUnnamedImages As Long
                Dim tmpImgName As String
                
                Dim i As Long
                For i = 0 To m_unsavedImageIDs.GetNumOfInts - 1
                    tmpImgName = PDImages.GetImageByID(m_unsavedImageIDs.GetInt(i)).ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
                    If (LenB(tmpImgName) <> 0) Then
                        cboImageNames.AddItem tmpImgName, i
                    Else
                        cboImageNames.AddItem g_Language.TranslateMessage("unnamed image %1", numUnnamedImages + 1)
                        numUnnamedImages = numUnnamedImages + 1
                    End If
                Next i
                
                'Find the image ID that matches the active image, and set the default list index to that.
                For i = 0 To m_unsavedImageIDs.GetNumOfInts - 1
                    If (m_unsavedImageIDs.GetInt(i) = m_imageBeingClosed) Then
                        cboImageNames.ListIndex = i
                        Exit For
                    End If
                Next i
                
                cboImageNames.SetAutomaticRedraws True, True
            
            End If
            
            'Change the text of the "repeat for all unsaved images" check box depending on how many unsaved images are present.
            If (m_numOfUnsavedImages = 2) Then
                chkRepeat.Caption = g_Language.TranslateMessage("Repeat this action for both unsaved images")
            Else
                chkRepeat.Caption = g_Language.TranslateMessage("Repeat this action for all unsaved images (%1 in total)", m_numOfUnsavedImages)
            End If
            
            g_WindowManager.SetSizeByHWnd Me.hWnd, wRect.x2 - wRect.x1, vDifference + (cboImageNames.GetTop + cboImageNames.GetHeight) + picPreview.GetTop + Interface.FixDPI(8)
            
        End If
    
    End If
    
    'Apply any custom styles to the form
    Interface.ApplyThemeAndTranslations Me
    
    'Prep a warning icon
    Dim warningIconSize As Long
    warningIconSize = Interface.FixDPI(32)
    
    If IconsAndCursors.LoadResourceToDIB("generic_warning", m_warningDIB, warningIconSize, warningIconSize, 0) Then
        picWarning.RequestRedraw True
    Else
        Set m_warningDIB = Nothing
        picWarning.Visible = False
    End If
    
    'Prep the unsaved changes preview
    If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
    If PDImages.IsImageActive(m_imageBeingClosed) Then
        PDImages.GetImageByID(m_imageBeingClosed).RequestThumbnail m_PreviewDIB, IIf(picPreview.GetWidth > picPreview.GetHeight, picPreview.GetHeight - 2, picPreview.GetWidth - 2), False
    End If
    picPreview.RequestRedraw True
    
    'Display the form
    ShowPDDialog vbModal, Me, True

End Sub

'Before this dialog closes, this routine is called to update the user's preference for applying this action to all unsaved images
Private Sub UpdateRepeatToAllUnsavedImages(ByVal actionToApply As VbMsgBoxResult)
    g_DealWithAllUnsavedImages = (chkRepeat.Visible And chkRepeat.Value)
    If g_DealWithAllUnsavedImages Then g_HowToDealWithAllUnsavedImages = actionToApply
End Sub

Private Sub cboImageNames_Click()

    If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
    If PDImages.IsImageActive(m_unsavedImageIDs.GetInt(cboImageNames.ListIndex)) Then
        PDImages.GetImageByID(m_unsavedImageIDs.GetInt(cboImageNames.ListIndex)).RequestThumbnail m_PreviewDIB, IIf(picPreview.GetWidth > picPreview.GetHeight, picPreview.GetHeight - 2, picPreview.GetWidth - 2), False
    End If
    picPreview.RequestRedraw True

End Sub

'The three choices available to the user correspond to message box responses of "Yes", "No", and "Cancel"
Private Sub cmdAnswer_Click(Index As Integer)

    Select Case Index
    
        Case 0
            m_userAnswer = vbYes
        
        Case 1
            m_userAnswer = vbNo
            
        Case 2
            m_userAnswer = vbCancel
        
    End Select
    
    UpdateRepeatToAllUnsavedImages m_userAnswer
    Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_PreviewDIB Is Nothing) Then picPreview.CopyDIB m_PreviewDIB, , True, , True
End Sub

Private Sub picWarning_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    GDI.FillRectToDC targetDC, 0, 0, ctlWidth, ctlHeight, g_Themer.GetGenericUIColor(UI_Background)
    If (Not m_warningDIB Is Nothing) Then m_warningDIB.AlphaBlendToDC targetDC, , (ctlWidth - m_warningDIB.GetDIBWidth) \ 2, (ctlHeight - m_warningDIB.GetDIBHeight) \ 2
End Sub
