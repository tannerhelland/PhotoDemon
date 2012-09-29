VERSION 5.00
Begin VB.Form FormImage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Image Window"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VBP_FormImage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.VScrollBar VScroll 
      Height          =   3615
      LargeChange     =   10
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.PictureBox PicCH 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3480
      Picture         =   "VBP_FormImage.frx":000C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "FormImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Form (Child MDI form)
'Copyright ©2002-2012 by Tanner Helland
'Created: 11/29/02
'Last updated: 09/September/12
'Last update: manually unload this image's main layer when the attached form is unloaded (to conserve memory)
'
'Every time the user loads an image, one of these forms is spawned.  This form also interfaces with several
' specialized program components in the MDIWindow module.  Look there for more information.
'
'***************************************************************************

Option Explicit

'These are used to track use of the Ctrl, Alt, and Shift keys
Dim ShiftDown As Boolean, CtrlDown As Boolean, AltDown As Boolean

'Track mouse button use on this form
Dim lMouseDown As Boolean, rMouseDown As Boolean

'Track initial mouse button locations
Dim initMouseX As Single, initMouseY As Single
    
'NOTE: _Activate and _GotFocus are confusing in VB6.  _Activate will be fired whenever a child form
' gains "focus."  _GotFocus will be pre-empted by controls on the form, so do not use it.

Private Sub Form_Activate()

    'Update the current form variable
    CurrentImage = Val(Me.Tag)
    
    'Display the size of this image in the status bar
    ' (NOTE: because this event will be fired when this form is first built, don't update the size values
    ' unless they actually exist.)
    If pdImages(CurrentImage).Width <> 0 Then DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height

    'If this MDI child is maximized, double-check that it's been drawn correctly.
    ' (This is necessary because VB doesn't handle _Resize() properly when switching between maximized MDI child forms)
    If Me.WindowState = 2 Then
        DoEvents
        PrepareViewport Me, "Maximized MDI child redraw"
    End If
    
    'Determine whether Undo, Redo, Fade-last are available
    tInit tUndo, pdImages(CurrentImage).UndoState
    tInit tRedo, pdImages(CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState
    
    'Determine whether save is enabled
    tInit tSave, Not pdImages(CurrentImage).HasBeenSaved
    
    'Restore the zoom value for this particular image (again, only if the form has been initialized)
    If pdImages(CurrentImage).Width <> 0 Then FormMain.CmbZoom.ListIndex = pdImages(CurrentImage).CurrentZoomValue
    
    'Finally, if the histogram window is open, redraw it
    If (FormHistogram.Visible = True) And pdImages(Me.Tag).loadedSuccessfully Then
        FormHistogram.TallyHistogramValues
        FormHistogram.DrawHistogram
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Add support for scrolling with the mouse wheel
    If IsProgramCompiled Then Call WheelHook(Me.HWnd)
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'Track which mouse buttons are pressed
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'These variables will hold the corresponding (x,y) coordinates on the image - NOT the viewport
    Dim imgX As Long, imgY As Long
    imgX = -1
    imgY = -1
    
    'Check mouse button use
    If Button = vbLeftButton Then
        
        'Check the location of the mouse to see if it's over the image
        If isMouseOverImage(x, y, Me) Then
        
            'Only track the mouse state if it is over the image
            lMouseDown = True
                
            Me.MousePointer = 2
        
            'Display the image coordinates under the mouse pointer
            displayImageCoordinates x, y, Me, imgX, imgY
        
            'Activate the selection and pass in the first two points
            pdImages(CurrentImage).selectionActive = True
            pdImages(CurrentImage).mainSelection.setInitialCoordinates imgX, imgY
    
            'Remember this location
            initMouseX = x
            initMouseY = y
    
        End If
        
    End If
    
    If Button = vbRightButton Then rMouseDown = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'These variables will hold the corresponding (x,y) coordinates on the image - NOT the viewport
    Dim imgX As Long, imgY As Long
    imgX = -1
    imgY = -1
    
    'Check the left mouse button
    If lMouseDown Then
    
        'First, check to see if a selection is active.  (In the future, we will be checking for other tools as well.)
        If pdImages(CurrentImage).selectionActive Then
        
            'Check the location of the mouse to see if it's over the image
            If isMouseOverImage(x, y, Me) Then
            
                'Display the image coordinates under the mouse pointer
                displayImageCoordinates x, y, Me, imgX, imgY
            
                'Pass new points to the active selection
                pdImages(CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
            
            'If the mouse coordinates are NOT over the image, we need to find the closest points in the image and pass those instead
            Else
        
                imgX = x
                imgY = y
                findNearestImageCoordinates imgX, imgY, Me
                
                'Pass those points to the active selection
                pdImages(CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
            
            End If
            
        End If
        
        'Force a redraw of the viewport
        ScrollViewport Me
    
    'This else means the LEFT mouse button is NOT down
    Else
    
        'Check the location of the mouse to see if it's over the image
        If isMouseOverImage(x, y, Me) Then
        
            Me.MousePointer = 2
        
            'Display the image coordinates under the mouse pointer
            displayImageCoordinates x, y, Me
        
        Else
        
            Me.MousePointer = 0
        
        End If
        
    End If
    
    makeFormPretty Me
    
End Sub

'Track which mouse buttons are released
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim imgX As Long, imgY As Long
    
    'Check mouse buttons
    If Button = vbLeftButton Then
    
        lMouseDown = False
        Me.MousePointer = 0
    
        'If a selection was being drawn, lock it into place
        If pdImages(CurrentImage).selectionActive Then
            
            'Lock the selection
            pdImages(CurrentImage).mainSelection.lockIn
            
            'Finally, check to see if this mouse location is the same as the initial mouse press.  If it is, clear the selection.
            If (x = initMouseX) And (y = initMouseY) Then
                pdImages(CurrentImage).mainSelection.lockRelease
                pdImages(CurrentImage).selectionActive = False
            End If
            
            'Force a redraw of the screen
            ScrollViewport Me
            
        End If
                
    End If
    
    If Button = vbRightButton Then rMouseDown = False
    
End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there)
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If FormMain.Enabled = False Then Exit Sub

    'Verify that the object being dragged is some sort of file or file list
    If Data.GetFormat(vbCFFiles) Then
        
        'Copy the filenames into an array
        Dim sFile() As String
        ReDim sFile(0 To Data.Files.Count) As String
        
        Dim oleFilename
        Dim tmpString As String
        
        Dim countFiles As Long
        countFiles = 0
        
        For Each oleFilename In Data.Files
            tmpString = CStr(oleFilename)
            If tmpString <> "" Then
                sFile(countFiles) = tmpString
                countFiles = countFiles + 1
            End If
        Next oleFilename
        
        'Because the OLE drop may include blank strings, verify the size of the array against countFiles
        ReDim Preserve sFile(0 To countFiles - 1) As String
        
        'Pass the list of filenames to PreLoadImage, which will load the images one-at-a-time
        PreLoadImage sFile
        
    End If
    
End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there)
Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If FormMain.Enabled = False Then Exit Sub

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Then
        'Inform the source (Explorer, in this case) that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files, don't allow a drop
        Effect = vbDropEffectNone
    End If

End Sub

'In VB6, _QueryUnload fires before _Unload.  We check for unsaved images here.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'If the user wants to be prompted about unsaved images, do it now
    If (ConfirmClosingUnsaved = True) And (pdImages(Me.Tag).IsActive = True) And (pdImages(Me.Tag).forInternalUseOnly = False) Then
    
        'Check the .HasBeenSaved property of the image associated with this form
        If pdImages(Me.Tag).HasBeenSaved = False Then
            
            'Generate our save message
            Dim saveMsg As String
            If FormMain.Enabled Then
                Me.SetFocus
                DoEvents
            End If
            saveMsg = "This image (""" & pdImages(Me.Tag).OriginalFileNameAndExtension & """) has not been saved.  Would you like to save it now?"
            
            'If this file exists on disk, warn them that this will initiate a SAVE, not a SAVE AS
            If pdImages(Me.Tag).LocationOnDisk <> "" Then saveMsg = saveMsg & vbCrLf & vbCrLf & "NOTE: if you click 'Yes', PhotoDemon will save this image using its current file name.  If you would like to save it with a different file name, please select 'Cancel', then click the File -> Save As menu."
            
            'Get the user's input
            Dim confirmReturn As VbMsgBoxResult
            confirmReturn = MsgBox(saveMsg, vbYesNoCancel + vbApplicationModal + vbQuestion, "Unsaved image data")
        
            'There are now three possible courses of action:
            ' 1) The user canceled.  Quit and abandon all notion of closing.
            ' 2) The user asked us to save this image.  Pass control to MenuSave (which will in turn call SaveAs if necessary)
            ' 3) The user doesn't give a shit.  Exit without saving.
            If confirmReturn = vbCancel Then
                Cancel = True
            ElseIf confirmReturn = vbYes Then
                
                'Attempt to save.  Note that the user can still cancel at this point, and we want to honor their cancellation
                Dim saveSuccessful As Boolean
                saveSuccessful = MenuSave(CLng(Me.Tag))
                
                'If something went wrong, or the user canceled the save dialog, stop the unload process
                Cancel = Not saveSuccessful
 
                'If we make it here and the save was successful, force an immediate unload
                If Cancel = False Then
                    Me.Visible = False
                    Unload Me
                End If
                
            ElseIf confirmReturn = vbNo Then
                Me.Visible = False
                Unload Me
           End If
        
        End If
    
    End If
    
End Sub

Private Sub Form_Resize()
    
    'Redraw this form if certain criteria are met (image loaded, form visible, viewport adjustments allowed)
    If (pdImages(Me.Tag).Width > 0) And (pdImages(Me.Tag).Height > 0) And (Me.Visible = True) Then
        DrawSpecificCanvas Me
        PrepareViewport Me, "Form_Resize(" & Me.ScaleWidth & "," & Me.ScaleHeight & ")"
    End If
    
    'The height of a newly created form is automatically set to 1.  This is normally changed when the image is
    ' resized to fit on screen, but if an image is loaded into a maximized window, the height value will remain
    ' at 1.  If the user ever un-maximized the window, it will leave a bare title bar behind, which looks
    ' terrible.  Thus, let's check for a height of 1, and if found resize the form to a larger (arbitrary) value.
    If Me.ScaleHeight <= 1 Then
        Me.Height = 6000
        Me.Width = 8000
    End If
    
    Dim i As Long
    
    'If the window is being un-maximized, it's necessary to redraw every image buffer (to check for scroll bar enabling/disabling)
    If pdImages(Me.Tag).WindowState = vbMaximized And Me.WindowState = 0 Then
        
        'Run a loop through every child form to see if all windows are being un-maximized
        ' (This will only happen when the user presses the "unmaximize" window button)
        Dim allShrunk As Boolean
        allShrunk = True
        
        Dim tForm As Form
        For Each tForm In VB.Forms
            If tForm.Name = "FormImage" Then
                If tForm.WindowState = vbMaximized Then allShrunk = False
            End If
        Next
        
        'If the user has unmaximized all windows, we need to redraw them
        If allShrunk = True Then
        
            'Loop through every image, redrawing as we go
            For i = 1 To CurrentImage
                If pdImages(i).IsActive = True Then
                    
                    'Remember this new window state and redraw the form containing this image
                    pdImages(i).WindowState = 0
                    PrepareViewport pdImages(i).containingForm, "Form_Resize(), user unmaximized MDI children"
                    
                    'While we're at it, make sure the images aren't still hidden off-form (which can happen if they were loaded while the window was maximized)
                    If pdImages(i).containingForm.Left >= FormMain.ScaleWidth Then pdImages(i).containingForm.Left = pdImages(i).WindowLeft
                    If pdImages(i).containingForm.Top >= FormMain.ScaleHeight Then pdImages(i).containingForm.Top = pdImages(i).WindowTop
    
                End If
            Next i
        End If
        
    End If
    
    'Remember this window state in the relevant pdImages object
    pdImages(Me.Tag).WindowState = Me.WindowState
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Release mouse wheel support
    If IsProgramCompiled Then Call WheelUnHook(Me.HWnd)
    
    Message "Closing image..."
    
    NumOfWindows = NumOfWindows - 1
    Me.Visible = False
    pdImages(Me.Tag).mainLayer.eraseLayer
    Set pdImages(Me.Tag).mainLayer = Nothing
    pdImages(Me.Tag).IsActive = False
    ClearUndo Me.Tag
    
    Message "Finished."

    'If this was the last (or only) open image and the histogram is loaded, unload the histogram
    ' (If we don't do this, the histogram may attempt to update, and without an active image it will throw an error)
    If NumOfWindows = 0 Then Unload FormHistogram
    
    UpdateMDIStatus
    
End Sub

Private Sub HScroll_Change()
    ScrollViewport Me
End Sub

Private Sub HScroll_Scroll()
    ScrollViewport Me
End Sub

Private Sub VScroll_Change()
    ScrollViewport Me
End Sub

Private Sub VScroll_Scroll()
    ScrollViewport Me
End Sub

'In VB6, a routine this like is required to support use of a mouse wheel.
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  
  On Error Resume Next
  
  'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
  If (VScroll.Visible = True) And (Not ShiftDown) And (Not CtrlDown) Then
  
    If Rotation < 0 Then
        
        If VScroll.Value + VScroll.LargeChange > VScroll.Max Then
            VScroll.Value = VScroll.Max
        Else
            VScroll.Value = VScroll.Value + VScroll.LargeChange
        End If
        
        ScrollViewport Me
    
    ElseIf Rotation > 0 Then
        
        If VScroll.Value - VScroll.LargeChange < VScroll.Min Then
            VScroll.Value = VScroll.Min
        Else
            VScroll.Value = VScroll.Value - VScroll.LargeChange
        End If
        
        ScrollViewport Me
        
    End If
  End If
  
  'Horizontal scrolling - only trigger if the horizontal scroll bar is visible AND a shift key has been pressed
  If (HScroll.Visible = True) And ShiftDown And (Not CtrlDown) Then
  
    If Rotation < 0 Then
        
        If HScroll.Value + HScroll.LargeChange > HScroll.Max Then
            HScroll.Value = HScroll.Max
        Else
            HScroll.Value = HScroll.Value + HScroll.LargeChange
        End If
        
        ScrollViewport Me
    
    ElseIf Rotation > 0 Then
        
        If HScroll.Value - HScroll.LargeChange < HScroll.Min Then
            HScroll.Value = HScroll.Min
        Else
            HScroll.Value = HScroll.Value - HScroll.LargeChange
        End If
        
        ScrollViewport Me
        
    End If
  End If
  
  'Zooming - only trigger when Ctrl has been pressed
  If CtrlDown Then
  
    If Rotation > 0 Then
        
        If FormMain.CmbZoom.ListIndex > 0 Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex - 1
    
    ElseIf Rotation < 0 Then
        
        If FormMain.CmbZoom.ListIndex < (FormMain.CmbZoom.ListCount - 1) Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex + 1
        
    End If
  End If
    
End Sub

