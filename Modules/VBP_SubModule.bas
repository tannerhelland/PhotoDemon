Attribute VB_Name = "Misc_Uncategorized"
'***************************************************************************
'Miscellaneous Operations Handler
'Copyright ©2001-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 13/June/13
'Last update: removed many functions into new dedicated Math and Color modules
'
'If a function doesn't have a home in a more appropriate module, it gets stuck here. Over time, I'm
' hoping to clear out most of this module in favor of a more organized approach.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Types and API calls for processing ESC keypresses mid-loop
Private Type winMsg
    hWnd As Long
    sysMsg As Long
    wParam As Long
    lParam As Long
    msgTime As Long
    ptX As Long
    ptY As Long
End Type
'Private Declare Function TranslateMessage Lib "user32" (lpMsg As winMsg) As Long
'Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As winMsg) As Long
Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Const WM_KEYFIRST As Long = &H100
Private Const WM_KEYLAST As Long = &H108
Private Const PM_REMOVE As Long = &H1

Public cancelCurrentAction As Boolean

'Distance value for mouse_over events and selections; a literal "radius" below which the mouse cursor is considered "over" a point
Private Const mouseSelAccuracy As Double = 8

'This function will quickly and efficiently check the last unprocessed keypress submitted by the user.  If an ESC keypress was found,
' this function will return TRUE.  It is then up to the calling function to determine how to proceed.
Public Function userPressedESC(Optional ByVal displayConfirmationPrompt As Boolean = True) As Boolean

    Dim tmpMsg As winMsg
    
    'GetInputState returns a non-0 value if key or mouse events are pending.  By Microsoft's own admission, it is much faster
    ' than PeekMessage, so to keep things fast we check it before manually inspecting individual messages
    ' (see http://support.microsoft.com/kb/35605 for more details)
    If GetInputState() Then
    
        'Use the WM_KEYFIRST/LAST constants to explicitly request only keypress messages.  If the user has pressed multiple
        ' keys besides just ESC, this function may not operate as intended.  (Per the MSDN documentation: "...the first queued
        ' message that matches the specified filter is retrieved.")  We could technically parse all keypress messages and look
        ' for just ESC, but this would slow the function without providing any clear benefit.
        PeekMessage tmpMsg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE
        If tmpMsg.wParam = vbKeyEscape Then
            If displayConfirmationPrompt Then
                Dim msgReturn As VbMsgBoxResult
                msgReturn = pdMsgBox("Are you sure you want to cancel %1?", vbInformation + vbYesNo + vbApplicationModal, "Cancel image processing", LastProcess.Id)
                If msgReturn = vbYes Then cancelCurrentAction = True Else cancelCurrentAction = False
            Else
                cancelCurrentAction = True
            End If
        Else
            cancelCurrentAction = False
        End If
    Else
        cancelCurrentAction = False
    End If
    
    userPressedESC = cancelCurrentAction
    
End Function

'Populate a text box with a given integer value. This is done constantly across the program, so I use a sub to handle it, as
' there may be additional validations that need to be performed, and it's nice to be able to adjust those from a single location.
Public Sub copyToTextBoxI(ByRef dstTextBox As TextBox, ByVal srcValue As Long)

    'Remember the current cursor position
    Dim cursorPos As Long
    cursorPos = dstTextBox.SelStart

    'Overwrite the current text box value with the new value
    dstTextBox = CStr(srcValue)
    dstTextBox.Refresh
    
    'Restore the cursor to its original position
    If cursorPos >= Len(dstTextBox) Then cursorPos = Len(dstTextBox)
    dstTextBox.SelStart = cursorPos

End Sub

'Populate a text box with a given floating-point value. This is done constantly across the program, so I use a sub to handle it, as
' there may be additional validations that need to be performed, and it's nice to be able to adjust those from a single location.
Public Sub copyToTextBoxF(ByVal srcValue As Double, ByRef dstTextBox As TextBox, Optional ByVal numOfSD As Long = 2)

    'Remember the current cursor position
    Dim cursorPos As Long
    cursorPos = dstTextBox.SelStart

    'PhotoDemon never allows more than two significant digits for floating-point text boxes
    If numOfSD = 2 Then
        dstTextBox = Format(CStr(srcValue), "#0.00")
    Else
        dstTextBox = Format(CStr(srcValue), "#0.0")
    End If
    dstTextBox.Refresh
    
    'Restore the cursor to its original position
    If cursorPos >= Len(dstTextBox) Then cursorPos = Len(dstTextBox)
    dstTextBox.SelStart = cursorPos

End Sub

'Let a form know whether the mouse pointer is over its image or just the viewport
Public Function isMouseOverImage(ByVal x1 As Long, ByVal y1 As Long, ByRef srcForm As Form) As Boolean

    If (x1 >= pdImages(srcForm.Tag).targetLeft) And (x1 <= pdImages(srcForm.Tag).targetLeft + pdImages(srcForm.Tag).targetWidth) Then
        If (y1 >= pdImages(srcForm.Tag).targetTop) And (y1 <= pdImages(srcForm.Tag).targetTop + pdImages(srcForm.Tag).targetHeight) Then
            isMouseOverImage = True
            Exit Function
        Else
            isMouseOverImage = False
        End If
        isMouseOverImage = False
    End If

End Function

'Calculate and display the current mouse position.
' INPUTS: x and y coordinates of the mouse cursor, current form, and optionally two Double-type variables to receive the relative
' coordinates (e.g. location on the image) of the current mouse position.
Public Sub displayImageCoordinates(ByVal x1 As Double, ByVal y1 As Double, ByRef srcForm As Form, Optional ByRef copyX As Double, Optional ByRef copyY As Double)

    'Grab the current zoom value
    Dim ZoomVal As Double
    ZoomVal = g_Zoom.ZoomArray(pdImages(srcForm.Tag).CurrentZoomValue)
                
    'Calculate x and y positions, while taking into account zoom and scroll values
    x1 = srcForm.HScroll.Value + Int((x1 - pdImages(srcForm.Tag).targetLeft) / ZoomVal)
    y1 = srcForm.VScroll.Value + Int((y1 - pdImages(srcForm.Tag).targetTop) / ZoomVal)
    
    'When zoomed very far out, the values might be calculated incorrectly. Force them to the image dimensions if necessary.
    'If x1 < 0 Then x1 = 0
    'If y1 < 0 Then y1 = 0
    'If x1 > pdImages(srcForm.Tag).Width Then x1 = pdImages(srcForm.Tag).Width
    'If y1 > pdImages(srcForm.Tag).Height Then y1 = pdImages(srcForm.Tag).Height
        
    'If the user has requested copies of these coordinates, assign them now
    If copyX Then copyX = x1
    If copyY Then copyY = y1
    
    FormMain.lblCoordinates.Caption = "(" & x1 & "," & y1 & ")"
    FormMain.lblCoordinates.Refresh
    
End Sub

'If an x or y location is NOT in the image, find the nearest coordinate that IS in the image
Public Sub findNearestImageCoordinates(ByRef x1 As Double, ByRef y1 As Double, ByRef srcForm As Form)

    'Grab the current zoom value
    Dim ZoomVal As Double
    ZoomVal = g_Zoom.ZoomArray(pdImages(srcForm.Tag).CurrentZoomValue)

    'Calculate x and y positions, while taking into account zoom and scroll values
    x1 = srcForm.HScroll.Value + Int((x1 - pdImages(srcForm.Tag).targetLeft) / ZoomVal)
    y1 = srcForm.VScroll.Value + Int((y1 - pdImages(srcForm.Tag).targetTop) / ZoomVal)
    
    'Force any invalid values to their nearest matching point in the image
    If x1 < 0 Then x1 = 0
    If y1 < 0 Then y1 = 0
    If x1 >= pdImages(srcForm.Tag).Width Then x1 = pdImages(srcForm.Tag).Width - 1
    If y1 >= pdImages(srcForm.Tag).Height Then y1 = pdImages(srcForm.Tag).Height - 1

End Sub

'This sub will return a constant correlating to the nearest selection point. Its return values are:
' 0 - Cursor is not near a selection point
' 1 - NW corner
' 2 - NE corner
' 3 - SE corner
' 4 - SW corner
' 5 - N edge
' 6 - E edge
' 7 - S edge
' 8 - W edge
' 9 - interior of selection, not near a corner or edge
Public Function findNearestSelectionCoordinates(ByRef x1 As Single, ByRef y1 As Single, ByRef srcForm As Form) As Long

    'Grab the current zoom value
    Dim ZoomVal As Double
    ZoomVal = g_Zoom.ZoomArray(pdImages(srcForm.Tag).CurrentZoomValue)

    'Calculate x and y positions, while taking into account zoom and scroll values
    x1 = srcForm.HScroll.Value + Int((x1 - pdImages(srcForm.Tag).targetLeft) / ZoomVal)
    y1 = srcForm.VScroll.Value + Int((y1 - pdImages(srcForm.Tag).targetTop) / ZoomVal)
    
    'With x1 and y1 now representative of a location within the image, it's time to start calculating distances.
    Dim tLeft As Double, tTop As Double, tRight As Double, tBottom As Double
    
    If (pdImages(srcForm.Tag).mainSelection.getSelectionShape = sRectangle) Or (pdImages(srcForm.Tag).mainSelection.getSelectionShape = sCircle) Then
        tLeft = pdImages(srcForm.Tag).mainSelection.selLeft
        tTop = pdImages(srcForm.Tag).mainSelection.selTop
        tRight = pdImages(srcForm.Tag).mainSelection.selLeft + pdImages(srcForm.Tag).mainSelection.selWidth
        tBottom = pdImages(srcForm.Tag).mainSelection.selTop + pdImages(srcForm.Tag).mainSelection.selHeight
    Else
        tLeft = pdImages(srcForm.Tag).mainSelection.boundLeft
        tTop = pdImages(srcForm.Tag).mainSelection.boundTop
        tRight = pdImages(srcForm.Tag).mainSelection.boundLeft + pdImages(srcForm.Tag).mainSelection.boundWidth
        tBottom = pdImages(srcForm.Tag).mainSelection.boundTop + pdImages(srcForm.Tag).mainSelection.boundHeight
    End If
    
    'Adjust the mouseAccuracy value based on the current zoom value
    Dim mouseAccuracy As Double
    mouseAccuracy = mouseSelAccuracy * (1 / ZoomVal)
    
    'Before doing anything else, make sure the pointer is actually worth checking - e.g. make sure it's near the selection
    'If (x1 < tLeft - mouseAccuracy) Or (x1 > tRight + mouseAccuracy) Or (y1 < tTop - mouseAccuracy) Or (y1 > tBottom + mouseAccuracy) Then
    '    findNearestSelectionCoordinates = 0
    '    Exit Function
    'End If
    
    'Find the smallest distance for this mouse position
    Dim minDistance As Double
    minDistance = mouseAccuracy
    
    Dim closestPoint As Long
    
    'If we made it here, this mouse location is worth evaluating.  How we evaluate it depends on the shape of the current selection.
    Select Case g_CurrentTool
    
        Case SELECT_RECT, SELECT_CIRC
    
            'Corners get preference, so check them first.
            Dim nwDist As Double, neDist As Double, seDist As Double, swDist As Double
            
            nwDist = distanceTwoPoints(x1, y1, tLeft, tTop)
            neDist = distanceTwoPoints(x1, y1, tRight, tTop)
            swDist = distanceTwoPoints(x1, y1, tLeft, tBottom)
            seDist = distanceTwoPoints(x1, y1, tRight, tBottom)
            
            'Find the smallest distance for this mouse position
            closestPoint = -1
            
            If nwDist <= minDistance Then
                minDistance = nwDist
                closestPoint = 1
            End If
            
            If neDist <= minDistance Then
                minDistance = neDist
                closestPoint = 2
            End If
            
            If seDist <= minDistance Then
                minDistance = seDist
                closestPoint = 3
            End If
            
            If swDist <= minDistance Then
                minDistance = swDist
                closestPoint = 4
            End If
            
            'Was a close point found? If yes, then return that value
            If closestPoint <> -1 Then
                findNearestSelectionCoordinates = closestPoint
                Exit Function
            End If
        
            'If we're at this line of code, a closest corner was not found. So check edges next.
            Dim nDist As Double, eDist As Double, sDist As Double, wDist As Double
            
            nDist = distanceOneDimension(y1, tTop)
            eDist = distanceOneDimension(x1, tRight)
            sDist = distanceOneDimension(y1, tBottom)
            wDist = distanceOneDimension(x1, tLeft)
            
            If (nDist <= minDistance) Then
                minDistance = nDist
                closestPoint = 5
            End If
            
            If (eDist <= minDistance) Then
                minDistance = eDist
                closestPoint = 6
            End If
            
            If (sDist <= minDistance) Then
                minDistance = sDist
                closestPoint = 7
            End If
            
            If (wDist <= minDistance) Then
                minDistance = wDist
                closestPoint = 8
            End If
            
            'Was a close point found? If yes, then return that value.
            If closestPoint <> -1 Then
                findNearestSelectionCoordinates = closestPoint
                Exit Function
            End If
        
            'If we're at this line of code, a closest edge was not found. Perform one final check to ensure that the mouse is within the
            ' image's boundaries, and if it is, return the "move selection" ID, then exit.
            If (x1 > tLeft) And (x1 < tRight) And (y1 > tTop) And (y1 < tBottom) Then
                findNearestSelectionCoordinates = 9
            Else
                findNearestSelectionCoordinates = 0
            End If
            
        Case SELECT_LINE
    
            'Line selections are simple - we only care if the mouse is by (x1,y1) or (x2,y2)
            Dim xCoord As Double, yCoord As Double
            Dim firstDist As Double, secondDist As Double
            
            closestPoint = 0
            
            pdImages(srcForm.Tag).mainSelection.getSelectionCoordinates 1, xCoord, yCoord
            firstDist = distanceTwoPoints(x1, y1, xCoord, yCoord)
            
            pdImages(srcForm.Tag).mainSelection.getSelectionCoordinates 2, xCoord, yCoord
            secondDist = distanceTwoPoints(x1, y1, xCoord, yCoord)
                        
            If firstDist <= minDistance Then closestPoint = 1
            If secondDist <= minDistance Then closestPoint = 2
            
            'Was a close point found? If yes, then return that value.
            findNearestSelectionCoordinates = closestPoint
            Exit Function
            
    End Select

End Function
