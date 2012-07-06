Attribute VB_Name = "Outside_modMouseWheel"
'Note: this file has been modified for use within PhotoDemon.

'This code was adopted from http://www.vbforums.com/showthread.php?s=ff827c56c69cb7ad5dcbab38a92b5799&t=388222, accessed on 20 April 2012
' Thank you to user "bushmobile" for supplying the initial version of this code.

Option Explicit



' Store WndProcs
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
                ByVal HWnd As Long, _
                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" ( _
                ByVal HWnd As Long, _
                ByVal lpString As String, _
                ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" ( _
                ByVal HWnd As Long, _
                ByVal lpString As String) As Long

' Hooking
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc As Long, _
                ByVal HWnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal HWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                ByVal HWnd As Long, _
                ByVal Msg As Long, _
                wParam As Any, _
                lParam As Any) As Long

' Position Checking
Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal HWnd As Long, _
                lpRect As RECT) As Long
                
Private Declare Function GetParent Lib "user32" ( _
                ByVal HWnd As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157

'Tanner's addition: extra constants for checking mouse forward/back keys
Private Const WM_XBUTTONDOWN As Long = &H20B
Private Const WM_XBUTTONUP As Long = &H20C
Private Const WM_XBUTTONDBLCLK As Long = &H20D
Private Const WM_NCXBUTTONUP As Long = &H20F
Private Const WM_NCXBUTTONDBCLK As Long = &H210
Private Const WM_NCXBUTTONDOWN  As Long = &HAB

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

' Check Messages
' ================================================
Private Function WindowProc(ByVal lWnd As Long, ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim MouseKeys As Long
  Dim Rotation As Long
  Dim Xpos As Long
  Dim Ypos As Long
  Dim fFrm As Form

    On Error Resume Next

  Select Case lMsg
  
    Case WM_MOUSEWHEEL
    
      MouseKeys = wParam And 65535
      Rotation = wParam / 65536
      Xpos = lParam And 65535
      Ypos = lParam / 65536
      
      Set fFrm = GetForm(lWnd)
      If fFrm Is Nothing Then
        ' it's not a form
        If Not IsOver(lWnd, Xpos, Ypos) And IsOver(GetParent(lWnd), Xpos, Ypos) Then
          ' it's not over the control and is over the form,
          ' so fire mousewheel on form (if it's not a dropped down combo)
          If SendMessage(lWnd, CB_GETDROPPEDSTATE, 0&, 0&) <> 1 Then
            GetForm(GetParent(lWnd)).MouseWheel MouseKeys, Rotation, Xpos, Ypos
            Exit Function ' Discard scroll message to control
          End If
        End If
      Else
        ' it's a form so fire mousewheel
        If IsOver(fFrm.HWnd, Xpos, Ypos) Then fFrm.MouseWheel MouseKeys, Rotation, Xpos, Ypos
      End If
      
    'Forgive my use of arbitrary numbers here, but I used brute force testing to discover what messages my
    ' mouse's back/forward keys were sending.  These results were found by tracking the lMsg/wParam/lParam values
    ' manually.  I could redefine them as constants, but have not gone to the trouble of it... yet.
    Case 793
            
        'Mouse back key; I have no idea if this value is consistent between different hardware vendors
        If lParam = -2147418112 Then
            If pdImages(CurrentImage).UndoState = True Then Process Undo
        'Mouse forward key
        ElseIf lParam = -2147352576 Then
            If pdImages(CurrentImage).RedoState = True Then Process Redo
        End If
            
        'This line of code can be used to display the parameters in PhotoDemon's status bar - it's useful for tracking
        ' arbitrary key presses or functions on devices VB doesn't inherently support.
        'If lMsg > 700 Then Message CStr(lMsg) & "," & CStr(wParam) & "," & CStr(lParam)
            
  End Select
  
  WindowProc = CallWindowProc(GetProp(lWnd, "PrevWndProc"), lWnd, lMsg, wParam, lParam)
End Function

' Hook / UnHook
' ================================================
Public Sub WheelHook(ByVal HWnd As Long)
  On Error Resume Next
  SetProp HWnd, "PrevWndProc", SetWindowLong(HWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub WheelUnHook(ByVal HWnd As Long)
  On Error Resume Next
  SetWindowLong HWnd, GWL_WNDPROC, GetProp(HWnd, "PrevWndProc")
  RemoveProp HWnd, "PrevWndProc"
End Sub

' Window Checks
' ================================================
Public Function IsOver(ByVal HWnd As Long, ByVal Lx As Long, ByVal lY As Long) As Boolean
  Dim rectCtl As RECT
  GetWindowRect HWnd, rectCtl
  With rectCtl
    IsOver = (Lx >= .Left And Lx <= .Right And lY >= .Top And lY <= .Bottom)
  End With
End Function

Private Function GetForm(ByVal HWnd As Long) As Form
  For Each GetForm In Forms
    If GetForm.HWnd = HWnd Then Exit Function
  Next GetForm
  Set GetForm = Nothing
End Function

Public Sub PictureBoxZoom(ByRef PicBox As PictureBox, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  PicBox.Cls
  PicBox.Print "MouseWheel " & IIf(Rotation < 0, "Down", "Up")
End Sub
