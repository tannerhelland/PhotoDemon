Attribute VB_Name = "Outside_SetIcon"
'Note: this file has been modified for use within PhotoDemon.

'This code was originally written by Steve McMahon.  You may download the original from this link:
' http://www.vbaccelerator.com/home/VB/Tips/Setting_the_App_Icon_Correctly/article.asp

'To the best of my knowledge, this code is released under a CC-BY-1.0 license.  (Assumed from the footer text of vbaccelerator.com: "All contents of this web site are licensed under a Creative Commons Licence, except where otherwise noted.")
' You may access a complete copy of this license at the following link:
' http://creativecommons.org/licenses/by/1.0/

'Many thanks to Steve and vbaccelerator.com for this excellent icon-related code

Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
   
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
   
'Private Const LR_DEFAULTCOLOR = &H0
'Private Const LR_MONOCHROME = &H1
'Private Const LR_COLOR = &H2
'Private Const LR_COPYRETURNORG = &H4
'Private Const LR_COPYDELETEORG = &H8
'Private Const LR_LOADFROMFILE = &H10
'Private Const LR_LOADTRANSPARENT = &H20
'Private Const LR_DEFAULTSIZE = &H40
'Private Const LR_VGACOLOR = &H80
'Private Const LR_LOADMAP3DCOLORS = &H1000
'Private Const LR_CREATEDIBSECTION = &H2000
'Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

'This variable will hold the hWnd of the hidden top-most parent of the program (created by VB)
Dim lHwndTop As Long

'Rather than constantly re-load the original icons from file, store them once generated
Public origIcon32 As Long, origIcon16 As Long

Public Sub SetIcon(ByVal hWnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)

    Dim lHwnd As Long
    Dim cx As Long
    Dim cy As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long
      
    If (bSetAsAppIcon) Then
        ' Find VB's hidden parent window:
        lHwnd = hWnd
        lHwndTop = lHwnd
        Do While Not (lHwnd = 0)
            lHwnd = GetWindow(lHwnd, GW_OWNER)
            If Not (lHwnd = 0) Then
                lHwndTop = lHwnd
            End If
        Loop
    End If
       
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    origIcon32 = hIconLarge
    
    If bSetAsAppIcon Then SendMessageLong lHwndTop, WM_SETICON, ICON_BIG, hIconLarge
    SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
       
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    origIcon16 = hIconSmall
    
    If bSetAsAppIcon Then SendMessageLong lHwndTop, WM_SETICON, ICON_SMALL, hIconSmall
    SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall
   
End Sub

'During run-time, the user has an option to use the current MDI child window's icon as the task bar icon as well.
Public Sub setNewTaskbarIcon(ByVal iconhWnd32 As Long)
    If Not ALLOW_DYNAMIC_ICONS Then Exit Sub
    SendMessageLong FormMain.hWnd, WM_SETICON, ICON_BIG, iconhWnd32
End Sub

Public Sub setNewAppIcon(ByVal iconhWnd16 As Long)
    If Not ALLOW_DYNAMIC_ICONS Then Exit Sub
    SendMessageLong FormMain.hWnd, WM_SETICON, ICON_SMALL, iconhWnd16
End Sub

