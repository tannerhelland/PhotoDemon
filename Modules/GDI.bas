Attribute VB_Name = "GDI"
'***************************************************************************
'GDI interop manager
'Copyright 2001-2016 by Tanner Helland
'Created: 03/April/2001
'Last updated: 28/June/16
'Last update: continued clean-up of PD-specific code
'
'To improve performance, pd2D falls back to GDI in cases where GDI behavior is functionally identical.  This module
' manages all GDI-specific code paths.
'
'(I should probably mention that some non-GDI bits also exist in this module, like retrieving data from hWnds which
' actually happens in user32, but it doesn't really make sense to split those into their own module.)
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'For clarity, GDI's "BITMAP" type is referred to as "GDI_BITMAP" throughout PD.
Public Type GDI_Bitmap
    Type As Long
    Width As Long
    Height As Long
    WidthBytes As Long
    Planes As Integer
    BitsPerPixel As Integer
    Bits As Long
End Type

Private Enum GDI_PenStyle
    PS_SOLID = 0
    PS_DASH = 1
    PS_DOT = 2
    PS_DASHDOT = 3
    PS_DASHDOTDOT = 4
End Enum

#If False Then
    Private Const PS_SOLID = 0, PS_DASH = 1, PS_DOT = 2, PS_DASHDOT = 3, PS_DASHDOTDOT = 4
#End If

Private Const GDI_OBJ_BITMAP As Long = 7&

Private Declare Function BitBlt Lib "gdi32" (ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal rastOp As Long) As Boolean
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rastOp As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As GDI_PenStyle, ByVal nWidth As Long, ByVal srcColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal srcColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal srcDC As Long, ByVal srcObjectType As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal sizeOfBuffer As Long, ByVal ptrToBuffer As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal pointerToRectOfOldCoords As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Helper functions from user32
Private Declare Function FillRect Lib "user32" (ByVal hDstDC As Long, ByVal ptrToRect As Long, ByVal hSrcBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByVal ptrToRectL As Long) As Long

Public Function BitBltWrapper(ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, Optional ByVal rastOp As Long = vbSrcCopy) As Boolean
    BitBltWrapper = CBool(BitBlt(hDstDC, dstX, dstY, dstWidth, dstHeight, hSrcDC, srcX, srcY, rastOp) <> 0)
End Function

Public Function StretchBltWrapper(ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal rastOp As Long = vbSrcCopy) As Boolean
    StretchBltWrapper = CBool(StretchBlt(hDstDC, dstX, dstY, dstWidth, dstHeight, hSrcDC, srcX, srcY, srcWidth, srcHeight, rastOp) <> 0)
End Function

Public Function GetClientRectWrapper(ByVal srcHWnd As Long, ByVal ptrToDestRect As Long) As Boolean
    GetClientRectWrapper = CBool(GetClientRect(srcHWnd, ptrToDestRect) <> 0)
End Function

Public Function GetBitmapHeaderFromDC(ByVal srcDC As Long) As GDI_Bitmap
    
    Dim hBitmap As Long
    hBitmap = GetCurrentObject(srcDC, GDI_OBJ_BITMAP)
    If (hBitmap <> 0) Then
        If (GetObject(hBitmap, LenB(GetBitmapHeaderFromDC), VarPtr(GetBitmapHeaderFromDC)) = 0) Then
            InternalGDIError "GetObject failed on source hDC", , Err.LastDllError
        End If
    Else
        InternalGDIError "No bitmap in source hDC", "You can't query a DC for bitmap data if the DC doesn't have a bitmap selected into it!", Err.LastDllError
    End If
                        
End Function

'Need a quick and dirty DC for something?  Call this.  (Just remember to free the DC when you're done!)
Public Function GetMemoryDC() As Long
    GetMemoryDC = CreateCompatibleDC(0&)
End Function

Public Sub FreeMemoryDC(ByVal srcDC As Long)
    If (srcDC <> 0) Then DeleteDC srcDC
End Sub

Public Sub ForceGDIFlush()
    GdiFlush
End Sub

'Basic wrapper to line-drawing via GDI
Public Sub DrawLineToDC(ByVal targetDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal crColor As Long)
    
    'Create a pen with the specified color
    Dim newPen As Long
    newPen = CreatePen(PS_SOLID, 1, crColor)
    
    'Select the pen into the target DC
    Dim oldObject As Long
    oldObject = SelectObject(targetDC, newPen)
    
    'Render the line
    MoveToEx targetDC, x1, y1, 0&
    LineTo targetDC, x2, y2
    
    'Remove the pen and delete it
    SelectObject targetDC, oldObject
    DeleteObject newPen

End Sub

'Basic wrappers for rect-filling and rect-tracing via GDI
Public Sub FillRectToDC(ByVal targetDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal rectWidth As Long, ByVal rectHeight As Long, ByVal crColor As Long)

    'Create a brush with the specified color
    Dim tmpBrush As Long
    tmpBrush = CreateSolidBrush(crColor)
    
    'Select the brush into the target DC
    Dim oldObject As Long
    oldObject = SelectObject(targetDC, tmpBrush)
    
    'Fill the rect
    Dim tmpRect As RECTL
    With tmpRect
        .Left = x1
        .Top = y1
        .Right = x1 + rectWidth + 1
        .Bottom = y1 + rectHeight + 1
    End With
    
    FillRect targetDC, VarPtr(tmpRect), tmpBrush
    
    'Remove the brush and delete it
    SelectObject targetDC, oldObject
    DeleteObject tmpBrush

End Sub

'Add your own error-handling behavior here, as desired
Private Sub InternalGDIError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  The GDI interface encountered an error: """ & errName & """ - " & errDescription
        If (ErrNum <> 0) Then pdDebug.LogAction "(Also, an error number was reported: " & ErrNum & ")"
    #End If
End Sub

