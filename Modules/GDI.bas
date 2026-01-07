Attribute VB_Name = "GDI"
'***************************************************************************
'GDI interop manager
'Copyright 2001-2026 by Tanner Helland
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
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
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

Private Type GDI_BitmapInfoHeader
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
 
Private Type GDI_RGBQuad
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
 
Private Type GDI_BitmapInfo
    bmiHeader As GDI_BitmapInfoHeader
    bmiColors(0 To 255) As GDI_RGBQuad
End Type

'Private Enum GDI_PenStyle
'    PS_SOLID = 0
'    PS_DASH = 1
'    PS_DOT = 2
'    PS_DASHDOT = 3
'    PS_DASHDOTDOT = 4
'End Enum
'
'#If False Then
'    Private Const PS_SOLID = 0, PS_DASH = 1, PS_DOT = 2, PS_DASHDOT = 3, PS_DASHDOTDOT = 4
'#End If

Private Const GDI_OBJ_BITMAP As Long = 7&
Private Const GDI_CBM_INIT As Long = &H4
Private Const GDI_DIB_RGB_COLORS As Long = &H0

Private Declare Function BitBlt Lib "gdi32" (ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal rastOp As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rastOp As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hDC As Long, ByRef lpInfoHeader As GDI_BitmapInfoHeader, ByVal dwUsage As Long, ByVal ptrToInitBits As Long, ByVal ptrToInitBitmapInfo As Long, ByVal wUsage As Long) As Long
'Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As GDI_PenStyle, ByVal nWidth As Long, ByVal srcColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal srcColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function GdiFlush Lib "gdi32" () As Long    'We don't require this right now, but it may be useful in the future
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal srcDC As Long, ByVal srcObjectType As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal sizeOfBuffer As Long, ByVal ptrToBuffer As Long) As Long
'Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal pointerToRectOfOldCoords As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Helper functions from user32
Private Declare Function FillRect Lib "user32" (ByVal hDstDC As Long, ByVal ptrToRect As Long, ByVal hSrcBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByVal ptrToRectL As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Public Function BitBltWrapper(ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, Optional ByVal rastOp As Long = vbSrcCopy) As Boolean
    BitBltWrapper = (BitBlt(hDstDC, dstX, dstY, dstWidth, dstHeight, hSrcDC, srcX, srcY, rastOp) <> 0)
End Function

Public Function StretchBltWrapper(ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal rastOp As Long = vbSrcCopy) As Boolean
    StretchBltWrapper = (StretchBlt(hDstDC, dstX, dstY, dstWidth, dstHeight, hSrcDC, srcX, srcY, srcWidth, srcHeight, rastOp) <> 0)
End Function

Public Function GetClientRectWrapper(ByVal srcHWnd As Long, ByVal ptrToDestRect As Long) As Boolean
    GetClientRectWrapper = (GetClientRect(srcHWnd, ptrToDestRect) <> 0)
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
Public Function GetMemoryDC(Optional ByVal compatDC As Long = 0&) As Long
    GetMemoryDC = CreateCompatibleDC(compatDC)
    If (GetMemoryDC <> 0) Then
        PDDebug.UpdateResourceTracker PDRT_hDC, 1
    Else
        PDDebug.LogAction "WARNING!  GDI.GetMemoryDC() failed to create a compatible DC.  DLL Error: #" & Err.LastDllError
    End If
End Function

Public Sub FreeMemoryDC(ByRef srcDC As Long)
    If (srcDC <> 0) Then
        If (DeleteDC(srcDC) <> 0) Then
            PDDebug.UpdateResourceTracker PDRT_hDC, -1
            srcDC = 0
        Else
            PDDebug.LogAction "WARNING!  GDI.FreeMemoryDC() failed to release the requested DC.  DLL Error: #" & Err.LastDllError
        End If
    End If
End Sub

'PD doesn't require this right now, but it may be useful in the future
'Public Sub ForceGDIFlush()
'    GdiFlush
'End Sub

'Basic wrappers for rect-filling and rect-tracing via GDI
Public Sub FillRectToDC(ByVal targetDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal rectWidth As Long, ByVal rectHeight As Long, ByVal crColor As Long)
    
    'Failsafe checks
    If (targetDC <> 0) Then
    
        'Create a brush with the specified color
        Dim tmpBrush As Long
        tmpBrush = CreateSolidBrush(crColor)
        
        If (tmpBrush <> 0) Then
        
            'Fill the rect
            Dim tmpRect As RectL
            With tmpRect
                .Left = x1
                .Top = y1
                .Right = x1 + rectWidth + 1
                .Bottom = y1 + rectHeight + 1
            End With
            
            FillRect targetDC, VarPtr(tmpRect), tmpBrush
            If (DeleteObject(tmpBrush) = 0) Then PDDebug.LogAction "WARNING!  GDI.FillRectToDC failed to free the brush it allocated."
            
        Else
            PDDebug.LogAction "WARNING!  GDI.FillRectToDC failed to create a solid brush"
        End If
        
    End If

End Sub

'Given a DIB, return a DDB.  IMPORTANT!  If the DIB is 32-bpp, you should (generally) unpremultiply alpha first.
' Most DDB-related functions do not handle premultiplied alpha correctly.
Public Function GetDDBFromDIB(ByRef srcDIB As pdDIB) As Long
    
    Dim tmpDC As Long
    tmpDC = GetDC(0)
    
    Dim tmpBIHeader As GDI_BitmapInfoHeader
    With tmpBIHeader
        .biSize = LenB(tmpBIHeader)
        .biWidth = srcDIB.GetDIBWidth
        .biHeight = -1 * srcDIB.GetDIBHeight
        .biBitCount = srcDIB.GetDIBColorDepth
        .biPlanes = 1
    End With
    
    Dim tmpBitmapInfo As GDI_BitmapInfo
    CopyMemoryStrict VarPtr(tmpBitmapInfo.bmiHeader), srcDIB.GetDIBHeader, LenB(tmpBitmapInfo.bmiHeader)
    GetDDBFromDIB = CreateDIBitmap(tmpDC, tmpBIHeader, GDI_CBM_INIT, srcDIB.GetDIBPointer, VarPtr(tmpBitmapInfo), GDI_DIB_RGB_COLORS)
    
    ReleaseDC 0, tmpDC
    
End Function

'Add your own error-handling behavior here, as desired
Private Sub InternalGDIError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal errNum As Long = 0)
    PDDebug.LogAction "WARNING!  The GDI interface encountered an error: """ & errName & """ - " & errDescription
    If (errNum <> 0) Then PDDebug.LogAction "(Also, an error number was reported: " & errNum & ")"
End Sub
