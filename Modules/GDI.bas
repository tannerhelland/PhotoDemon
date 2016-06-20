Attribute VB_Name = "GDI"
'***************************************************************************
'GDI interop manager
'Copyright 2001-2016 by Tanner Helland
'Created: 03/April/2001
'Last updated: 20/June/16
'Last update: split the GDI parts of the massive Drawing module into this dedicated module
'
'Like any Windows application, PD frequently interacts with GDI.  This module tries to manage the messiest bits
' of interop code.
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

Private Const GDI_OBJ_BITMAP As Long = 7&
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal srcDC As Long, ByVal srcObjectType As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal sizeOfBuffer As Long, ByVal ptrToBuffer As Long) As Long

Public Function GetBitmapHeaderFromDC(ByVal srcDC As Long) As GDI_Bitmap
    
    Dim hBitmap As Long
    hBitmap = GetCurrentObject(srcDC, GDI_OBJ_BITMAP)
    If (hBitmap <> 0) Then
        If (GetObject(hBitmap, Len(GetBitmapHeaderFromDC), VarPtr(GetBitmapHeaderFromDC)) = 0) Then
            InternalGDIError "GetObject failed on source hDC", , Err.LastDllError
        End If
    Else
        InternalGDIError "No bitmap in source hDC", "You can't query a DC for bitmap data if the DC doesn't have a bitmap selected into it!", Err.LastDllError
    End If
                        
End Function

'Add your own error-handling behavior here, as desired
Private Sub InternalGDIError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  The GDI interface encountered an error: """ & errName & """ - " & errDescription
        If (ErrNum <> 0) Then pdDebug.LogAction "(Also, an error number was reported: " & ErrNum & ")"
    #End If
End Sub

'Need a quick and dirty DC for something?  Call this.  (Just remember to free the DC when you're done!)
Public Function GetMemoryDC() As Long
    
    GetMemoryDC = CreateCompatibleDC(0&)
    
    'In debug mode, track how many DCs the program requests
    #If DEBUGMODE = 1 Then
        If GetMemoryDC <> 0 Then
            g_DCsCreated = g_DCsCreated + 1
        Else
            pdDebug.LogAction "WARNING!  GDI.GetMemoryDC() failed to create a new memory DC!"
        End If
    #End If
    
End Function

Public Sub FreeMemoryDC(ByVal srcDC As Long)
    
    If srcDC <> 0 Then
        
        Dim delConfirm As Long
        delConfirm = DeleteDC(srcDC)
    
        'In debug mode, track how many DCs the program frees
        #If DEBUGMODE = 1 Then
            If delConfirm <> 0 Then
                g_DCsDestroyed = g_DCsDestroyed + 1
            Else
                pdDebug.LogAction "WARNING!  GDI.FreeMemoryDC() failed to release DC #" & srcDC & "."
            End If
        #End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  GDI.FreeMemoryDC() was passed a null DC.  Fix this!"
        #End If
    End If
    
End Sub

Public Sub ForceGDIFlush()
    GdiFlush
End Sub

