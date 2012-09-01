Attribute VB_Name = "GDI_Plus"
'***************************************************************************
'GDI+ Interface
'Copyright ©2011-2012 by Tanner Helland
'Created: 1/September/12
'Last updated: 1/September/12
'Last update: initial build
'
'This interface provides a means for interacting with the unnecessarily complex (and overwrought) GDI+ module.  GDI+ is
' primarily used as a fallback for image loading if the FreeImage DLL cannot be found.
'
'These routines are adapted from the work of a number of other talented VB programmers.  Since GDI+ is not well-documented
' for VB users, I have pieced this module together from the following pieces of code:
' Avery P's initial GDI+ deconstruction: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
' Carles P.V.'s iBMP modifications: http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Robert Rayment's PaintRR modifications: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1
' Many thanks to both these individuals for their outstanding work on graphics in VB.
'
'***************************************************************************

Option Explicit

'GDI+ Enums
Public Enum GDIPlusStatus
    [OK] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow ] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized ] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

'GDI+ required types
Private Type GDIPlusStartupInput
    GDIPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Declare Function GDIPlusStartup Lib "gdiplus" (ByRef Token As Long, ByRef InputBuf As GDIPlusStartupInput, Optional ByVal OutputBuffer As Long = 0&) As GDIPlusStatus
Private Declare Function GDIPlusShutdown Lib "gdiplus" (ByVal Token As Long) As GDIPlusStatus

'When GDI+ is initialized, it will assign us a token.  We use this to release GDI+ when the program terminates.
Private GDIPlusToken As Long


'At start-up, this function is called to determine whether or not we have GDI+ available on this machine.
Public Function isGDIPlusAvailable() As Boolean

    Dim GDICheck As GDIPlusStartupInput
    GDICheck.GDIPlusVersion = 1
    
    If (GDIPlusStartup(GDIPlusToken, GDICheck) <> [OK]) Then
        isGDIPlusAvailable = False
    Else
        isGDIPlusAvailable = True
    End If

End Function

Public Function releaseGDIPlus()
    GDIPlusShutdown GDIPlusToken
End Function
