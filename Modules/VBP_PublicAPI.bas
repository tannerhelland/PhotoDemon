Attribute VB_Name = "Public_API"
'Any and all *publicly* necessary API declarations can be found here
' (Note 1: privately declared API calls have been left in their respective forms/modules)
' (Note 2: it makes more sense to keep API-related constants here than in the constants
'  module, so don't be surprised to find constants here)
' (Note 3: I haven't searched every form/module/class for duplicate API calls... yet)

Option Explicit

Public Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data)
Public Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)

Public Type SAFEARRAYBOUND
    cElements As Long
    lBound   As Long
End Type

Public Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Public Type SAFEARRAY1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements As Long
    lBound   As Long
End Type

'These functions are used to scroll through consecutive MDI windows without flickering
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Any, lParam As Any) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

'Drawing calls
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal rastOp As Long) As Boolean
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rastOp As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDestDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
Public Const STRETCHBLT_COLORONCOLOR As Long = 3
Public Const STRETCHBLT_HALFTONE As Long = 4

'API calls for explicitly calling dlls.  This allows us to build DLL paths at runtime, and it also allows
' us to call any DLL we like without first passing them through regsvr32.
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'ShellExecute is preferable to VB's 'Shell' command; I use it for launching URLs using the system's default web browser
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

'Various API calls for manually downloading files from the Internet
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, ByRef lpszBuffer As String, ByRef lpdwBufferLength As Long) As Boolean
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Boolean

'Constants used by the Windows Internet APIs
Public Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const HTTP_QUERY_CONTENT_LENGTH As Long = 5
'I've experimented with this constant with no luck; VB simply doesn't like asynchronous connections :(
'Public Const INTERNET_FLAG_ASYNC = &H10000000

'Some PhotoDemon functions are capable of timing themselves.  GetTickCount is used to do this.
Public Declare Function GetTickCount Lib "kernel32" () As Long

'RGB <-> HSL conversion
Public Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Public Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Public Declare Function ColorAdjustLuma Lib "shlwapi" (ByVal clrRGB As Long, ByVal n As Long, ByVal fScale As Long) As Long

'Request mouse tracking for a given object
Public Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hWndTrack As Long
    dwHoverTime As Long
End Type
Public Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long
