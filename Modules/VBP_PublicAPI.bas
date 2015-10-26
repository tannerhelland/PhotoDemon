Attribute VB_Name = "Public_API"
'Any and all *publicly* necessary API declarations can be found here
' (Note 1: privately declared API calls have obviously been left in their respective forms/modules)
' (Note 2: it makes more sense to keep API-related constants here than in the constants module)

Option Explicit

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Type POINTFLOAT
   x As Single
   y As Single
End Type

Public Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'Most API calls handle window position and movement in terms of a rect-type variable
Public Type winRect
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data)
Public Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)

Public Const FADF_AUTO As Long = (&H1)
Public Const FADF_FIXEDSIZE As Long = (&H10)

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

'These functions are used to interact with various windows
Public Const MONITOR_DEFAULTTONEAREST As Long = &H2
Public Const CB_SETMINVISIBLE As Long = 339
Public Const CB_SHOWDROPDOWN As Long = &H14F
Public Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function MonitorFromWindow Lib "user32" (ByVal myHwnd As Long, ByVal dwFlags As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal bEnable As Long) As Long

'NOTE!  By 6.6's release, I hope to remove the need for this A equivalent.  Search to see which functions make use of it.
Public Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Drawing calls
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal rastOp As Long) As Boolean
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rastOp As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDestDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal nScan As Long, ByVal NumScans As Long, ByRef lpBits As Any, ByRef BitsInfo As Any, ByVal wUsage As Long) As Long
Public Declare Function SetDIBitsToDC Lib "gdi32" Alias "SetDIBitsToDevice" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal nScan As Long, ByVal NumScans As Long, ByVal lpBits As Long, ByVal lpBitsInfo As Long, ByVal wUsage As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByVal lpBits As Long, ByVal lpBitsInfo As Long, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Const STRETCHBLT_COLORONCOLOR As Long = 3
Public Const STRETCHBLT_HALFTONE As Long = 4

'API calls for explicitly calling dlls.  This allows us to build DLL paths at runtime, and it also allows
' us to call any DLL we like without first passing them through regsvr32.
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'ShellExecute is preferable to VB's 'Shell' command; I use it for launching URLs using the system's default web browser.
' (Note the use of pointers
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal ptrToOperationString As Long, ByVal ptrToFileString As Long, ByVal ptrToParameters As Long, ByVal ptrToDirectory As Long, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

'Various API calls for manually downloading files from the Internet
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, ByRef lpszBuffer As String, ByRef lpdwBufferLength As Long) As Boolean
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal ptrToBuffer As Long, ByVal dwNumberOfBytesToRead As Long, ByRef lNumberOfBytesRead As Long) As Integer
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Boolean

'Constants used by the Windows Internet APIs
Public Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const HTTP_QUERY_CONTENT_LENGTH As Long = 5

'Some PhotoDemon functions are capable of timing themselves.  GetTickCount is used to do this.
Public Declare Function GetTickCount Lib "kernel32" () As Long

'RGB <-> HSL conversion
Public Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Public Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Public Declare Function ColorAdjustLuma Lib "shlwapi" (ByVal clrRGB As Long, ByVal n As Long, ByVal fScale As Long) As Long

'LockWindowUpdate has many purposes, but I primarily use it to add items to a listbox without a visual refresh occurring
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

'MoveWindow is used to seamlessly reposition windows as necessary
Public Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'API for drawing colored rectangles
Public Declare Function SetRect Lib "user32" (lpRect As RECTL, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECTL, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECTL, ByVal hBrush As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Wait for external actions to finish by using Sleep
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Retrieve various system metrics.  (Constants for this function are typically declared in the module where they are relevant.)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'Hook functions generally need to return the value of this function if they don't process a given message
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

'PhotoDemon allows the user to resize some toolboxes to their liking.  These constants are used with the SendMessage API
' to enable this behavior.  (In the case of the image tabstrip, note that it can be aligned to any window edge, which
' means the inside edge type can change; we thus have to support all edge types in order to interact with SendMessage
' regardless of tabstrip alignment.)
Public Const WM_NCLBUTTONDOWN As Long = &HA1
Public Const HTLEFT As Long = 10
Public Const HTTOP As Long = 12
Public Const HTRIGHT As Long = 11
Public Const HTBOTTOM As Long = 15

'Sometimes it is necessary to realign the cursor, particularly while painting
Public Declare Sub SetCursorPos Lib "user32" (ByVal newX As Long, ByVal newY As Long)
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'This painting struct stores the data passed between BeginPaint and EndPaint
Public Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate  As Long
    rgbReserved(0 To 31) As Byte
End Type

Public Declare Function GetUpdateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
