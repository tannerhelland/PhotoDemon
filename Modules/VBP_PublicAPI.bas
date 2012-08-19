Attribute VB_Name = "Public_API"
'Any and all *publicly* necessary API declarations can be found here
' (Note 1: privately declared API calls have been left in their respective forms/modules)
' (Note 2: it makes more sense to keep API-related constants here than in the constants
'  module, so don't be surprised to find constants here)
' (Note 3: I haven't searched every form/module/class for duplicate API calls... yet)

Option Explicit

'Drawing calls
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal rastOp As Long) As Boolean
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal rastOp As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDestDC As Long, ByVal nStretchMode As Long) As Long
Public Const STRETCHBLT_COLORONCOLOR As Long = 3
Public Const STRETCHBLT_HALFTONE As Long = 4

'API calls for explicitly calling dlls.  This allows us to build DLL paths at runtime, and it also allows
' us to call any DLL we like without first passing them through regsvr32.
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'ShellExecute is preferable to VB's 'Shell' command; I use it for two items in the "Help" menu - sending
' me an email, and opening the PhotoDemon website (currently just tannerhelland.com)
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

'Various API calls for manually downloading files from the Internet
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Boolean

'Constants used by the Windows Internet APIs
Public Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Public Const INTERNET_FLAG_EXISTING_CONNECT As Long = &H20000000
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const HTTP_QUERY_CONTENT_LENGTH As Long = 5

'Some PhotoDemon functions are capable of timing themselves.  GetTickCount is used to do this.
Public Declare Function GetTickCount Lib "kernel32" () As Long

