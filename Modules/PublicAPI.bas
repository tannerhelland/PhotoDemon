Attribute VB_Name = "Public_API"
'Any and all *publicly* necessary API declarations can be found here.  My current goal is to remove as many of these as possible,
' in favor of local declarations.  (This is especially true for A/W variants, as we sometimes need to switch between them
' depending on whether we're interacting with VB windows or windows we've created ourselves.)

Option Explicit

'These functions are used to interact with various windows
Public Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

