Attribute VB_Name = "Public_Variables"

Option Explicit

'Contains any and all publicly-declared variables.  I am trying to move
' all public variables to here for obvious reasons, but the transition may
' not be completely done as long as this comment remains!



'Progress bar class
Public cProgBar As cProgressBar

'Color variables
Public EmbossEngraveColor As Long 'last used emboss/engrave color

'Rotation variables
Public Type TDPoint
    x As Integer
    y As Integer
    z As Integer
End Type
Public rm(0 To 3, 0 To 3) As Double 'rotation matrix


'Filter variables
'The array containing the filter data
Public FM() As Long
'The size (1x1, 3x3, 5x5, etc) of the filter array
Public FilterSize As Byte
'The weight (i.e. / by)
Public FilterWeight As Long
'The bias (i.e. +/-)
Public FilterBias As Long
'Area filter data
Public FTransfer() As Long
Public tData() As Byte

'Picture dimensions
Public PicWidthL As Long
Public PicHeightL As Long

'Others

'Zoom data
Public Type ZoomData
    ZoomCount As Byte
    ZoomArray() As Double
    ZoomFactor() As Byte
End Type

Public Zoom As ZoomData

'Whether or not to resize large images to fit on-screen (preference is stored in the INI file; 0 means "yes," 1 means "no")
Public AutosizeLargeImages As Long

'The path where this application is located
Public ProgramPath As String

'DLLs and related support libraries
Public PluginPath As String

'Command line (used here for processing purposes)
Public CommandLine As String

'Commonly used loop variables
Public x As Long
Public y As Long
Public z As Long

'The file size of the image
Public BitmapSize As Single

'Name of file to save (necessary because forms may take control and we need something to track the file in question)
Public SaveFileName As String

'Whether or not scanning is enabled
Public ScanEnabled As Boolean

'Whether or not zLib is enabled
Public zLibEnabled As Boolean

'Whether or not FreeImage is enabled
Public FreeImageEnabled As Boolean

'Whether or not the user has created a custom filter
Public HasCreatedFilter As Boolean

'How to draw the background of image forms; -1 is checkerboard, any other value is treated as an RGB long
Public CanvasBackground As Long

'Whether or not to log program messages in a separate file - this is useful for debugging
Public LogProgramMessages As Boolean

'Whether or not we are running in the IDE or compiled
Public IsProgramCompiled As Boolean

'Temporary loading variable to disable Autozoom feature
Public FixScrolling As Boolean

'For the Open and Save common dialog boxes, it's polite to remember what format the user used last, then default
' the boxes to that.  (Note that these values are stored in the INI file as well, but that is only accessed
' upon program load and unload
Public LastOpenFilter As Long
Public LastSaveFilter As Long


