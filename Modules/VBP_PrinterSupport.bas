Attribute VB_Name = "Printing"
'***************************************************************************
'Printer support functions
'Copyright ©2003-2013 by Tanner Helland
'Created: 4/April/03
'Last updated: 12/November/13
'Last update: moved a bunch of functions out of the print dialog and into this support module.
'
'This module includes code based off an article written by Cassandra Roads of Professional Logics Corporation (PLC).
' You can download the original, unmodified version of Cassandra's code from this link (good as of 12 Nov 2013):
' http://www.tek-tips.com/faqs.cfm?fid=3603
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Per its name, DeviceCapabilities is used to retrieve printer capabilities
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, ByVal lpDevMode As Long) As Long

'Public Type POINTAPI
'    x As Long
'    y As Long
'End Type

Private Const DC_PAPERS As Long = 2
Private Const DC_PAPERSIZE As Long = 3
Private Const DC_PAPERNAMES As Long = 16

'Use the API to retrieve all supported paper sizes for the current printer
Public Function getPaperSizes(ByVal printerIndex As Long, ByRef paperSizeNames() As String, ByRef paperIDs() As Integer, ByRef exactPaperSizes() As POINTAPI) As Boolean

    'We're going to use the printer name and port frequently, so cache their names in advance
    Dim pName As String, pPort As String
    pName = Printers(printerIndex).DeviceName
    pPort = Printers(printerIndex).Port

    'Start by retrieving the paper size count; we need this to prep all our arrays
    Dim numOfPaperSizes As Long
    numOfPaperSizes = DeviceCapabilities(pName, pPort, DC_PAPERNAMES, ByVal vbNullString, 0)
    
    'Prep the various size-related arrays
    ReDim paperSizeNames(0 To numOfPaperSizes - 1) As String
    ReDim paperIDs(0 To numOfPaperSizes - 1) As Integer
    ReDim exactPaperSizes(0 To numOfPaperSizes - 1) As POINTAPI
    
    'Paper size names are returned as one giant-ass string.  Each individual name occupies 64 characters, and each
    ' is null-terminated (unless it consumes all 64 characters, in which case we have to terminate it manually).
    Dim giantPaperNameList As String
    giantPaperNameList = String(numOfPaperSizes * 64, 0)
    
    DeviceCapabilities pName, pPort, DC_PAPERNAMES, ByVal giantPaperNameList, 0
    
    'Now we have to manually parse the returned string into the array
    Dim i As Long
    Dim tmpString As String
    
    For i = 0 To numOfPaperSizes - 1
        tmpString = Mid$(giantPaperNameList, (i * 64) + 1, 64)
        tmpString = TrimNull(tmpString)
        paperSizeNames(i) = tmpString
    Next i
    
    'Next comes the matching list of paper size IDs.  See the matching list of dmPaperSize constants at:
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/dd183565%28v=vs.85%29.aspx
    DeviceCapabilities pName, pPort, DC_PAPERS, paperIDs(0), 0
    
    'Next comes the list of paper widths and heights.  These are mm-accurate measurements of each paper size,
    ' which is hugely helpful for rendering our print preview accurately.
    DeviceCapabilities pName, pPort, DC_PAPERSIZE, exactPaperSizes(0), 0
        
    getPaperSizes = True

End Function
        

