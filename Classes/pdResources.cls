VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Resource Manager
'Copyright 2016-2016 by Tanner Helland
'Created: 13/December/16
'Last updated: 27/December/16
'Last update: add debug code for testing
'
'PhotoDemon needs to include a whole swatch of custom resources.  These resources take up a lot of space,
' and we also need to pull different resources depending on things like screen DPI.  To simplify this process,
' we manage resources manually, rather than relying on an external .rc file to do it for us.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'If a resource file was loaded successfully, this will be set to TRUE.  You *must* check this value before
' attempting to retrieve individual resources.
Private m_ResourcesAvailable As Boolean

'Resources are stored in a persistent pdPackager instance.  You cannot rely on instantiation of this class
' for correct behavior; instead, query m_ResourcesAvailable.
Private m_Resources As pdPackager2

'We use a temporary DIB to pull resources from file, prior to applying any real-time coloring changes.
Private m_TmpDIB As pdDIB

'Called early in PD's startup, this function checks to see if we've already extracted PD's resource file to a
' separate working folder.  If we have, grab the data from there instead of from the .exe.  (This is helpful
' during debugging.)
'
'In production builds, the initial collection is pulled from the .exe itself.
Public Function LoadInitialResourceCollection() As Boolean

    #If DEBUGMODE = 1 Then
    
        'In debug builds, we load the resource from a static location instead of pulling it from the .exe
        Dim srcResFile As String
        srcResFile = g_UserPreferences.GetThemePath & "Core_Icons.pdrc"
        LoadInitialResourceCollection = LoadResourceFile(srcResFile)
        
    #End If
    
    'TODO!  Someday we'll build the (finished) resource file directly into the .exe, but for now,
    ' load it from a static debug folder.
    
End Function

'Given a path to a .pdr file, attempt to load and validate it
Public Function LoadResourceFile(ByRef srcPath As String) As Boolean
    Set m_Resources = New pdPackager2
    m_ResourcesAvailable = m_Resources.ReadPackageFromFile(srcPath, PD_RES_IDENTIFIER, PD_SM_MemoryBacked)
    LoadResourceFile = m_ResourcesAvailable
End Function

Public Function AreResourcesAvailable() As Boolean
    AreResourcesAvailable = m_ResourcesAvailable
End Function

Public Function LoadImageResource(ByRef imgResName As String, ByRef dstDIB As pdDIB, Optional ByVal desiredWidth As Long = 0, Optional ByVal desiredHeight As Long = 0, Optional ByVal desiredBorders As Single = 0#, Optional ByVal dstIsMenu As Boolean = False) As Boolean
    
    LoadImageResource = False
    
    If m_ResourcesAvailable Then
    
        'See if this resource exists in the collection.  (It may not, as I'm still in the process of migrating
        ' all image resources to the new resource format.)
        Dim nodeIndex As Long
        nodeIndex = m_Resources.GetNodeIndexFromName(imgResName)
        
        '-1 is returned if the node doesn't exist in the collection
        If (nodeIndex >= 0) Then
            
            'Retrieve the header XML packet for this node
            Dim retBytes() As Byte, retString As String, retSize As Long
            If m_Resources.GetNodeDataByIndex(nodeIndex, True, retBytes, False, retSize) Then
                
                'Copy the received bytes into a string, then initialize an XML object from it
                retString = Space$(retSize \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), retSize
                
                Dim cXML As pdXML: Set cXML = New pdXML
                If cXML.LoadXMLFromString(retString) Then
                
                    'Retrieve the image's dimensions
                    Dim imgWidth As Long, imgHeight As Long, imgBPP As Long
                    imgWidth = cXML.GetUniqueTag_Long("w")
                    imgHeight = cXML.GetUniqueTag_Long("h")
                    imgBPP = cXML.GetUniqueTag_Long("bpp")
                    
                    'Prep a temporary DIB (as we may need to resize the DIB to meet the user's request)
                    If (m_TmpDIB Is Nothing) Then Set m_TmpDIB = New pdDIB
                    
                    Dim needToCreateDIB As Boolean: needToCreateDIB = False
                    If (m_TmpDIB.GetDIBWidth <> imgWidth) Then
                        needToCreateDIB = True
                    ElseIf (m_TmpDIB.GetDIBHeight <> imgHeight) Then
                        needToCreateDIB = True
                    ElseIf (m_TmpDIB.GetDIBColorDepth <> imgBPP) Then
                        needToCreateDIB = True
                    End If
                    
                    If needToCreateDIB Then m_TmpDIB.CreateBlank imgWidth, imgHeight, imgBPP, 0, 0
                    
                    'We now have to use one of two strategies to retrieve the DIB
                    ' 1) Pull the actual DIB bits out of the file.  This is only an option if the image
                    '    does *not* support run-time coloration.
                    ' 2) Pull the DIB's alpha channel only out of the file.  If this image *does* support
                    '    run-time coloration, there will not be any color data inside the file.
                    If cXML.GetUniqueTag_Boolean("rt-clr") Then
                        Dim srcBytes() As Byte
                        ReDim srcBytes(0 To imgWidth - 1, 0 To imgHeight - 1) As Byte
                        LoadImageResource = m_Resources.GetNodeDataByIndex_UnsafeDstPointer(nodeIndex, False, VarPtr(srcBytes(0, 0)))
                        If LoadImageResource Then LoadImageResource = DIB_Support.ApplyTransparencyTable(m_TmpDIB, srcBytes)
                    Else
                        Dim tmpDIBPointer As Long, tmpDIBLength As Long
                        m_TmpDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                        LoadImageResource = m_Resources.GetNodeDataByIndex_UnsafeDstPointer(nodeIndex, False, tmpDIBPointer)
                    End If
                    
                    'At present, all resources contain premultiplied alpha, so force the corresponding state now
                    m_TmpDIB.SetInitialAlphaPremultiplicationState True
                    
                    'See if the DIB requires real-time coloration
                    
                    'Resize the DIB into the destination
                    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
                    If (desiredWidth = 0) Then desiredWidth = 16#
                    If (desiredHeight = 0) Then desiredHeight = 16#
                    
                    Dim desiredWidthInt As Long, desiredHeightInt As Long
                    desiredWidthInt = Int(desiredWidth + 0.9999)
                    desiredHeightInt = Int(desiredHeight + 0.9999)
                    
                    If (dstDIB.GetDIBWidth <> desiredWidthInt) Or (dstDIB.GetDIBHeight <> desiredHeightInt) Then
                        dstDIB.CreateBlank desiredWidthInt, desiredHeightInt, 32, 0, 0
                    Else
                        dstDIB.ResetDIB 0
                    End If
                    
                    dstDIB.SetAlphaPremultiplication True
                    GDI_Plus.GDIPlus_StretchBlt dstDIB, desiredBorders, desiredBorders, desiredWidth - desiredBorders * 2, desiredHeight - desiredBorders * 2, m_TmpDIB, 0, 0, imgWidth, imgHeight, , GP_IM_HighQualityBicubic, , , , True
                    
                    'If the image supports real-time coloration, apply it now (based on the currently selected theme).
                    If cXML.GetUniqueTag_Boolean("rt-clr") Then
                    
                        'Retrieve the image's color.  (NOTE: the specified color will be overridden with
                        ' monochrome if the monochrome icon preference is active.)
                        ' TODO: implement monochrome handling here
                        Dim targetColor As Long
                        If ((g_Themer.GetCurrentThemeClass = PDTC_Light) Or dstIsMenu) Then
                            targetColor = cXML.GetUniqueTag_Long("clr-l")
                        ElseIf (g_Themer.GetCurrentThemeClass = PDTC_Dark) Then
                            targetColor = cXML.GetUniqueTag_Long("clr-d")
                        End If
                        
                        DIB_Support.ColorizeDIB dstDIB, targetColor
                    
                    End If
                    
                End If
                
            End If
            
        End If
    
    End If
    
End Function

Private Sub Class_Initialize()
    m_ResourcesAvailable = False
End Sub
