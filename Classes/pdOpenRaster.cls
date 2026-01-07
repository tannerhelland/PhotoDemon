VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdOpenRaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon OpenRaster Container and Parser
'Copyright 2019-2026 by Tanner Helland
'Created: 08/January/19
'Last updated: 05/January/23
'Last update: update to latest cZipArchive; minor changes to match
'
'OpenRaster is currently the "best" mechanism for sharing multi-layer images between PhotoDemon
' and other open-source photo editors (GIMP, Krita, MyPaint, etc).
'
'The OpenRaster format is described in detail at Wikipedia:
' https://en.wikipedia.org/wiki/OpenRaster
'
'The formal spec is available here (link good as of January 2019):
' https://www.openraster.org/
'
'This class requires a copy of cZipArchive, an MIT-licensed zip library by wqweto@gmail.com.
' Many thanks to wqweto for not only sharing his class under a permissive license, but also being
' very responsive to bug reports and feature requests.  An original, un-altered copy of cZipArchive
' can be downloaded from GitHub, as can its attached MIT license (link good as of January 2019):
' https://github.com/wqweto/ZipArchive
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Some of our own internal parsers can perform perf timings;
' enable this value to send format-specific timing reports to PDDebug.
Private Const GENERATE_PERF_REPORTS As Boolean = False

'cZipArchive handles all zip extraction duties
Private m_ZipArchive As cZipArchive

'When loading a (valid) ORA file, a basic header with key image attributes gets populated first.
' Note that images "in the wild" may dump a bunch of undocumented attributes into the header as well;
' I currently make no attempt to understand or preserve such settings (looking at you, MyPaint).
Private Type OpenRasterHeader
    orVersion As String 'Semantic version string
    orWidth As Long     'Image width; child layer width may be </> than this
    orHeight As Long    'Image height; child layer width may be </> than this
    orXRes As Single    'X-resolution, in PPI.  Defaults to 72.
    orYRes As Single    'X-resolution, in PPI.  Defaults to 72.
End Type

Private m_Header As OpenRasterHeader

'In a valid ORA file, layers can potentially possess many "attributes".  PhotoDemon supports most but not
' all possible properties (e.g. some blend-modes are not currently implemented).  This "PD-friendly" struct
' holds any attributes we currently support, with everything translated to its most-similar feature in PD.
Private Type OpenRasterLayer
    orlName As String
    orlX As Long
    orlY As Long
    orlOpacity As Single
    orlVisibility As Boolean
    orlBlendMode As PD_BlendMode
    orlSelected As Boolean      'Not in the spec, but appears to describe a currently active layer in Krita;
                                ' we use it to set the active layer, if found.
    orlSourceFile As String     'Used to know which PNG file to extract from the resources section,
                                ' but we don't store this attribute outside of this class
End Type

'When validating an ORA file, we store the original filename that was validated; if the validation succeeds,
' we can reuse this same m_ZipArchive instance for actual parsing duties.
Private m_Filename As String

'Validate a source filename as ORA format.  Validation *does* touch the file - we actually open the first
' file in the archive to make sure it meets the required mimetype descriptor.
Friend Function IsFileORA(ByRef srcFilename As String, Optional ByVal requireValidFileExtension As Boolean = True) As Boolean
    
    If (m_ZipArchive Is Nothing) Then Set m_ZipArchive = New cZipArchive
    
    Dim potentiallyORA As Boolean
    potentiallyORA = True
    
    'Check extension up front, if requested
    If requireValidFileExtension Then potentiallyORA = Strings.StringsEqual(Files.FileGetExtension(srcFilename), "ora", True)
    
    'Proceed with deeper validation as necessary
    If potentiallyORA Then
        
        'Attempt to load the archive
        potentiallyORA = m_ZipArchive.OpenArchive(srcFilename)
        If potentiallyORA Then
        
            'The file appears to be a valid zip archive.  Retrieve a list of included filenames;
            ' we're looking for "mimetype" as the first entry.
            Const IDX_FILENAME As Long = 0
            potentiallyORA = Strings.StringsEqual(CStr(m_ZipArchive.FileInfo(IDX_FILENAME, zipIdxFileName)), "mimetype", False)
            
            'If the mimetype file is present, extract it and ensure the contents are, per the spec:
            ' "...the string "image/openraster", with no whitespace or trailing newline."
            If potentiallyORA Then
                
                Dim utf8Bytes() As Byte
                potentiallyORA = m_ZipArchive.Extract(utf8Bytes, "mimetype")
                If potentiallyORA Then potentiallyORA = Strings.StringsEqual(Strings.StringFromUTF8(utf8Bytes), "image/openraster")
                
                'On success, don't free the cZipArchive instance; we want to reuse the current instance
                ' for actual file import.  Instead, note the successfully validated filename so that the
                ' loader knows to skip validation steps.
                If potentiallyORA Then m_Filename = srcFilename
                
            End If
        
        End If
        
    End If
    
    IsFileORA = potentiallyORA
    
End Function

'Before loading an actual ORA file, consider running IsFileORA(), above - this will validate the file for you,
' and you can avoid calling this function for non-ORA files.
Friend Function LoadORA(ByRef srcFilename As String, ByRef dstImage As pdImage) As Boolean
    
    LoadORA = False
    
    On Error GoTo CouldNotLoadFile
    
    'If we haven't validated the target file, do so now
    If Strings.StringsNotEqual(m_Filename, srcFilename) Then
        If (Not Me.IsFileORA(m_Filename)) Then
            InternalError "LoadORA", "target file isn't in OpenRaster format"
            Exit Function
        End If
    End If
    
    Dim startTime As Currency, firstTime As Currency
    Dim zipTime As Double, pngTime As Double, postTime As Double
    If GENERATE_PERF_REPORTS Then
        VBHacks.GetHighResTime firstTime
        VBHacks.GetHighResTime startTime
    End If
    
    'The validator will have already loaded the target file, so this exists only as a failsafe.
    If (m_ZipArchive Is Nothing) Then
        Set m_ZipArchive = New cZipArchive
        m_ZipArchive.OpenArchive srcFilename
    End If
    
    'The first thing we want to do is extract the stack.xml file.  This contains crucial attributes like
    ' the image's width and height, which we need to know before we start parsing individual layers.
    Dim utf8Bytes() As Byte
    If m_ZipArchive.Extract(utf8Bytes, "stack.xml") Then
        
        'Validate and load the XML
        Dim xmlEngine As Object 'MSXML2.DOMDocument
        Set xmlEngine = CreateObject("MSXML2.DOMDocument")
        xmlEngine.async = False
        xmlEngine.validateOnParse = True
        
        Dim srcXML As String
        srcXML = Strings.StringFromUTF8(utf8Bytes)
        
        'Want to know what the source XML looks like?  MyPaint XML in particular contains a load of non-spec data,
        ' and in the future, we could potentially expand support for some of their more esoteric features.
        'Debug.Print srcXML
        
        If xmlEngine.loadXML(srcXML) Then
        
            'The root node should be <image>.
            Dim xmlRoot As Object 'MSXML2.IXMLDOMNode
            Set xmlRoot = xmlEngine.documentElement
            If (xmlRoot.Attributes.Length > 0) Then
                
                'Prep a default header; we will overwrite these attributes if/when we encounter them
                With m_Header
                    .orVersion = vbNullString
                    .orWidth = 0
                    .orHeight = 0
                    .orXRes = 72    'x/y res default to 72 per the spec; these attributes are optional, however
                    .orYRes = 72
                End With
                
                'The final validation we need to perform involves checking the root image's attributes.
                ' We need to make sure basic things are intact, like non-zero width/height
                Dim i As Long, curNodeName As String, curNodeValue As String
                For i = 0 To xmlRoot.Attributes.Length - 1
                    
                    curNodeName = LCase$(xmlRoot.Attributes(i).nodeName)
                    curNodeValue = xmlRoot.Attributes(i).NodeValue
                    
                    Select Case curNodeName
                        Case "version"
                            m_Header.orVersion = curNodeValue
                        Case "w"
                            m_Header.orWidth = CLng(curNodeValue)
                        Case "h"
                            m_Header.orHeight = CLng(curNodeValue)
                        Case "xres"
                            m_Header.orXRes = TextSupport.CDblCustom(curNodeValue)
                        Case "yres"
                            m_Header.orYRes = TextSupport.CDblCustom(curNodeValue)
                        'The spec doesn't define any other image attributes, but who knows what we'll
                        ' encounter in the wild!
                        Case Else
                            InternalError "LoadORA", "Unknown <image> attribute encountered: " & curNodeName & "(" & curNodeValue & ")"
                    End Select
                    
                Next i
                
                'Make sure we found a valid width and height.  (The spec wants a version check as well,
                ' but this seems overkill if everything else is valid.)
                If (m_Header.orWidth > 0) And (m_Header.orHeight > 0) Then
                
                    'The image passes validation!  This appears to be a valid OpenRaster file.
                    With m_Header
                        dstImage.Width = .orWidth
                        dstImage.Height = .orHeight
                        dstImage.SetDPI .orXRes, .orYRes
                    End With
                    
                    ' We can now begin the messy business of traversing the rest of the stack.xml file
                    ' and extracting associated PNG data as we go.
                    
                    'TODO: the spec allows for layer grouping ("stacks"), which PhotoDemon does not currently support.
                    ' There's no easy workaround for this until we implement groups, and I have no ETA for this at present.
                    ' As such, I currently ignore stacks and simply iterate layer entries.  This will still generate
                    ' a correct image for the majority of "in the wild" ORA files.
                    
                    'Because of this, instead of traversing the DOM we can simply grab a list of layer nodes.
                    ' Without grouping support, the only thing we really care about is order, not nesting.
                    Dim listOfLayers As Object 'MSXML2.IXMLDOMNodeList
                    Set listOfLayers = xmlEngine.getElementsByTagName("layer")
                    
                    'Traverse the list of layers in *reverse* order, as the last layer in the file is the base layer,
                    ' while the first layer is the top-most one.
                    Dim layerProps() As OpenRasterLayer
                    ReDim layerProps(0 To listOfLayers.Length - 1) As OpenRasterLayer
                    
                    Dim j As Long
                    For i = listOfLayers.Length - 1 To 0 Step -1
                    
                        With layerProps(i)
                        
                            For j = 0 To listOfLayers(i).Attributes.Length - 1
                                
                                curNodeName = LCase$(listOfLayers(i).Attributes(j).nodeName)
                                curNodeValue = listOfLayers(i).Attributes(j).NodeValue
                                
                                Select Case curNodeName
                                    
                                    'Spec-defined attributes follow
                                    Case "name"
                                        .orlName = Strings.UnEscapeXMLString(curNodeValue)
                                    Case "opacity"
                                        .orlOpacity = TextSupport.CDblCustom(curNodeValue) * 100#
                                    Case "src"
                                        .orlSourceFile = Strings.UnEscapeXMLString(curNodeValue)
                                    Case "visibility"
                                        .orlVisibility = Strings.StringsEqual(curNodeValue, "visible", True)
                                    Case "x"
                                        .orlX = CLng(TextSupport.CDblCustom(curNodeValue))
                                    Case "y"
                                        .orlY = CLng(TextSupport.CDblCustom(curNodeValue))
                                    Case "composite-op"
                                        .orlBlendMode = GetBlendModeFromSVGOp(curNodeValue)
                                    
                                    'Non-spec defined attributes found in the wild follow
                                    
                                    'Krita uses "selected" to mark active layers; PD only supports one "selected"
                                    ' layer at a time, so we'll end up grabbing the *lowest in the stack* selected
                                    ' layer for the time being.
                                    Case "selected"
                                        .orlSelected = Strings.StringsEqual(curNodeValue, "true", True)
                                    
                                    'The spec doesn't define any other layer attributes, but who knows what we'll
                                    ' encounter in the wild!
                                    Case Else
                                        InternalError "LoadORA", "Unknown <layer> attribute encountered: " & curNodeName & "(" & curNodeValue & ")"
                                End Select
                                
                            Next j
                        End With
                        
                    Next i
                    
                    If GENERATE_PERF_REPORTS Then
                        zipTime = 0#
                        pngTime = 0#
                        postTime = 0#
                    End If
                    
                    'We now have all the information we need to reconstruct individual image layers.  Note that the OpenRaster spec
                    ' only allows individual layers to be embedded PNG images, but some software (MyPaint) may also embed SVG layers.
                    ' We can't parse SVGs at present, so we're SOL on such layers.
                    Dim cPNG As pdPNG, tmpPngBytes() As Byte
                    Dim tmpDIB As pdDIB, tmpLayer As pdLayer, newLayerID As Long, selLayerID As Long
                    selLayerID = -1
                    For i = listOfLayers.Length - 1 To 0 Step -1
                        
                        If GENERATE_PERF_REPORTS Then VBHacks.GetHighResTime startTime
                        
                        'Before creating a layer, we want to first attempt to retrieve the layer's PNG data.
                        ' (If this fails, we'll discard the layer entirely.)
                        
                        'Many ORA writers (e.g. GIMP, MyPaint) choose to write their PNGs to the zip container as-is,
                        ' without an additional deflate pass over-the-top.  This is a great idea as deflating a PNG
                        ' is largely pointless (little size reduction), and it adds unwanted complexity to the
                        ' extraction process because we effectively have to deflate pixel data twice (zip -> png -> final).
                        '
                        'To avoid wasting time "extracting" uncompressed PNGs into a temporary byte array, and then
                        ' decoding THAT, it makes a lot more sense to simply parse the PNG data directly from its
                        ' file offset on-disk when we can.  I've added a small "peek" function to cZipArchive to allow
                        ' us to more easily do this - but note that if the PNG file *is* deflated, we have no choice but
                        ' to inflate the data to a temporary buffer, *then* decode the resulting PNG.
                        Dim srcOffset As Long, srcSizeCompressed As Long, srcSizeOriginal As Long, srcIsDeflated As Boolean
                        If m_ZipArchive.PeekSingleFile(layerProps(i).orlSourceFile, srcOffset, srcSizeCompressed, srcSizeOriginal, srcIsDeflated) Then

                            Dim pngLoadSuccess As Boolean
                            Set tmpDIB = New pdDIB
                            Set cPNG = New pdPNG

                            'If the source data is compressed, we have no choice but to decompress it into a temporary array.
                            If srcIsDeflated Then
                                
                                If m_ZipArchive.Extract(tmpPngBytes, layerProps(i).orlSourceFile) Then
                                    
                                    If GENERATE_PERF_REPORTS Then
                                        zipTime = zipTime + VBHacks.GetTimerDifferenceNow(startTime)
                                        VBHacks.GetHighResTime startTime
                                    End If

                                    pngLoadSuccess = (cPNG.LoadPNG_Simple(vbNullString, Nothing, tmpDIB, False, VarPtr(tmpPngBytes(0)), srcSizeOriginal) = png_Success)
                                Else
                                    pngLoadSuccess = False
                                End If
                                
                            'If, however, the source data is *not* compressed (as is common for e.g. OpenRaster files
                            ' from GIMP), then it's a waste to copy the data into memory.  Instead, load the data
                            ' directly from the file into our PNG class, bypassing a temporary memory copy completely.
                            Else
                                If GENERATE_PERF_REPORTS Then
                                    zipTime = zipTime + VBHacks.GetTimerDifferenceNow(startTime)
                                    VBHacks.GetHighResTime startTime
                                End If
                                pngLoadSuccess = (cPNG.LoadPNG_Simple(srcFilename, Nothing, tmpDIB, False, offsetInSrcFile:=srcOffset) = png_Success)
                            End If
                            
                            If pngLoadSuccess Then
                                
                                If GENERATE_PERF_REPORTS Then
                                    pngTime = pngTime + VBHacks.GetTimerDifferenceNow(startTime)
                                    VBHacks.GetHighResTime startTime
                                End If
                                
                                'Because color-management has already been handled (if applicable), this is a great time to premultiply alpha
                                tmpDIB.SetAlphaPremultiplication True
                                
                                'Prep a new layer object and initialize it with the image bits we've retrieved
                                newLayerID = dstImage.CreateBlankLayer()
                                Set tmpLayer = dstImage.GetLayerByID(newLayerID)
                                tmpLayer.InitializeNewLayer PDL_Image, layerProps(i).orlName, tmpDIB
                                
                                'Fill in any remaining layer properties
                                With layerProps(i)
                                    tmpLayer.SetLayerBlendMode .orlBlendMode
                                    tmpLayer.SetLayerOpacity .orlOpacity
                                    tmpLayer.SetLayerOffsetX .orlX
                                    tmpLayer.SetLayerOffsetY .orlY
                                    tmpLayer.SetLayerVisibility .orlVisibility
                                    If .orlSelected Then selLayerID = newLayerID
                                End With
                                
                                'Notify the layer of new changes, so it knows to regenerate internal caches on next access
                                tmpLayer.NotifyOfDestructiveChanges
                                
                                If GENERATE_PERF_REPORTS Then postTime = postTime + VBHacks.GetTimerDifferenceNow(startTime)
                                
                            Else
                                InternalError "LoadORA", "Retrieved bytes for layer #" & CStr(i) & " were not PNG format; filename is: " & layerProps(i).orlSourceFile
                            End If
                            
                        Else
                            InternalError "LoadORA", "Could not extract PNG bytes for layer #" & CStr(i) & ": " & layerProps(i).orlSourceFile
                        End If
                        
                    Next i
                    
                    'All layers have been iterated.  If the target image contains at least one valid layer,
                    ' consider this a successful load.
                    LoadORA = (dstImage.GetNumOfLayers > 0)
                    If LoadORA And (selLayerID <> -1) Then dstImage.SetActiveLayerByID selLayerID
                    
                Else
                    InternalError "LoadORA", "Image width and/or height invalid (" & m_Header.orWidth & "x" & m_Header.orHeight & ")"
                End If
                
            Else
                InternalError "LoadORA", "root image node provides no attributes"
                LoadORA = False
            End If
            
        Else
            InternalError "LoadORA", "stack.xml didn't validate"
            LoadORA = False
        End If
    
    Else
        'On failure, there's nothing we can do - we lack enough information to construct even a placeholder image
        InternalError "LoadORA", "no stack.xml in file"
        LoadORA = False
    End If
    
    If GENERATE_PERF_REPORTS Then
        InternalWarning "LoadORA", "Total time to load OpenRaster file: " & VBHacks.GetTimeDiffNowAsString(firstTime)
        InternalWarning "LoadORA", "Time spent in zip extraction: " & Format$(zipTime * 1000#, "0.0") & " ms"
        InternalWarning "LoadORA", "Time spent in PNG parsing: " & Format$(pngTime * 1000#, "0.0") & " ms"
        InternalWarning "LoadORA", "Time spent in post-processing: " & Format$(postTime * 1000#, "0.0") & " ms"
    End If
    
    Set m_ZipArchive = Nothing
    
    Exit Function
    
CouldNotLoadFile:
    InternalError "LoadORA", "Internal VB error #" & Err.Number & ": " & Err.Description
    LoadORA = False

End Function

'Save a new OpenRaster file to disk.  (With minor modifications, this could also be used to save to a memory stream,
' but this is *not* currently implemented.)
Friend Function SaveORA(ByRef srcPDImage As pdImage, ByVal dstFile As String) As Boolean
    
    On Error GoTo CouldNotSaveFile
    
    If (srcPDImage Is Nothing) Or (LenB(dstFile) = 0) Then
        InternalError "SaveORA", "invalid function parameters"
        SaveORA = False
        Exit Function
    End If
    
    'Reset the archiver
    Set m_ZipArchive = New cZipArchive
    
    'Setting codepage here is not strictly necessary as the final .CompressArchive call accepts a UTF-8 parameter,
    ' but the OpenRaster spec is explicit on this point (see https://www.openraster.org/baseline/file-layout-spec.html)
    ' so I've added this as a reminder to myself.
    Const CP_UTF8 As Long = 65001
    m_ZipArchive.CodePage = CP_UTF8
    
    'The first entry in the archive must always be a file named "mimetype", stored without compression,
    ' with string content "image/openraster"
    Dim utf8Bytes() As Byte, lenInBytes As Long
    Strings.UTF8FromString "image/openraster", utf8Bytes, lenInBytes, 0, True
    If Not m_ZipArchive.AddFile(File:=utf8Bytes, Name:="mimetype", Level:=0&) Then InternalError "SaveORA", "could not write mimetype"
    
    'Next comes a file called "stack.xml".  This XML file contains all layer attribute data.
    ' Rather than use a full-blown XML manager, let's just use a lightweight stringbuilder to construct the file.
    Dim cString As pdString
    Set cString = New pdString
    cString.AppendLine "<?xml version='1.0' encoding='UTF-8'?>"
    
    'The root tag is an <image> tag; it defines the underlying image's dimensions and DPI, as well as a
    ' (currently meaningless) OpenRaster version number.
    cString.Append "<image version=""0.0.5"" w="""
    cString.Append CStr(srcPDImage.Width)
    cString.Append """ h="""
    cString.Append CStr(srcPDImage.Height)
    cString.Append """ xres="""
    cString.Append CStr(srcPDImage.GetDPI)
    cString.Append """ yres="""
    cString.Append CStr(srcPDImage.GetDPI)
    cString.AppendLine """>"""
    cString.AppendLine vbTab & "<stack>"
    
    'Next, we need to write out a header for each layer entry.  Note that we weirdly write these in
    ' *reverse* order, as the spec requires you to list layers in top-down order.  (While at it,
    ' we also invent a PNG "file path" identifier for each layer, based on its index.)
    Dim xOffset As Long, yOffset As Long
    Dim tmpDIB As pdDIB, pathInsideZip As String
    
    Dim i As Long, numLayersBound As Long
    numLayersBound = srcPDImage.GetNumOfLayers - 1
    For i = numLayersBound To 0 Step -1
        
        If Not (srcPDImage.GetLayerByIndex(i) Is Nothing) Then
            With srcPDImage.GetLayerByIndex(i)
                
                'It sounds funny, but if the target layer has non-destructive transformations active,
                ' we actually need to grab a copy of its image data before writing it out to file.
                ' The reason for this is that most other software does not support non-destructive editing,
                ' which means we need to make temporary destructive copies of this layer, which are likely
                ' to be offset to a different (x, y) position.
                If .AffineTransformsActive(True) Then
                    .GetAffineTransformedDIB tmpDIB, xOffset, yOffset
                Else
                    Set tmpDIB = New pdDIB
                    tmpDIB.CreateFromExistingDIB .GetLayerDIB
                    xOffset = .GetLayerOffsetX
                    yOffset = .GetLayerOffsetY
                End If
                
                'Saved PNGs must (obviously) contain un-premultiplied alpha
                tmpDIB.SetAlphaPremultiplication False
                
                'Add the PNG stream to the zip file
                pathInsideZip = "data/" & Format$(i, "000") & ".png"
                If AddImageToZipAsPNG(tmpDIB, pathInsideZip) Then
                
                    'The PNG was added successfully.  Add its header to the stack XML.
                    cString.Append String$(2, vbTab) & "<layer src="""
                    cString.Append pathInsideZip
                    cString.Append """ name="""
                    cString.Append Strings.EscapeXMLString(.GetLayerName)
                    cString.Append """ visibility="""
                    cString.Append IIf(.GetLayerVisibility, "visible", "hidden")
                    cString.Append """ opacity="""
                    cString.Append Format$((.GetLayerOpacity * 0.01), "0.0###")
                    cString.Append """ x="""
                    cString.Append Trim$(Str$(xOffset))
                    cString.Append """ y="""
                    cString.Append Trim$(Str$(yOffset))
                    cString.Append """ composite-op="""
                    cString.Append GetSVGOpFromBlendMode(.GetLayerBlendMode)
                    If (srcPDImage.GetActiveLayerIndex = i) Then cString.Append """ selected=""true"
                    cString.AppendLine """/>"
                    
                Else
                    InternalError "SaveORA", "Couldn't add layer """ & .GetLayerName & """ to zip file; omitting layer from stack.xml to match."
                End If
                
            End With
        End If
        
    Next i
    
    'Close all remaining tags
    cString.AppendLine vbTab & "</stack>"
    cString.AppendLine "</image>"
    
    'Convert the string to UTF-8 (in case there are Unicode chars in places like layer names) and add it to the zip
    Strings.UTF8FromString cString.ToString(), utf8Bytes, lenInBytes, , True
    If Not m_ZipArchive.AddFile(File:=utf8Bytes, Name:="stack.xml") Then InternalError "SaveORA", "WARNING!  ExportORA could not write stack.xml."
    
    'Next, we need to add a thumbnail image to the zip.  This is basically identical to the previous step
    ' of adding individual layers to the zip file as PNGs: generate a thumbnail DIB, save it to a PNG bytestream,
    ' and add the bytestream to the file at a hard-coded location.
    srcPDImage.RequestThumbnail tmpDIB, 256, False
    tmpDIB.SetAlphaPremultiplication False
    AddImageToZipAsPNG tmpDIB, "Thumbnails/thumbnail.png"
    
    'Finally, we need to add a fully composited copy of the image.
    srcPDImage.GetCompositedImage tmpDIB, False
    AddImageToZipAsPNG tmpDIB, "mergedimage.png"
    
    'And we're done!  Write the finished archive out to disk.
    SaveORA = m_ZipArchive.CompressArchive(ArchiveFile:=dstFile, useUTF8:=vbTrue, UseZip64:=vbFalse)
    Set m_ZipArchive = Nothing
    
    Exit Function
    
CouldNotSaveFile:
    InternalError "SaveORA", "Internal VB error #" & Err.Number & ": " & Err.Description
    SaveORA = False
    
End Function

Private Function AddImageToZipAsPNG(ByRef srcDIB As pdDIB, ByRef pathInsideZip As String) As Boolean

    AddImageToZipAsPNG = False
    
    Dim pngBytes() As Byte
    
    'We preferentially use our own internal PNG encoder, as it's both faster than FreeImage *and* it
    ' produces smaller files.  (If for some reason the export fails, we'll silently fall back to GDI+.
    ' GDI+ adds a bunch of unwanted PNG chunks - sRGB, cHRM - which fuck up the subsequent PNG's display
    ' on color-managed software, and it compresses terribly - but that pathway is only an emergency
    ' fallback, so it should never trigger!)
    Dim cPNG As pdPNG
    Set cPNG = New pdPNG
    If (cPNG.SavePNG_ToMemory(pngBytes, srcDIB, Nothing, png_AutoColorType, 0, 9, vbNullString, png_FilterAuto) <> 0) Then

        'Add the PNG stream to the zip file.  At present, we mimic GIMPs implementation and do *NOT* request
        ' DEFLATE compression.  This is because the PNG file is already compressed using DEFLATE, and an
        ' additional compression atop that gains little, but eats up a ton of time.
        AddImageToZipAsPNG = m_ZipArchive.AddFile(File:=pngBytes, Name:=pathInsideZip, Level:=0&)
        If (Not AddImageToZipAsPNG) Then InternalError "AddImageToZipAsPNG", "Couldn't add """ & pathInsideZip & """ to zip file"
        
    Else
        InternalError "AddImageToZipAsPNG", "Couldn't generate PNG for """ & pathInsideZip & """"
    End If

End Function

Private Function GetBlendModeFromSVGOp(ByRef srcString As String) As PD_BlendMode

    Select Case LCase$(srcString)
        
        'Normal / Source Over
        Case "svg:src-over"
            GetBlendModeFromSVGOp = BM_Normal
            
        'Multiply / Source Over
        Case "svg:multiply"
            GetBlendModeFromSVGOp = BM_Multiply
        
        'Screen / Source Over
        Case "svg:screen"
            GetBlendModeFromSVGOp = BM_Screen
        
        'Overlay / Source Over
        Case "svg:overlay"
            GetBlendModeFromSVGOp = BM_Overlay
        
        'Darken / Source Over
        Case "svg:darken"
            GetBlendModeFromSVGOp = BM_Darken
        
        'Lighten / Source Over
        Case "svg:lighten"
            GetBlendModeFromSVGOp = BM_Lighten
            
        'Color Dodge / Source Over
        Case "svg:color-dodge"
            GetBlendModeFromSVGOp = BM_ColorDodge
        
        'Color Burn / Source Over
        Case "svg:color-burn"
            GetBlendModeFromSVGOp = BM_ColorBurn
        
        'Hard Light / Source Over
        Case "svg:hard-light"
            GetBlendModeFromSVGOp = BM_HardLight
        
        'Soft Light / Source Over
        Case "svg:soft-light"
            GetBlendModeFromSVGOp = BM_SoftLight
        
        'Difference / Source Over
        Case "svg:difference"
            GetBlendModeFromSVGOp = BM_Difference
        
        'Color / Source Over
        Case "svg:color"
            GetBlendModeFromSVGOp = BM_Color
        
        'Luminosity / Source Over
        Case "svg:luminosity"
            GetBlendModeFromSVGOp = BM_Luminosity
        
        'Hue / Source Over
        Case "svg:hue"
            GetBlendModeFromSVGOp = BM_Hue
        
        'Saturation / Source Over
        Case "svg:saturation"
            GetBlendModeFromSVGOp = BM_Saturation
        
        'Normal / Lighter
        Case "svg:plus"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BM_Normal
        
        'Normal / Destination In
        Case "svg:dst-in"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BM_Normal
        
        'Normal / Destination Out
        Case "svg:dst-out"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BM_Normal
        
        'Normal / Source Atop
        Case "svg:src-atop"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BM_Normal
        
        'Normal / Destination Atop
        Case "svg:dst-atop"
            InternalError "GetBlendModeFromSVGOp", "unsupported blend mode: " & srcString
            GetBlendModeFromSVGOp = BM_Normal
        
        Case Else
            InternalError "GetBlendModeFromSVGOp", "unknown blend mode: " & srcString
            GetBlendModeFromSVGOp = BM_Normal
    
    End Select

End Function

Private Function GetSVGOpFromBlendMode(ByVal srcBlendMode As PD_BlendMode) As String
    Select Case srcBlendMode
        Case BM_Normal
            GetSVGOpFromBlendMode = "svg:src-over"
        Case BM_Multiply
            GetSVGOpFromBlendMode = "svg:multiply"
        Case BM_Screen
            GetSVGOpFromBlendMode = "svg:screen"
        Case BM_Overlay
            GetSVGOpFromBlendMode = "svg:overlay"
        Case BM_Darken
            GetSVGOpFromBlendMode = "svg:darken"
        Case BM_Lighten
            GetSVGOpFromBlendMode = "svg:lighten"
        Case BM_ColorDodge
            GetSVGOpFromBlendMode = "svg:color-dodge"
        Case BM_ColorBurn
            GetSVGOpFromBlendMode = "svg:color-burn"
        Case BM_HardLight
            GetSVGOpFromBlendMode = "svg:hard-light"
        Case BM_SoftLight
            GetSVGOpFromBlendMode = "svg:soft-light"
        Case BM_Difference
            GetSVGOpFromBlendMode = "svg:difference"
        Case BM_Color
            GetSVGOpFromBlendMode = "svg:color"
        Case BM_Luminosity
            GetSVGOpFromBlendMode = "svg:luminosity"
        Case BM_Hue
            GetSVGOpFromBlendMode = "svg:hue"
        Case BM_Saturation
            GetSVGOpFromBlendMode = "svg:saturation"
        
        'TODO:
        'Case "svg:plus"
        'Case "svg:dst-in"
        'Case "svg:dst-out"
        'Case "svg:src-atop"
        'Case "svg:dst-atop"
        
        Case Else
            InternalError "GetSVGOpFromBlendMode", "unknown blend mode: " & srcBlendMode
            GetSVGOpFromBlendMode = "svg:src-over"
    
    End Select

End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdOpenRaster." & funcName & "() reported an error on file """ & m_Filename & """: " & errDescription
    Else
        Debug.Print "pdOpenRaster." & funcName & "() reported an error on file """ & m_Filename & """: " & errDescription
    End If
End Sub

Private Sub InternalWarning(ByRef funcName As String, ByRef warnText As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdOpenRaster." & funcName & "() warned: " & warnText
    Else
        Debug.Print "pdOpenRaster." & funcName & "() warned: " & warnText
    End If
End Sub
