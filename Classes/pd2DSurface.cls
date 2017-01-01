VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon 2D Rendering Surface class
'Copyright 2012-2017 by Tanner Helland
'Created: 01/September/12
'Last updated: 10/May/16
'Last update: split surface-specific capabilities out from GDI+ and into a safe wrapper
'
'In the pd2D drawing model, "Surfaces" are objects onto which you can paint.  At present, there are two types of surfaces:
'
' 1) Wrappers around existing DCs.  This is help for UI painting, as wrapping a DC is effectively instantaneous
'    (compared to making a copy of the DC's bitmap, painting to that, then flipping back to the DC when finished)
'
' 2) Bitmap (raster) surfaces.  The creation and destruction of such surfaces is handled internally, using a pdDIB object.
'
'Regardless of which surface type you use, you can always access an hDC for external GDI interop purposes.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'The current surface type.  At present, this is immutable; once set, it cannot be migrated
Private m_SurfaceBackend As PD_2D_RENDERING_BACKEND

'The current surface handle.  Will be 0 if the surface is uninitialized.  The meaning of this value varies by backend.
Private m_SurfaceHandle As Long

'GDI+ uses a number of different image handle types.  A "Graphics" handle is like a DC.  It's an abstract value into
' which other, concrete image types can be select.  The above "SurfaceHandle" is actually a "Graphics" handle in disguise.
' Some GDI operations require an "Image" handle.  This handle may a bitmap or metafile.  (And if loaded from file, it may
' be even more abstract, like an EXIF file.)  In PD, it is almost certainly a bitmap object.  This handle is created
' on-demand only as necessary.  Any time the surface handle, above, is cleared, this value must also be cleared.
Private m_GdipImageHandle As Long

'If this surface was created by wrapping an existing hDC, this value will be non-zero.
Private m_SurfaceDC As Long

'The caller can optionally provide the hWnd, if any, associated with the DC above.  This allows us to automagically
' detect size changes.
Private m_SurfaceHWnd As Long
Private m_ClientRect As RECTL

'If the caller does *not* supply an hWnd, they need to pass surface dimensions.  (DCs are dimension-less.)
Private m_SurfaceWidthCaller As Single, m_SurfaceHeightCaller As Single

'If we need to query a wrapped DC for bitmap information (e.g. width/height), we will use this temporary object to retrieve
' data associated with the bitmap currently selected into the DC.  This system is *not* recommended, as bitmaps may not
' correspond in size to the surface (for example, a memory DC without a bitmap selected into it will return a monochrome
' 1x1 bitmap - see https://support.microsoft.com/en-us/kb/139165).  VB objects also tend to have bitmaps that are a
' different size from their parent window, for reasons unknown.
'
'In other words, this is an absolute last resort.  Supply the surface's size at wrap-time, either implicitly via attached
' hWnd or explicitly.
Private m_SurfaceDCBitmap As GDI_Bitmap

'When wrapping a DC, a surface needs to know the size of the object being painted on.  If an hWnd is supplied alongside
' the DC, we'll use that to auto-detect dimensions; otherwise, the caller needs to provide them.  (If the size is
' unknown, we'll use the size of the bitmap currently selected into the DC, but that's *not* reliable - so don't use it
' unless you know what you're doing!)
'
'This value is only used internally, and is set automatically at wrap-time.  It is ignored for internal surfaces,
' because we will always query the DIB directly.
Private m_SurfaceSizeMode As PD_2D_SIZE_DETECTION

'If this surface was created as a new surface, this DIB will serve as the surface backer.  (Similarly, the surface DC will simply
' wrap this DIB's hDC.)
Private m_SurfaceDIB As pdDIB

'This class internally manages what kind of surface it is (e.g. a lightweight wrapper around an existing hDC,
' or a fully-managed DIB surface).  This value is exposed via the GetSurfaceType function, but users should never
' have to mess with this; we only use it internally, in places like the Painter class, to figure out the best
' GDI+ interop method.  (Wrapped hDCs can't be used as 32-bpp surfaces, because alpha data is unreliable.)
Private m_SurfaceType As PD_2D_SurfaceType

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Surface settings that support get/set operations
Private m_SurfaceAntialiasing As PD_2D_Antialiasing
Private m_SurfaceBlendUsingSRGBGamma As Boolean
Private m_SurfacePixelOffset As PD_2D_PixelOffset
Private m_SurfaceRenderingOriginX As Single, m_SurfaceRenderingOriginY As Single
Private m_SurfaceResizeQuality As PD_2D_ResizeQuality

'Get/set individual settings.  Note that these just wrap the generic Get/SetPenProperty functions, below.
Public Function GetSurfaceAntialiasing() As PD_2D_Antialiasing
    GetSurfaceAntialiasing = Me.GetSurfaceProperty(P2_SurfaceAntialiasing)
End Function

Public Function GetSurfaceBlendUsingSRGBGamma() As Boolean
    GetSurfaceBlendUsingSRGBGamma = Me.GetSurfaceProperty(P2_SurfaceBlendUsingSRGBGamma)
End Function

Public Function GetSurfacePixelOffset() As PD_2D_PixelOffset
    GetSurfacePixelOffset = Me.GetSurfaceProperty(P2_SurfacePixelOffset)
End Function

Public Function GetSurfaceRenderingOriginX() As Single
    GetSurfaceRenderingOriginX = Me.GetSurfaceProperty(P2_SurfaceRenderingOriginX)
End Function

Public Function GetSurfaceRenderingOriginY() As Single
    GetSurfaceRenderingOriginY = Me.GetSurfaceProperty(P2_SurfaceRenderingOriginY)
End Function

Public Function GetSurfaceResizeQuality() As PD_2D_ResizeQuality
    GetSurfaceResizeQuality = Me.GetSurfaceProperty(P2_SurfaceResizeQuality)
End Function

Public Sub SetSurfaceAntialiasing(ByVal newSetting As PD_2D_Antialiasing)
    Me.SetSurfaceProperty P2_SurfaceAntialiasing, newSetting
End Sub

Public Sub SetSurfaceBlendUsingSRGBGamma(ByVal newSetting As Boolean)
    Me.SetSurfaceProperty P2_SurfaceBlendUsingSRGBGamma, newSetting
End Sub

Public Sub SetSurfacePixelOffset(ByVal newSetting As PD_2D_PixelOffset)
    Me.SetSurfaceProperty P2_SurfacePixelOffset, newSetting
End Sub

Public Sub SetSurfaceRenderingOriginX(ByVal newSetting As Single)
    Me.SetSurfaceProperty P2_SurfaceRenderingOriginX, newSetting
End Sub

Public Sub SetSurfaceRenderingOriginY(ByVal newSetting As Single)
    Me.SetSurfaceProperty P2_SurfaceRenderingOriginY, newSetting
End Sub

Public Sub SetSurfaceResizeQuality(ByVal newSetting As PD_2D_ResizeQuality)
    Me.SetSurfaceProperty P2_SurfaceResizeQuality, newSetting
End Sub

'Get/set individual settings by enum type
Public Function GetSurfaceProperty(ByVal propID As PD_2D_SURFACE_SETTINGS) As Variant

    Select Case propID
    
        Case P2_SurfaceAntialiasing
            GetSurfaceProperty = m_SurfaceAntialiasing
        
        Case P2_SurfacePixelOffset
            GetSurfaceProperty = m_SurfacePixelOffset
            
        Case P2_SurfaceRenderingOriginX
            GetSurfaceProperty = m_SurfaceRenderingOriginX
        
        Case P2_SurfaceRenderingOriginY
            GetSurfaceProperty = m_SurfaceRenderingOriginY
        
        Case P2_SurfaceBlendUsingSRGBGamma
            GetSurfaceProperty = m_SurfaceBlendUsingSRGBGamma
            
        Case P2_SurfaceResizeQuality
            GetSurfaceProperty = m_SurfaceResizeQuality
    
    End Select

End Function

Public Sub SetSurfaceProperty(ByVal propID As PD_2D_SURFACE_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
    
        Case P2_SurfaceAntialiasing
            m_SurfaceAntialiasing = CLng(newSetting)
        
        Case P2_SurfacePixelOffset
            m_SurfacePixelOffset = CLng(newSetting)
            
        Case P2_SurfaceRenderingOriginX
            m_SurfaceRenderingOriginX = CSng(newSetting)
        
        Case P2_SurfaceRenderingOriginY
            m_SurfaceRenderingOriginY = CSng(newSetting)
            
        Case P2_SurfaceBlendUsingSRGBGamma
            m_SurfaceBlendUsingSRGBGamma = CBool(newSetting)
            
        Case P2_SurfaceResizeQuality
            m_SurfaceResizeQuality = CLng(newSetting)
            
    End Select
    
    RelayPropertyChangeToBackend propID
    
End Sub

'Some property changes exist only at run-time.  Because these properties are not well-defined, they do not support
' serialization (e.g. you cannot save/load them to a string).  Clipping is one such property, because clipping may
' be undefined, a simple rect, or a complex region, and pd2D doesn't implement all those features... yet.
' (GDIPlus can retrieve region data as a byte array, so we could theoretically convert that to Base64 and store it
'  inside an XML tag... maybe someday!)
Public Function GetSurfaceClippingRegion(ByRef dstRegion As pd2DRegion) As Boolean

    If (m_SurfaceHandle <> 0) Then
        
        If (dstRegion Is Nothing) Then Set dstRegion = New pd2DRegion
        
        Select Case m_SurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                Dim tmpHandle As Long
                tmpHandle = GDI_Plus.GDIPlus_GraphicsGetClipRegion(m_SurfaceHandle)
                dstRegion.AssignExternalHandleDirectly tmpHandle
                GetSurfaceClippingRegion = CBool(tmpHandle <> 0)
        End Select
        
    Else
        GetSurfaceClippingRegion = False
    End If
        
End Function

Public Function SetSurfaceClip_FromRectangle(ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean

    If (m_SurfaceHandle <> 0) Then
    
        Select Case m_SurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                SetSurfaceClip_FromRectangle = GDI_Plus.GDIPlus_GraphicsSetClipRect(m_SurfaceHandle, rectLeft, rectTop, rectWidth, rectHeight, useCombineMode)
        End Select
        
    Else
        SetSurfaceClip_FromRectangle = False
    End If
        
End Function

Friend Function SetSurfaceClip_FromRectF(ByRef srcRectF As RECTF, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    With srcRectF
        SetSurfaceClip_FromRectF = SetSurfaceClip_FromRectangle(.Left, .Top, .Width, .Height)
    End With
End Function

Friend Function SetSurfaceClip_FromRectL(ByRef srcRectL As RECTL, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    With srcRectL
        SetSurfaceClip_FromRectL = SetSurfaceClip_FromRectangle(.Left, .Top, .Right - .Left, .Bottom - .Top)
    End With
End Function

Friend Function SetSurfaceClip_FromRegion(ByRef srcRegion As pd2DRegion, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean

    If (m_SurfaceHandle <> 0) Then
    
        Select Case m_SurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                SetSurfaceClip_FromRegion = GDI_Plus.GDIPlus_GraphicsSetClipRegion(m_SurfaceHandle, srcRegion.GetHandle, useCombineMode)
        End Select
        
    Else
        SetSurfaceClip_FromRegion = False
    End If
    
End Function

Public Function SetSurfaceClip_None() As Boolean

    If (m_SurfaceHandle <> 0) Then
    
        Select Case m_SurfaceBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                SetSurfaceClip_None = GDI_Plus.GDIPlus_GraphicsResetClipRegion(m_SurfaceHandle)
        End Select
        
    Else
        SetSurfaceClip_None = False
    End If
    
End Function

'Make an exact copy of an existing surface.
Public Function CloneSurface(ByRef srcSurface As pd2DSurface, Optional ByVal cloneAlphaChannelToo As Boolean = False, Optional ByVal clonePropertiesToo As Boolean = False, Optional ByVal cloneClippingToo As Boolean = False) As Boolean

    If (Not srcSurface Is Nothing) Then
        
        'Note that surface creation from this CreateBlankSurface call will report this surface's state back to the main Drawing2D module.
        ' If subsequent cloning steps fail (very unlikely), then the main surface count will be off by (1).
        CloneSurface = Me.CreateBlankSurface(srcSurface.GetSurfaceWidth, srcSurface.GetSurfaceHeight, IIf(srcSurface.GetSurfaceAlphaSupport, cloneAlphaChannelToo, False))
        
        If CloneSurface Then
        
            'Clone all surface properties
            If clonePropertiesToo Then
                Dim i As PD_2D_SURFACE_SETTINGS
                For i = 0 To [_P2_NumOfSurfaceSettings] - 1
                    Me.SetSurfaceProperty i, srcSurface.GetSurfaceProperty(i)
                Next i
            End If
            
            'Clipping regions are somewhat messy, as their shape can be horrifically complex
            Dim srcRegion As pd2DRegion
            If cloneClippingToo Then
                If srcSurface.GetSurfaceClippingRegion(srcRegion) Then Me.SetSurfaceClip_FromRegion srcRegion, P2_CM_Replace
            End If
            
            'Copy the surface contents into place
            Dim cPainter As pd2DPainter
            Drawing2D.QuickCreatePainter cPainter
            CloneSurface = cPainter.CopySurfaceI(Me, 0, 0, srcSurface)
            
        End If
        
    Else
        InternalError "Failed to clone surface", "Source surface was empty"
    End If

End Function

'Copy the current contents of this surface to some target DC.  (This is a nice shorthand function for flipping backbuffers, for example.)
' I know it's weird, but if you don't supply a destination width and height, the surface's current width/height values will be automatically
' plugged in.
Public Function CopySurfaceToDC(ByVal dstDC As Long, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal dstWidth As Long = 0, Optional ByVal dstHeight As Long = 0, Optional ByVal srcX As Long = 0, Optional ByVal srcY As Long = 0)
    If (dstWidth = 0) Then dstWidth = Me.GetSurfaceWidth
    If (dstHeight = 0) Then dstHeight = Me.GetSurfaceHeight
    CopySurfaceToDC = GDI.BitBltWrapper(dstDC, dstX, dstY, dstWidth, dstHeight, Me.GetSurfaceDC, srcX, srcY)
End Function

'Create a blank in-memory surface.  pd2DSurface will automatically manage the memory for this surface.
Public Function CreateBlankSurface(ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, Optional ByVal surfaceHasAlphaChannel As Boolean = True, Optional ByVal surfaceBackColor As Long = vbWhite, Optional ByVal surfaceOpacity As Single = 100#) As Boolean
    
    If (surfaceWidth <> 0) And (surfaceHeight <> 0) Then
        
        If (m_SurfaceHandle <> 0) Then ReleaseSurface
    
        'Prep a GDI DIB for use as our blank surface
        If (m_SurfaceDIB Is Nothing) Then Set m_SurfaceDIB = New pdDIB
        Dim targetColorDepth As Long
        If surfaceHasAlphaChannel Then targetColorDepth = 32 Else targetColorDepth = 24
        CreateBlankSurface = m_SurfaceDIB.CreateBlank(surfaceWidth, surfaceHeight, targetColorDepth, surfaceBackColor, surfaceOpacity * 2.55)
        
        If CreateBlankSurface Then
        
            m_SurfaceType = P2_ST_Bitmap
            
            'Finally, wrap the appropriate backend around this newly created DC
            Select Case m_SurfaceBackend
                
                Case P2_DefaultBackend, P2_GDIPlusBackend
                    m_SurfaceHandle = GDI_Plus.GetGDIPlusGraphicsFromDC(m_SurfaceDIB.GetDIBDC, IIf(m_SurfaceAntialiasing >= P2_AA_HighQuality, GP_SM_Antialias, GP_SM_None), IIf(m_SurfacePixelOffset >= P2_PO_Half, GP_POM_Half, GP_POM_None))
                    CreateBlankSurface = (m_SurfaceHandle <> 0)
                    
                    'Apply any other stored settings to the newly created surface
                    If CreateBlankSurface Then
                        RelayPropertyChangeToBackend P2_SurfaceBlendUsingSRGBGamma
                        RelayPropertyChangeToBackend P2_SurfaceRenderingOriginX
                        RelayPropertyChangeToBackend P2_SurfaceRenderingOriginY
                        RelayPropertyChangeToBackend P2_SurfaceResizeQuality
                    End If
                    
                    'When debug mode is active, all object creations are reported back to the master Drawing2D module
                    If (CreateBlankSurface And m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, True
                    
            End Select
            
        End If
    Else
        InternalError "Invalid surface width/height", "When creating a blank surface, the width and height must be non-zero."
    End If
    
End Function

'Create a surface from a given filename.  To reduce the chance of unpredictable behavior (including security concerns), please pass
' fully qualified paths, not relative ones.
'
'Supported formats include: BMP, GIF, JPEG, PNG, TIFF, WMF, and EMF.  For best results, load images on Vista or later; XP has some
' limitations beyond our control.  (This is mainly a concern with esoteric formats, like CMYK JPEGs.)
'
'Extra parameters include:
' 1) fixJPEGOrientation: if the source file contains an orientation tag (like photos coming from smartphones), this function can
'                        automatically rotate the image to match.  This is the default behavior, and should generally be disabled
'                        only if the source file contains faulty EXIF data.
Public Function CreateSurfaceFromFile(ByVal srcFilename As String, Optional ByVal fixJPEGOrientation As Boolean = True) As Boolean
    
    If (m_SurfaceHandle <> 0) Then ReleaseSurface
    
    Select Case m_SurfaceBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            
            'To simplify handling of surface(s) created from file, we are only going to lean on GDI+ for the initial image
            ' file parsing.  Once the image is loaded, we will transfer it to an internally managed DIB; this frees us from
            ' obnoxious GDI+ constraints (like being forced to maintain file access for the lifetime of the image object).
            '
            'One downside of this approach is that metafiles will be forcibly rasterized.  This behavior could be "fixed"
            ' in a future release, but honestly, knowing all surfaces are raster-type saves us a lot of trouble.
            
            'Start by using GDI+ to load the file into a GDI+ container.
            Dim hGdipImage As Long, isImageMetafile As Boolean
            hGdipImage = GDI_Plus.GDIPlus_ImageCreateFromFile(srcFilename, isImageMetafile)
            
            If (hGdipImage <> 0) Then
                CreateSurfaceFromFile = InternalGDIPlusLoadImage(hGdipImage, isImageMetafile, fixJPEGOrientation)
            Else
                CreateSurfaceFromFile = False
            End If
            
            'When debug mode is active, all object creations are reported back to the master Drawing2D module
            If (CreateSurfaceFromFile And m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, True
            
    End Select
    
End Function

'GDI+ load helper; this lets us share code between images loaded from file, and loaded from array (memory)
Private Function InternalGDIPlusLoadImage(ByVal hGdipImage As Long, Optional ByVal isImageMetafile As Boolean = False, Optional ByVal fixJPEGOrientation As Boolean = True) As Boolean

    If (hGdipImage <> 0) Then
        
        Dim emfPlusConversionSuccessful As Boolean
        emfPlusConversionSuccessful = False
        
        'Retrieve the exact file format; this may enable additional features
        Dim srcFileFormat As PD_2D_FileFormatImport
        srcFileFormat = GDI_Plus.GDIPlus_ImageGetFileFormat(hGdipImage)
        
        'Retrieve the image's dimensions
        Dim imgWidth As Long, imgHeight As Long
        GDI_Plus.GDIPlus_ImageGetDimensions hGdipImage, imgWidth, imgHeight
        
        Dim tmpPropertyBuffer() As Byte
        
        'If the source image is a metafile, we can enable some bonus features on Vista+
        If isImageMetafile Then
            
            'Ideally, we'd want to keep the metafile in vector format, but in practice, this adds a lot of complexity to the library.
            ' (For example, implementing something like a distort filter on a vector object is impossible, so we'd need to add
            ' auto-rasterization capabilities to a LOT of places, and I don't want to do that right now.)
            '
            'Instead, pd2D rasterizes metafiles at import-time to the size suggested by the file itself.  This can be reevaluated
            ' in the future as time allows.
            
            'Before doing anything else, we need to differentiate between WMF and EMF files.  WMF report their size differently.
            If (srcFileFormat = P2_FFI_WMF) Then
                
                'Retrieve the metafile's horizontal and vertical resolution
                Dim imgHResolution As Single, imgVResolution As Single
                GDI_Plus.GDIPlus_ImageGetResolution hGdipImage, imgHResolution, imgVResolution
                
                'I assume 96 is used because it's the default DPI value in Windows.  I have not tested if different system DPI values
                ' affect the way GDI+ reports metafile size.
                If (imgHResolution <> 0) Then imgWidth = imgWidth * (96# / imgHResolution)
                If (imgVResolution <> 0) Then imgHeight = imgHeight * (96# / imgVResolution)
                
            End If
            
            'If GDI+ v1.1 is available, we can translate EMFs and WMFs into the new GDI+ EMF+ format, which supports antialiasing
            ' and alpha channels (among other cool features).
            If GDI_Plus.IsGDIPlusV11Available Then
                
                'Create a temporary GDI+ graphics object, whose properties will be used to control the render state of the EMF
                Dim tmpSettingsDIB As pdDIB
                Set tmpSettingsDIB = New pdDIB
                tmpSettingsDIB.CreateBlank 8, 8, 32, 0, 0
                
                Dim tmpGraphics As Long
                tmpGraphics = GDI_Plus.GetGDIPlusGraphicsFromDC(tmpSettingsDIB.GetDIBDC, GP_SM_Antialias, GP_POM_Half)
                
                'Attempt to convert the EMF to EMF+ format
                Dim newImageHandle As Long
                If GDI_Plus.GDIPlus_ImageUpgradeMetafile(hGdipImage, tmpGraphics, newImageHandle) Then
                    emfPlusConversionSuccessful = True
                    GDI_Plus.ReleaseGDIPlusImage hGdipImage
                    hGdipImage = newImageHandle
                Else
                    InternalError "Metafile upgrade failed", "Could not convert old-style WMF or EMF to new EMF+ format"
                End If
                
                'Release all the temporary objects we created for the conversion
                If (tmpGraphics <> 0) Then GDI_Plus.ReleaseGDIPlusGraphics tmpGraphics
                Set tmpSettingsDIB = Nothing
                
            End If
            
        'If the source image is a raster file, we can also enable some bonus features
        Else
        
            'Automatically apply any orientation metadata, if present
            If GDI_Plus.GDIPlus_ImageGetProperty(hGdipImage, GP_PT_Orientation, tmpPropertyBuffer) Then
            
                'The first 16 bytes of a GDI+ property are ignored.  For the orientation property, we need the MSB of the
                ' 2-byte trailer of the returned array.
                Select Case tmpPropertyBuffer(UBound(tmpPropertyBuffer) - 1)
                
                    'For details on what the different orientation values mean, visit https://msdn.microsoft.com/en-us/library/ms534416(v=vs.85).aspx
                    Case 1
                        'Standard orientation; no modifications needed
                    Case 2
                        GDI_Plus.GDIPlus_ImageRotateFlip hGdipImage, GP_RF_NoneFlipX
                    Case 3
                        GDI_Plus.GDIPlus_ImageRotateFlip hGdipImage, GP_RF_180FlipNone
                    Case 4
                        GDI_Plus.GDIPlus_ImageRotateFlip hGdipImage, GP_RF_NoneFlipY
                    Case 5
                        GDI_Plus.GDIPlus_ImageRotateFlip hGdipImage, GP_RF_270FlipY
                    Case 6
                        GDI_Plus.GDIPlus_ImageRotateFlip hGdipImage, GP_RF_90FlipNone
                    Case 7
                        GDI_Plus.GDIPlus_ImageRotateFlip hGdipImage, GP_RF_90FlipY
                    Case 8
                        GDI_Plus.GDIPlus_ImageRotateFlip hGdipImage, GP_RF_270FlipNone
                End Select
            
            End If
        
        End If
        
        'Check for the presence of an alpha channel
        Dim imgPixelFormat As GP_PixelFormat, hasAlpha As Boolean
        imgPixelFormat = GDI_Plus.GDIPlus_ImageGetPixelFormat(hGdipImage)
        hasAlpha = CBool((imgPixelFormat And GP_PF_Alpha) <> 0) Or CBool((imgPixelFormat And GP_PF_PreMultAlpha) <> 0)
        
        'Metafiles require special handling on Vista and earlier
        Set m_SurfaceDIB = New pdDIB
        
        If isImageMetafile Then
            If emfPlusConversionSuccessful Or hasAlpha Then
                m_SurfaceDIB.CreateBlank imgWidth, imgHeight, 32, 0, 0
            Else
                m_SurfaceDIB.CreateBlank imgWidth, imgHeight, 24
            End If
        Else
            m_SurfaceDIB.CreateBlank imgWidth, imgHeight, 32, 0, 0
        End If
        
        'It's finally time to copy the GDI+ image data into our internal DIB container.  Note that 32-bpp images are
        ' handled differently from 24-bpp images, due to some quirks with GDI+.
        If hasAlpha Then
            
            'Make sure the image is in 32-bpp premultiplied ARGB format
            If (imgPixelFormat <> GP_PF_32bppPARGB) And (Not isImageMetafile) Then GDI_Plus.GDIPlus_ImageForcePremultipliedAlpha hGdipImage, imgWidth, imgHeight
            m_SurfaceDIB.SetInitialAlphaPremultiplicationState True
            
            'We are now going to copy the image's data directly into our destination DIB by using LockBits.  Very fast, and not much code!
            
            'Start by preparing a GP_BitmapData variable, which describes how we want GDI+ to lock the data.  (In our case, we instruct
            ' it to copy the data into a buffer of our own creation.)
            Dim copyBitmapData As GP_BitmapData
            With copyBitmapData
                .BD_Width = imgWidth
                .BD_Height = imgHeight
                .BD_PixelFormat = GP_PF_32bppPARGB
                .BD_Stride = m_SurfaceDIB.GetDIBStride
                .BD_Scan0 = m_SurfaceDIB.GetDIBPointer
            End With
            
            'Next, prepare a clipping rect
            Dim tmpRect As RECTL
            With tmpRect
                .Left = 0
                .Top = 0
                .Right = imgWidth
                .Bottom = imgHeight
            End With
            
            'Use LockBits to perform the copy for us.
            InternalGDIPlusLoadImage = GDI_Plus.GDIPlus_ImageLockBits(hGdipImage, tmpRect, copyBitmapData, GP_BLM_UserInputBuf Or GP_BLM_Write Or GP_BLM_Read, GP_PF_32bppPARGB)
            GDI_Plus.GDIPlus_ImageUnlockBits hGdipImage, copyBitmapData
            
        '24-bpp images can be rendered directly onto our internal DIB
        Else
            Dim hGraphics As Long
            hGraphics = GDI_Plus.GetGDIPlusGraphicsFromDC(m_SurfaceDIB.GetDIBDC, GP_SM_Antialias, GP_POM_None)
            InternalGDIPlusLoadImage = GDI_Plus.GDIPlus_DrawImageRectI(hGraphics, hGdipImage, 0, 0, imgWidth, imgHeight)
            GDI_Plus.ReleaseGDIPlusGraphics hGraphics
        End If
        
        'Release any remaining GDI+ handles and exit
        GDI_Plus.ReleaseGDIPlusImage hGdipImage
        
        'If successful, relay any remaining surface properties to the GDI+ container associated with this surface
        If InternalGDIPlusLoadImage Then
            m_SurfaceType = P2_ST_Bitmap
            m_SurfaceHandle = GDI_Plus.GetGDIPlusGraphicsFromDC(m_SurfaceDIB.GetDIBDC, IIf(m_SurfaceAntialiasing >= P2_AA_HighQuality, GP_SM_Antialias, GP_SM_None), IIf(m_SurfacePixelOffset >= P2_PO_Half, GP_POM_Half, GP_POM_None))
            InternalGDIPlusLoadImage = (m_SurfaceHandle <> 0)
            
            'Apply any other stored settings to the newly created surface
            If InternalGDIPlusLoadImage Then
                RelayPropertyChangeToBackend P2_SurfaceBlendUsingSRGBGamma
                RelayPropertyChangeToBackend P2_SurfaceRenderingOriginX
                RelayPropertyChangeToBackend P2_SurfaceRenderingOriginY
                RelayPropertyChangeToBackend P2_SurfaceResizeQuality
            End If
            
        End If
        
    Else
        InternalGDIPlusLoadImage = False
    End If
            
End Function

'If you already have an image file loaded into a VB byte array, you can use this function to load it.  Note that the source array
' must contain a valid image file - not just a bare array of RGB values.
'
'Note also that you can free the array after calling this function; it does not need to be maintained post-load.
Public Function CreateSurfaceFromArray(ByRef srcArray() As Byte, Optional ByVal fixJPEGOrientation As Boolean = True) As Boolean
    
    If (m_SurfaceHandle <> 0) Then ReleaseSurface
    
    Select Case m_SurfaceBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            
            'Start by using GDI+ to load the array into a GDI+ container.
            Dim hGdipImage As Long, isImageMetafile As Boolean
            hGdipImage = GDI_Plus.GDIPlus_ImageCreateFromArray(srcArray, isImageMetafile)
            
            If (hGdipImage <> 0) Then
                CreateSurfaceFromArray = InternalGDIPlusLoadImage(hGdipImage, isImageMetafile, fixJPEGOrientation)
            Else
                CreateSurfaceFromArray = False
            End If
            
            'When debug mode is active, all object creations are reported back to the master Drawing2D module
            If (CreateSurfaceFromArray And m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, True
            
    End Select
    
End Function

'Wrap this surface around some existing hDC.  The DC can be any type (display or memory; printer may also work but I haven't
' tested them *at all*).
'
'If this DC is associated with a window or a VB object, you should also pass the object's hWnd.  hWnds are more reliable for
' retrieving surface dimensions, because we can retrieve the surface's size on-demand.  (For example, if you wrap a picture box
' or form, then resize the picture box or form, this class will auto-magically update itself to match.)
'
'If you *don't* provide an hWnd, you need to provide the surface's width and height.  This allows pd2D to auto-calculate things
' like fill boundaries and implicit draw operations (draw operations where the caller isn't required to pass width and height).
' If you don't provide width and height, pd2D will try to guess dimensions from the active object inside the DC, but this is
' subject to a bunch of problems and is unlikely to provide desired results.
Public Function WrapSurfaceAroundDC(ByVal srcDC As Long, Optional ByVal srcHwnd As Long = 0, Optional ByVal surfaceWidth As Single = 0#, Optional ByVal surfaceHeight As Single = 0#) As Boolean
    
    If (srcDC <> 0) Then
    
        If (srcDC = m_SurfaceDC) Then
            WrapSurfaceAroundDC = True
        Else
            
            If (m_SurfaceHandle <> 0) Then ReleaseSurface
            
            'Call the backend-specific surface creation function
            Select Case m_SurfaceBackend
                
                Case P2_DefaultBackend, P2_GDIPlusBackend
                    
                    m_SurfaceHandle = GDI_Plus.GetGDIPlusGraphicsFromDC(srcDC, IIf(m_SurfaceAntialiasing >= P2_AA_HighQuality, GP_SM_Antialias, GP_SM_None), IIf(m_SurfacePixelOffset >= P2_PO_Half, GP_POM_Half, GP_POM_None))
                    WrapSurfaceAroundDC = (m_SurfaceHandle <> 0)
                    
                    If WrapSurfaceAroundDC Then
                        m_SurfaceDC = srcDC
                        m_SurfaceType = P2_ST_WrapperOnly
                        
                        'Apply any other stored settings to the newly created surface
                        RelayPropertyChangeToBackend P2_SurfaceBlendUsingSRGBGamma
                        RelayPropertyChangeToBackend P2_SurfaceRenderingOriginX
                        RelayPropertyChangeToBackend P2_SurfaceRenderingOriginY
                        RelayPropertyChangeToBackend P2_SurfaceResizeQuality
                        
                        'Determine how we're gonna get size data from this DC
                        If (srcHwnd <> 0) Then
                            m_SurfaceHWnd = srcHwnd
                            m_SurfaceSizeMode = P2_SizeFromHWnd
                        Else
                            If (surfaceWidth <> 0) And (surfaceHeight <> 0) Then
                                m_SurfaceWidthCaller = surfaceWidth
                                m_SurfaceHeightCaller = surfaceHeight
                                m_SurfaceSizeMode = P2_SizeFromCaller
                            Else
                                m_SurfaceSizeMode = P2_SizeUnknown
                            End If
                        End If
                        
                    Else
                        InternalError "GDI+ DC wrapper failed", "GDI+ failed to wrap the hDC passed to pd2DSurface.WrapSurfaceAroundDC"
                    End If
                    
            End Select
            
            'When debug mode is active, all object creations are reported back to the master Drawing2D module
            If (WrapSurfaceAroundDC And m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, True
            
        End If
    
    Else
        InternalError "hDC is null", "pd2DSurface.WrapSurfaceAroundDC doesn't work on null DCs."
    End If
    
End Function

'Save the current surface to file.  To reduce the chance of unpredictable behavior (including security concerns), please pass
' fully qualified paths, not relative ones.
'
'Supported formats include: BMP, GIF, JPEG, PNG, and TIFF.  (You may be able to save metafiles, but it's completely untested.)
'
'Note that wrapped VB objects (like picture boxes) may demonstrate unstable alpha channel behavior.  For best results, do all rendering
' to in-memory surfaces, and use those surfaces for saving.  If the alpha channel of the source surface is ever in doubt, I recommend
' manually converting the underlying DIB to 24-bpp prior to saving.
'
'Extra parameters include:
' 1) dstFileFormat: file format of the saved image.  If no format is provided, lossless PNG will be used.
' 2) outputColorDepth: integer value that must be a power of two (1, 2, 4, 8, 16, 24, 32).  Not all color depths are supported for
'     all output formats.  (For example, GIFs are always 8-bpp.)  If no output depth is provided, standard GDI+ color depth rules
'     will be obeyed.
' 3) jpegQuality: integer on the range [0, 100] describing the output JPEG's quality.  100 = nearly lossless, 0 = extremely lossy
Public Function SaveSurfaceToFile(ByVal dstFilename As String, Optional ByVal dstFileFormat As PD_2D_FileFormatExport = P2_FFE_PNG, Optional ByVal jpegQuality As Long = 90) As Boolean

    Select Case m_SurfaceBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            
            Dim useTmpSurfaceCopy As Boolean: useTmpSurfaceCopy = False
            Dim tmpSurfaceCopy As pd2DSurface
            
            'If the output color depth is 24-bpp or less, forcibly composite our surface against white prior to saving.
            If ((dstFileFormat = P2_FFE_GIF) Or (dstFileFormat = P2_FFE_JPEG)) And (Me.GetSurfaceAlphaSupport) Then
                Drawing2D.QuickCreateBlankSurface tmpSurfaceCopy, Me.GetSurfaceWidth, Me.GetSurfaceHeight, False, True, vbWhite, 100
                Dim cPainter As pd2DPainter
                Drawing2D.QuickCreatePainter cPainter
                cPainter.DrawSurfaceCroppedI tmpSurfaceCopy, 0, 0, Me.GetSurfaceWidth, Me.GetSurfaceHeight, Me, 0, 0
                useTmpSurfaceCopy = True
            End If
            
            'Retrieve the relevant surface's GDI+ bitmap handle
            Dim gdipBitmap As Long
            If useTmpSurfaceCopy Then gdipBitmap = tmpSurfaceCopy.GetGdipImageHandle Else gdipBitmap = Me.GetGdipImageHandle
            
            'GDI+ handles the rest...
            If (gdipBitmap <> 0) Then
                SaveSurfaceToFile = GDI_Plus.GDIPlus_ImageSaveToFile(gdipBitmap, dstFilename, dstFileFormat, jpegQuality)
            Else
                InternalError "Could not create GDI+ Bitmap Object", "SaveSurfaceToFile() failed because GDI+ failed to create a matching Bitmap object prior to saving."
            End If
        
    End Select
    
End Function

'Identical function to SaveSurfaceToFile, above, except the target is a VB array.  Note that this function uses streams, so the hard drive
' isn't touched.
Public Function SaveSurfaceToArray(ByRef dstArray() As Byte, Optional ByVal dstFileFormat As PD_2D_FileFormatExport = P2_FFE_PNG, Optional ByVal jpegQuality As Long = 90) As Boolean

    Select Case m_SurfaceBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            
            Dim useTmpSurfaceCopy As Boolean: useTmpSurfaceCopy = False
            Dim tmpSurfaceCopy As pd2DSurface
            
            'If the output color depth is 24-bpp or less, forcibly composite our surface against white prior to saving.
            If ((dstFileFormat = P2_FFE_GIF) Or (dstFileFormat = P2_FFE_JPEG)) And (Me.GetSurfaceAlphaSupport) Then
                Drawing2D.QuickCreateBlankSurface tmpSurfaceCopy, Me.GetSurfaceWidth, Me.GetSurfaceHeight, False, True, vbWhite, 100
                Dim cPainter As pd2DPainter
                Drawing2D.QuickCreatePainter cPainter
                cPainter.DrawSurfaceCroppedI tmpSurfaceCopy, 0, 0, Me.GetSurfaceWidth, Me.GetSurfaceHeight, Me, 0, 0
                useTmpSurfaceCopy = True
            End If
            
            'Retrieve the relevant surface's GDI+ bitmap handle
            Dim gdipBitmap As Long
            If useTmpSurfaceCopy Then gdipBitmap = tmpSurfaceCopy.GetGdipImageHandle Else gdipBitmap = Me.GetGdipImageHandle
            
            'GDI+ handles the rest...
            If (gdipBitmap <> 0) Then
                SaveSurfaceToArray = GDI_Plus.GDIPlus_ImageSaveToArray(gdipBitmap, dstArray, dstFileFormat, jpegQuality)
            Else
                InternalError "Could not create GDI+ Bitmap Object", "SaveSurfaceToArray() failed because GDI+ failed to create a matching Bitmap object prior to saving."
            End If
        
    End Select
    
End Function

Public Function GetHandle() As Long
    GetHandle = m_SurfaceHandle
End Function

'Retrieving an Image-type handle is an expensive operation for graphics objects wrapped around a DC.  Whenever possible,
' restrict usage of this function to internally-managed DIBs (which are very fast by comparison).
Friend Function GetGdipImageHandle() As Long
    
    If (m_GdipImageHandle = 0) Then
        
        'If this surface is just a wrapper around a DC, we need to create a temporary copy of it, and wrap *that* instead.
        ' AFAIK, this is the fastest, most reliable way to generate bitmap data from an arbitrary DC (which may not contain
        ' bitmap data at all).
        
        'Note that alpha data *will* be ignored, by design.  The only way to fix this would be to ask the caller for
        ' permission to use alpha bytes, which may not be valid as many (most?) DCs will have DDBs selected into them,
        ' meaning alpha bytes are device-dependent and not necessarily relevant or even correct.
        If (m_SurfaceType = P2_ST_WrapperOnly) Then
            
            'Allocating memory for new DIBs is expensive, so whenever possible, skip these steps
            If (m_SurfaceDIB Is Nothing) Then Set m_SurfaceDIB = New pdDIB
            If (m_SurfaceDIB.GetDIBWidth <> Me.GetSurfaceWidth) Or (m_SurfaceDIB.GetDIBHeight <> Me.GetSurfaceHeight) Then m_SurfaceDIB.CreateBlank Me.GetSurfaceWidth, Me.GetSurfaceHeight, 24, 0
            
            'BitBlt lets us quickly copy the source DC's contents, regardless of the source DC's selected object
            GDI.BitBltWrapper m_SurfaceDIB.GetDIBDC, 0, 0, m_SurfaceDIB.GetDIBWidth, m_SurfaceDIB.GetDIBHeight, m_SurfaceDC, 0, 0, vbSrcCopy
            
        End If
        
        If (Not GDI_Plus.GetGdipBitmapHandleFromDIB(m_GdipImageHandle, m_SurfaceDIB)) Then InternalError "GetGdipImageHandle failed", "Could not create a GDI+ Image (bitmap) object from wrapped hDC."
        
    'If we already have a handle, but we are in "wrapper" mode, we need to make sure our duplicate DIB copy is still up-to-date.
    Else
    
        If (m_SurfaceType = P2_ST_WrapperOnly) Then
        
            If (m_SurfaceDIB.GetDIBWidth <> Me.GetSurfaceWidth) Or (m_SurfaceDIB.GetDIBHeight <> Me.GetSurfaceHeight) Then
                ReleaseExtraInternalObjects
                m_SurfaceDIB.CreateBlank Me.GetSurfaceWidth, Me.GetSurfaceHeight, 24, 0
            End If
            
            GDI.BitBltWrapper m_SurfaceDIB.GetDIBDC, 0, 0, m_SurfaceDIB.GetDIBWidth, m_SurfaceDIB.GetDIBHeight, m_SurfaceDC, 0, 0, vbSrcCopy
            
            If (m_GdipImageHandle = 0) Then
                If Not GDI_Plus.GetGdipBitmapHandleFromDIB(m_GdipImageHandle, m_SurfaceDIB) Then InternalError "GetGdipImageHandle failed", "Could not create a GDI+ Image (bitmap) object from wrapped hDC."
            End If
            
        End If
        
    End If
    
    GetGdipImageHandle = m_GdipImageHandle
    
End Function

Friend Function GetSurfaceBackend() As PD_2D_RENDERING_BACKEND
    GetSurfaceBackend = m_SurfaceBackend
End Function

'Users never need to call this function; it is only used to accelerate various drawing options
Friend Function GetSurfaceType() As PD_2D_SurfaceType
    GetSurfaceType = m_SurfaceType
End Function

'I'm still debating whether it's wise to expose this function externally.  GDI+ handles DCs in weird ways;
' as long as it is only ever wrapped around *existing* DCs, you can usually intermix GDI+ and GDI calls
' without running into problems.  However, if you create a native GDI+ surface then use GdipGetDC() to
' retrieve a DC for it, you *must* call GdipReleaseDC when you're done (which is very unintuitive, given the
' way DCs usually work for VB programmers).
'
'As such, my current inclination is to always keep a GDI-backed copy of any surfaces created by this class.
' This simplifies situations where the wants to intermix bare GDI calls (like BitBlt) and pd2D drawing calls,
' and it also circumvents some obnoxious GDI+ limitations (like being forced to keep a source file around for
' the lifetime of an image object created from that file).
'
'Anyway, I mention this here as an FYI, but tl;dr: for now, this function is safe to use, and you don't have
' to free the DC after an operation completes.
Public Function GetSurfaceDC() As Long
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
        GetSurfaceDC = m_SurfaceDC
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceDC = m_SurfaceDIB.GetDIBDC Else GetSurfaceDC = 0
    Else
        GetSurfaceDC = 0
    End If
End Function

Public Function GetSurfaceAlphaSupport() As Boolean
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
        m_SurfaceDCBitmap = GDI.GetBitmapHeaderFromDC(m_SurfaceDC)
        GetSurfaceAlphaSupport = CBool(m_SurfaceDCBitmap.BitsPerPixel = 32)
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceAlphaSupport = CBool(m_SurfaceDIB.GetDIBColorDepth = 32)
    Else
        GetSurfaceAlphaSupport = False
    End If
End Function

Public Function GetSurfaceHeight() As Long
    If (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceHeight = m_SurfaceDIB.GetDIBHeight
    ElseIf (m_SurfaceType = P2_ST_WrapperOnly) Then
        Select Case m_SurfaceSizeMode
            Case P2_SizeFromCaller
                GetSurfaceHeight = m_SurfaceHeightCaller
            Case P2_SizeFromHWnd
                If GDI.GetClientRectWrapper(m_SurfaceHWnd, VarPtr(m_ClientRect)) Then GetSurfaceHeight = m_ClientRect.Bottom
            Case P2_SizeUnknown
                m_SurfaceDCBitmap = GDI.GetBitmapHeaderFromDC(m_SurfaceDC)
                GetSurfaceHeight = m_SurfaceDCBitmap.Height
        End Select
    Else
        GetSurfaceHeight = 0
    End If
End Function

'Stride calculations are *not* guaranteed to be accurate for wrapped DCs, because the underlying surface format is device-dependent.
' 32-bpp is assumed but this is not a safe assumption pre-Vista.  (It shouldn't matter because this function is not used at present,
' but FYI for future use.)
Public Function GetSurfaceStride() As Long
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
         Select Case m_SurfaceSizeMode
            Case P2_SizeFromCaller
                GetSurfaceStride = m_SurfaceWidthCaller * 4
            Case P2_SizeFromHWnd
                GetSurfaceStride = Me.GetSurfaceWidth * 4
            Case P2_SizeUnknown
                m_SurfaceDCBitmap = GDI.GetBitmapHeaderFromDC(m_SurfaceDC)
                GetSurfaceStride = m_SurfaceDCBitmap.WidthBytes
        End Select
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceStride = m_SurfaceDIB.GetDIBStride
    Else
        GetSurfaceStride = 0
    End If
End Function

Public Function GetSurfaceWidth() As Long
    If (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then GetSurfaceWidth = m_SurfaceDIB.GetDIBWidth
    ElseIf (m_SurfaceType = P2_ST_WrapperOnly) Then
        Select Case m_SurfaceSizeMode
            Case P2_SizeFromCaller
                GetSurfaceWidth = m_SurfaceWidthCaller
            Case P2_SizeFromHWnd
                If GDI.GetClientRectWrapper(m_SurfaceHWnd, VarPtr(m_ClientRect)) Then GetSurfaceWidth = m_ClientRect.Right
            Case P2_SizeUnknown
                m_SurfaceDCBitmap = GDI.GetBitmapHeaderFromDC(m_SurfaceDC)
                GetSurfaceWidth = m_SurfaceDCBitmap.Width
        End Select
    Else
        GetSurfaceWidth = 0
    End If
End Function

'If this is an in-memory surface, you can directly access its underlying DIB via this function.  Note that the underlying DIB
' will be "Nothing" if this surface is not an in-memory surface (e.g. if it's just a wrapper around a DC).
Public Property Get UnderlyingDIB() As pdDIB
    Set UnderlyingDIB = m_SurfaceDIB
End Property

Public Function HasSurface() As Boolean
    HasSurface = CBool(m_SurfaceHandle <> 0)
End Function

'You only need to manually initialize a surface backend if you plan on using a non-default backend.  (Surfaces with the
' default backend are initialized automatically.)
Friend Function InitializeSurfaceBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    m_SurfaceBackend = targetBackend
    InitializeSurfaceBackend = True
End Function

'Erase the current surface's contents.  Roughly equivalent to .Cls in VB, except that this function accepts a
' color and/or alpha to use when erasing.
'
'Note that wrapped GDI surfaces will *not* touch alpha, by design, to avoid screwing up the target DC's behavior.
Public Sub EraseSurfaceContents(Optional ByVal newColor As Long = vbWhite, Optional ByVal newAlpha As Single = 0#)
    If (m_SurfaceType = P2_ST_WrapperOnly) Then
        GDI.FillRectToDC m_SurfaceDC, 0, 0, Me.GetSurfaceWidth + 1, Me.GetSurfaceHeight + 1, newColor
    ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
        If (Not m_SurfaceDIB Is Nothing) Then m_SurfaceDIB.FillWithColor newColor, newAlpha
    End If
End Sub

Public Function ReleaseSurface() As Boolean
    
    'If any extra, specialized handles have been created, free them first.  (Some of these handles may wrap data
    ' we are about to free, so releasing them first prevents double-freeing crashes.)
    ReleaseExtraInternalObjects
    
    If (m_SurfaceHandle <> 0) Then
        
        'Call the backend-specific release function
        Select Case m_SurfaceBackend
        
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ReleaseSurface = GDI_Plus.ReleaseGDIPlusGraphics(m_SurfaceHandle)
        
        End Select
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseSurface Then
            m_SurfaceHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifySurfaceCountChange m_SurfaceBackend, False
            
            'Also reset any surface-type-specific values
            If (m_SurfaceType = P2_ST_WrapperOnly) Then
                m_SurfaceDC = 0
                
                'Wrapped surfaces may use a temporary DIB for certain operations, so free that temp DIB if it exists
                If (Not m_SurfaceDIB Is Nothing) Then Set m_SurfaceDIB = Nothing
                
            ElseIf (m_SurfaceType = P2_ST_Bitmap) Then
                Set m_SurfaceDIB = Nothing
            End If
            
            m_SurfaceType = P2_ST_Uninitialized
            
        End If
    
    Else
        ReleaseSurface = True
    End If
    
    'Releasing a handle also releases any underlying trackers (hDCs, pointers, etc)
    m_SurfaceDC = 0

End Function

Private Sub ReleaseExtraInternalObjects()

    'GDI+ "Image" handles may be created on-demand, if required by a given paint function
    If (m_GdipImageHandle <> 0) Then
        GDI_Plus.ReleaseGDIPlusImage m_GdipImageHandle
        m_GdipImageHandle = 0
    End If
    
End Sub

Private Sub RelayPropertyChangeToBackend(ByVal propID As PD_2D_SURFACE_SETTINGS)
    
    'Property changes only need to be relayed if we've already created a handle
    If (m_SurfaceHandle <> 0) Then
    
        'Relays obviously vary by the active backend!
        Select Case m_SurfaceBackend
            
            Case P2_DefaultBackend, P2_GDIPlusBackend
                
                Select Case propID
                    Case P2_SurfaceAntialiasing
                        If (m_SurfaceAntialiasing >= P2_AA_HighQuality) Then
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_SM_Antialias
                        Else
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_SM_None
                        End If
                        
                    Case P2_SurfacePixelOffset
                        If (m_SurfacePixelOffset >= P2_PO_Half) Then
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_POM_Half
                        Else
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, GP_POM_None
                        End If
                    
                    Case P2_SurfaceRenderingOriginX
                        GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, m_SurfaceRenderingOriginX
                        
                    Case P2_SurfaceRenderingOriginY
                        GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, propID, m_SurfaceRenderingOriginY
                        
                    Case P2_SurfaceBlendUsingSRGBGamma
                        If m_SurfaceBlendUsingSRGBGamma Then
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceBlendUsingSRGBGamma, GP_CQ_GammaCorrected
                        Else
                            GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceBlendUsingSRGBGamma, GP_CQ_AssumeLinear
                        End If
                        
                    Case P2_SurfaceResizeQuality
                        Select Case m_SurfaceResizeQuality
                            Case P2_RQ_Fast
                                GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceResizeQuality, GP_IM_NearestNeighbor
                            Case P2_RQ_Bilinear
                                GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceResizeQuality, GP_IM_HighQualityBilinear
                            Case P2_RQ_Bicubic
                                GDI_Plus.SetGDIPlusGraphicsProperty m_SurfaceHandle, P2_SurfaceResizeQuality, GP_IM_HighQualityBicubic
                        End Select
                        
                End Select
                
            'In the future, other backends can be handled here...
            Case Else
                
        End Select
    
    End If
    
End Sub

Public Sub ResetAllProperties()
    Me.SetSurfaceProperty P2_SurfaceAntialiasing, P2_AA_None
    Me.SetSurfaceProperty P2_SurfacePixelOffset, P2_PO_Normal
    Me.SetSurfaceProperty P2_SurfaceRenderingOriginX, 0#
    Me.SetSurfaceProperty P2_SurfaceRenderingOriginY, 0#
    Me.SetSurfaceProperty P2_SurfaceBlendUsingSRGBGamma, False
    Me.SetSurfaceProperty P2_SurfaceResizeQuality, P2_RQ_Fast
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Initialize()
    m_SurfaceType = P2_ST_Uninitialized
    Me.ResetAllProperties
    m_DebugMode = Drawing2D.GetLibraryDebugMode()
End Sub

Private Sub Class_Terminate()
    Me.ReleaseSurface
End Sub

'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DSurface"
End Sub

