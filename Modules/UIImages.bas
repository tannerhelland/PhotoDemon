Attribute VB_Name = "UIImages"
'***************************************************************************
'PhotoDemon Central UI image cache
'Copyright 2018-2026 by Tanner Helland
'Created: 13/July/18
'Last updated: 01/June/21
'Last update: rewrite against pdSpriteSheet, to reuse all the great optimization work I did there
'             as part of animation support (e.g. built-in compression support).
'
'PhotoDemon uses a *lot* of UI images.  The amount of GDI objects required for these surfaces is
' substantial, and we can greatly reduce requirements by using something akin to "sprite sheets",
' e.g. shared image storage for images with similar dimensions.
'
'At present, this module accepts images of any size, but it only provides a benefit when images are
' the *same* size - this allows it to automatically "coalesce" images into shared sheets, which callers
' can then access by index (rather than managing their own pdDIB instance).
'
'This module really only makes sense for images that are kept alive for the duration of the program.
' One-off images (e.g. temp images) should *not* be used, as it is non-trivial to release shared images
' in a performance-friendly manner.
'
'At present, PD limits usage of this cache to pdButtonToolbox images.  (They are the perfect use-case
' for shared caching.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To force a sprite handle to an unsupported value, use this constant.  (This is useful for
' initializing sprite sheet values, to determine if a given sprite has been loaded yet.)
Public Const UI_SPRITE_UNDEFINED As Long = &HFFFFFFFF

'This module no longer uses a custom spritesheet implementation; instead, it wraps pdSpriteSheet
' (which was heavily optimized as part of work on animated image support).
' Note that one great side-effect of this is that this module now supports images of varying sizes.
' (Individual sheets are automatically created for each added size, as necessary.)
Private m_Sprites As pdSpriteSheet

'During a given session, we periodically compress latent UI images to memory buffers, which frees
' up previous resources.  Because compression requires a temporary target buffer of some safe
' minimally acceptable size (for uncompressible data), we can greatly reduce memory thrashing by
' reusing a single temp buffer for this.
Private m_TempCompressBuffer() As Byte, m_CompressBufferSize As Long

'Add an image to the cache.  The returned Long is the handle into the cache; you MUST remember it,
' as it's the only way to access the image again!
'
'When adding images to the cache, you must also pass a unique image name.  This ensures that cache
' entries are never duplicated, which is important as some images are reused throughout PD (for example,
' if every usage instance attempted to add that image to the cache, we would waste a lot of time and
' memory).  Note that the name is only required when *adding* images, so that we can perform a
' duplication check.  Once added, an image's handle is all that's required to retrieve it.
'
'RETURNS: non-zero value if successful; zero if the function fails.
Public Function AddImage(ByRef srcDIB As pdDIB, ByRef uniqueImageName As String) As Long
    
    'Initialize sprite manager on first-use
    If (m_Sprites Is Nothing) Then Set m_Sprites = New pdSpriteSheet
    AddImage = m_Sprites.AddImage(srcDIB, uniqueImageName)
    
End Function

'Free the shared compression buffer for UI images.  This carries a ripple effect for PD's internal
' suspend-to-memory operations, so do this only if the memory savings are large.
Public Sub FreeSharedCompressBuffer()
    If (m_CompressBufferSize <> 0) Then
        PDDebug.LogAction "Freeing shared memory buffer (size " & Files.GetFormattedFileSize(m_CompressBufferSize) & ")"
        Erase m_TempCompressBuffer
        m_CompressBufferSize = 0
    End If
End Sub

'Return a standalone DIB of a given sprite.  Do *not* use this more than absolutely necessary,
' as it is expensive to initialize sprites (and it sort of defeats the purpose of using a
' sprite sheet in the first place!)
Public Function GetCopyOfSprite(ByVal srcImgID As Long, Optional ByRef dstSpriteName As String = vbNullString) As pdDIB
    
    If (Not m_Sprites Is Nothing) Then
        
        Dim imgWidth As Long, imgHeight As Long
        imgWidth = m_Sprites.GetSpriteWidth(srcImgID)
        imgHeight = m_Sprites.GetSpriteHeight(srcImgID)
        
        If (imgWidth <> 0) And (imgHeight <> 0) Then
            Set GetCopyOfSprite = New pdDIB
            GetCopyOfSprite.CreateBlank imgWidth, imgHeight, 32, 0, 0
            GetCopyOfSprite.SetInitialAlphaPremultiplicationState True
            m_Sprites.CopyCachedImage GetCopyOfSprite.GetDIBDC, 0, 0, srcImgID
            dstSpriteName = m_Sprites.GetSpriteName(srcImgID)
        End If
        
    End If
    
End Function

'Get access to the shared compression buffer for UI images.  Do *not* use this buffer for other purposes,
' as it may grow excessively large (and it's not easily freed).
Public Function GetSharedCompressBuffer(ByRef dstBufferSize As Long, ByVal requiredSize As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Lz4, Optional ByVal cmpLevel As Long = -1) As Long

    'Figure out worst-case scenario size for this format, then resize the buffer accordingly
    dstBufferSize = Compression.GetWorstCaseSize(requiredSize, cmpFormat, cmpLevel)
    
    If (dstBufferSize > m_CompressBufferSize) Then
        ReDim m_TempCompressBuffer(0 To dstBufferSize - 1) As Byte
        m_CompressBufferSize = dstBufferSize
    Else
        dstBufferSize = m_CompressBufferSize
    End If
    
    GetSharedCompressBuffer = VarPtr(m_TempCompressBuffer(0))

End Function

Public Sub MinimizeCacheMemory()
    If (Not m_Sprites Is Nothing) Then m_Sprites.MinimizeMemory cf_Lz4, False
End Sub

Public Function PaintCachedImage(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcImgID As Long) As Boolean
    If (Not m_Sprites Is Nothing) Then PaintCachedImage = m_Sprites.PaintCachedImage(dstDC, dstX, dstY, srcImgID)
    If (Not PaintCachedImage) Then PDDebug.LogAction "WARNING!  UIImages.PaintCachedImage failed to paint image " & srcImgID
End Function

'After using a sprite, you can call this function to suspend the source sprite to a
' compressed memory stream.  This can help keep memory usage low during an extended
' editing session.
Public Sub SuspendSprite(ByVal srcImgID As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Lz4, Optional ByVal autoKeepIfLarge As Boolean = True)
    If (Not m_Sprites Is Nothing) Then m_Sprites.SuspendCachedImage srcImgID, cmpFormat, autoKeepIfLarge
End Sub
