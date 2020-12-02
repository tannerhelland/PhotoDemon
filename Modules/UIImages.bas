Attribute VB_Name = "UIImages"
'***************************************************************************
'PhotoDemon Central UI image cache
'Copyright 2018-2020 by Tanner Helland
'Created: 13/July/18
'Last updated: 27/August/20
'Last update: remove ResetCache function; this was causing issues if the user loaded the theme dialog,
'             made changes, then *canceled* the dialog (because the icon cache was being cleared,
'             but UI elements were detecting an identical theme to their previous render and were thus
'             skipping icon reloading steps).  Instead, duplicate icons are intelligently located and
'             updated in-place.  If the theme dialog is canceled, indices still point to valid
'             locations in the icon table, and no extra work is required on our end regardless of
'             whether the user accepts or cancels the theme dialog.
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

'Number of images allowed on a single sheet column.  Once the number of images on a sheet exceeds this,
' a new column will be created.  (The number of allowed columns is currently unbounded.)
Private Const MAX_SPRITES_IN_COLUMN As Long = 8

'Individual cache object.  This module manages a one-dimensional array of these headers.
Private Type ImgCacheEntry
    spriteWidth As Long
    spriteHeight As Long
    numImages As Long
    ImgSpriteSheet As pdDIB
    spriteNames As pdStringStack
End Type

'Cheap way to "fake" integer access inside a long
Private Type FakeDWord
    wordOne As Integer
    wordTwo As Integer
End Type

'The actual cache.  Resized dynamically as additional images are added.
Private m_ImageCache() As ImgCacheEntry
Private m_NumOfCacheObjects As Long

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
    
    'Failsafe checks
    If (srcDIB Is Nothing) Then
        PDDebug.LogAction "WARNING!  UIImages.AddImage was passed a null DIB"
        Exit Function
    End If
    
    If (LenB(uniqueImageName) = 0) Then
        PDDebug.LogAction "WARNING!  UIImages.AddImage was passed a zero-length DIB name"
        Exit Function
    End If
    
    Dim i As Long
    
    Dim targetWidth As Long, targetHeight As Long
    targetWidth = srcDIB.GetDIBWidth
    targetHeight = srcDIB.GetDIBHeight
        
    'Our first task is finding a matching spritesheet - specifically, a spritesheet where the sprites
    ' have the same dimensions as this image.
    Dim targetIndex As Long
    targetIndex = -1
    
    If (m_NumOfCacheObjects > 0) Then
        
        'Look for a cache with matching dimensions
        For i = 0 To m_NumOfCacheObjects - 1
            If (m_ImageCache(i).spriteWidth = targetWidth) Then
                If (m_ImageCache(i).spriteHeight = targetHeight) Then
                    targetIndex = i
                    Exit For
                End If
            End If
        Next i
        
    End If
    
    'The last piece of the puzzle is a "target ID", e.g. the location of this image within the
    ' relevant sprite sheet.
    Dim targetID As Long
    targetID = -1
    
    'If we found a sprite sheet that matches our target size, we just need to append this
    ' new image to it.
    If (targetIndex >= 0) Then
        
        Dim targetRow As Long, targetColumn As Long
        
        'Before adding this sprite, perform a quick check for duplicate IDs.  If one is found,
        ' return the existing sprite instead of adding it anew.
        targetID = m_ImageCache(targetIndex).spriteNames.ContainsString(uniqueImageName, True) + 1
        
        If (targetID = 0) Then
        
            'We have an existing sprite sheet with dimensions identical to this one!  Figure out
            ' if we need to resize the sprite sheet to account for another addition to it.
            GetNumRowsColumns targetIndex, m_ImageCache(targetIndex).numImages, targetRow, targetColumn
            
            'If this sprite sheet is still only one-column tall, we may need to resize it vertically
            Dim newDibRequired As Boolean
            With m_ImageCache(targetIndex)
            
                If (targetColumn = 0) Then
                    newDibRequired = ((targetRow + 1) * .spriteHeight) > .ImgSpriteSheet.GetDIBHeight
                
                'Otherwise, we may need to resize it horizontally
                Else
                    newDibRequired = ((targetColumn + 1) * .spriteWidth) > .ImgSpriteSheet.GetDIBWidth
                End If
            
            End With
            
            'If a new sprite sheet is required, create one now
            If newDibRequired Then
                
                With m_ImageCache(targetIndex)
                    
                    Dim tmpDIB As pdDIB
                    Set tmpDIB = New pdDIB
                    If (targetColumn = 0) Then
                        tmpDIB.CreateBlank .spriteWidth, .spriteHeight * (.numImages + 1), 32, 0, 0
                        tmpDIB.SetInitialAlphaPremultiplicationState True
                        GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, .spriteWidth, .ImgSpriteSheet.GetDIBHeight, .ImgSpriteSheet.GetDIBDC, 0, 0, vbSrcCopy
                        Set .ImgSpriteSheet = tmpDIB
                    Else
                        
                        'When adding a new column to a DIB, we *leave* the DIB at its maximum row size
                        tmpDIB.CreateBlank .spriteWidth * (targetColumn + 1), .ImgSpriteSheet.GetDIBHeight, 32, 0, 0
                        tmpDIB.SetInitialAlphaPremultiplicationState True
                        GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, .ImgSpriteSheet.GetDIBWidth, .ImgSpriteSheet.GetDIBHeight, .ImgSpriteSheet.GetDIBDC, 0, 0, vbSrcCopy
                        Set .ImgSpriteSheet = tmpDIB
                        
                    End If
                    
                End With
                
                'Suspend the previous DIB in line, as it may not be accessed again for awhile
                If (targetIndex > 0) Then m_ImageCache(targetIndex - 1).ImgSpriteSheet.SuspendDIB
            
            End If
            
            'Paint the new DIB into place, and update all target references to reflect the correct index
            With m_ImageCache(targetIndex)
                GDI.BitBltWrapper .ImgSpriteSheet.GetDIBDC, targetColumn * .spriteWidth, targetRow * .spriteHeight, .spriteWidth, .spriteHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
                .ImgSpriteSheet.FreeFromDC
                .numImages = .numImages + 1
                targetID = .numImages
                .spriteNames.AddString uniqueImageName
            End With
        
        'Duplicate entries are okay!  These can occur after the user changes the UI theme from
        ' e.g. color to monochrome icons; all icons already exist, but they need to be updated
        ' with their new monochrome equivalents.  To accomplish this, we just want to update
        ' the image in-place with whatever new version we've been passed.
        Else
            With m_ImageCache(targetIndex)
                GetNumRowsColumns targetIndex, targetID - 1, targetRow, targetColumn
                GDI.BitBltWrapper .ImgSpriteSheet.GetDIBDC, targetColumn * .spriteWidth, targetRow * .spriteHeight, .spriteWidth, .spriteHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
                .ImgSpriteSheet.FreeFromDC
            End With
        End If
            
    'If we didn't find a matching spritesheet, we must create a new one
    Else
        
        If (m_NumOfCacheObjects = 0) Then
            ReDim m_ImageCache(0) As ImgCacheEntry
        Else
            ReDim Preserve m_ImageCache(0 To m_NumOfCacheObjects) As ImgCacheEntry
        End If
        
        'Prep a generic header
        With m_ImageCache(m_NumOfCacheObjects)
            
            .spriteWidth = targetWidth
            .spriteHeight = targetHeight
            .numImages = 1
            targetID = .numImages
            
            'Create the first sprite sheet entry
            Set .ImgSpriteSheet = New pdDIB
            .ImgSpriteSheet.CreateFromExistingDIB srcDIB
            .ImgSpriteSheet.FreeFromDC
            
            'Add this sprite's name to the collection
            Set .spriteNames = New pdStringStack
            .spriteNames.AddString uniqueImageName
            
        End With
        
        targetIndex = m_NumOfCacheObjects
        
        'Increment the cache object count prior to exiting
        m_NumOfCacheObjects = m_NumOfCacheObjects + 1
        
    End If
    
    'Before exiting, we now need to return an index into our table.  We use a simple formula for this:
    ' 4-byte long
    '   - 1st 2-bytes: index into the cache
    '   - 2nd 2-bytes: index into that cache object's spritesheet
    Dim tmpDWord As FakeDWord
    tmpDWord.wordOne = targetIndex
    tmpDWord.wordTwo = targetID
    
    CopyMemoryStrict VarPtr(AddImage), VarPtr(tmpDWord), 4
    
    'Finally, free the target sprite sheet from its DC; the DC will automatically be re-created as necessary
    m_ImageCache(targetIndex).ImgSpriteSheet.FreeFromDC
    
End Function

Public Sub FreeSharedCompressBuffer()
    Erase m_TempCompressBuffer
    m_CompressBufferSize = 0
End Sub

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

Public Function PaintCachedImage(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcImgID As Long) As Boolean

    'Resolve the image ID into a target index and image number
    Dim targetIndex As Long, imgNumber As Long, tmpDWord As FakeDWord
    PutMem4 VarPtr(tmpDWord), srcImgID
    targetIndex = tmpDWord.wordOne
    imgNumber = tmpDWord.wordTwo - 1
    
    'Failsafe checks
    If (targetIndex > UBound(m_ImageCache)) Then
        PDDebug.LogAction "WARNING!  Failed to resolve index into UI image cache."
        Exit Function
    End If
    
    'Resolve the image number into a sprite row and column
    Dim targetRow As Long, targetColumn As Long
    GetNumRowsColumns targetIndex, imgNumber, targetRow, targetColumn
    
    'Paint the result!
    If (Not m_ImageCache(targetIndex).ImgSpriteSheet Is Nothing) Then
        With m_ImageCache(targetIndex)
            .ImgSpriteSheet.AlphaBlendToDCEx dstDC, dstX, dstY, .spriteWidth, .spriteHeight, targetColumn * .spriteWidth, targetRow * .spriteHeight, .spriteWidth, .spriteHeight
            .ImgSpriteSheet.SuspendDIB
        End With
    Else
        PDDebug.LogAction "WARNING!  UIImages.PaintCachedImage failed to paint image number " & imgNumber & " in spritesheet " & targetIndex
    End If
    
End Function

'Return the row and column location [0-based] of entry (n) in a target cache entry.
Private Sub GetNumRowsColumns(ByVal srcCacheIndex As Long, ByVal srcImageIndex As Long, ByRef dstRow As Long, ByRef dstColumn As Long)
    dstRow = srcImageIndex Mod MAX_SPRITES_IN_COLUMN
    dstColumn = srcImageIndex \ MAX_SPRITES_IN_COLUMN
End Sub
