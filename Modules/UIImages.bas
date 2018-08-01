Attribute VB_Name = "UIImages"
'***************************************************************************
'PhotoDemon Central UI image cache
'Copyright 2018-2018 by Tanner Helland
'Created: 13/July/18
'Last updated: 13/July/18
'Last update: initial build
'
'PhotoDemon uses a *lot* of UI images.  The sheer amount of GDI objects required for these surfaces
' is substantial, and we can greatly reduce our requirements by using something akin to "sprite sheets",
' e.g. shared image storage when images have similar dimensions.
'
'At present, this module accepts images of any size, but it only provides a benefit when images are
' the *same* size - this allows it to automatically "coalesce" images into shared sheets, which callers
' can then access by index (rather than managing their own pdDIB instance).
'
'This module really only makes sense for images that are kept alive for the duration of the program.
' One-off images (e.g. temp images) should *not* be used, as it is non-trivial to release shared images
' in a performance-friendly manner.
'
'At present, PD limits usage of this cache to pdButtonToolbox images.  (They are basically the perfect
' use-case for shared caching.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Number of images allowed on a single sheet column.  Once the number of images on a sheet exceeds this,
' a new column will be created.  (The number of allowed columns is currently unbounded.)
Private Const MAX_SPRITES_IN_COLUMN As Long = 8

'Individual cache object.  This module manages a one-dimensional array of these headers.
Private Type ImgCacheEntry
    SpriteWidth As Long
    SpriteHeight As Long
    NumImages As Long
    ImgSpriteSheet As pdDIB
    SpriteNames As pdStringStack
End Type

'Cheap way to "fake" integer access inside a long
Private Type FakeDWord
    WordOne As Integer
    WordTwo As Integer
End Type

'The actual cache.  Resized dynamically as additional images are added.
Private m_ImageCache() As ImgCacheEntry
Private m_NumOfCacheObjects As Long

'Add an image to the cache.  The returned Long is the handle into the cache; you MUST remember it,
' as it's the only way to access the image again!
'
'When adding images to the cache, you must also pass a unique image name.  This ensures that cache
' entries are never duplicated, which is important as some images are reused throughout PD (and if
' each place the images are used attempts to add them to the cache, we waste time and memory).
' Note that the name is only required when *adding* images, so that we can perform a duplicate check.
' Once added, an image's handle is all that's required to retrieve it.
'
'RETURNS: non-zero value if successful; zero if the function fails.
Public Function AddImage(ByRef srcDIB As pdDIB, ByRef uniqueImageName As String) As Long

    'Failsafe checks
    If (srcDIB Is Nothing) Then Exit Function
    If (LenB(uniqueImageName) = 0) Then Exit Function
    
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
            If (m_ImageCache(i).SpriteWidth = targetWidth) Then
                If (m_ImageCache(i).SpriteHeight = targetHeight) Then
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
        
        'Before adding this sprite, perform a quick check for duplicate IDs.  If one is found,
        ' return the existing sprite instead of adding it anew.
        targetID = m_ImageCache(targetIndex).SpriteNames.ContainsString(uniqueImageName, True) + 1
        
        If (targetID = 0) Then
        
            'We have an existing sprite sheet with dimensions identical to this one!  Figure out
            ' if we need to resize the sprite sheet to account for another addition to it.
            Dim targetRow As Long, targetColumn As Long
            GetNumRowsColumns targetIndex, m_ImageCache(targetIndex).NumImages, targetRow, targetColumn
            
            'If this sprite sheet is still only one-column tall, we may need to resize it vertically
            Dim newDibRequired As Boolean
            With m_ImageCache(targetIndex)
            
                If (targetColumn = 0) Then
                    newDibRequired = ((targetRow + 1) * .SpriteHeight) > .ImgSpriteSheet.GetDIBHeight
                
                'Otherwise, we may need to resize it horizontally
                Else
                    newDibRequired = ((targetColumn + 1) * .SpriteWidth) > .ImgSpriteSheet.GetDIBWidth
                End If
            
            End With
            
            'If a new sprite sheet is required, create one now
            If newDibRequired Then
                
                With m_ImageCache(targetIndex)
                    
                    Dim tmpDIB As pdDIB
                    Set tmpDIB = New pdDIB
                    If (targetColumn = 0) Then
                        tmpDIB.CreateBlank .SpriteWidth, .SpriteHeight * (.NumImages + 1), 32, 0, 0
                        tmpDIB.SetInitialAlphaPremultiplicationState True
                        GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, .SpriteWidth, .ImgSpriteSheet.GetDIBHeight, .ImgSpriteSheet.GetDIBDC, 0, 0, vbSrcCopy
                        Set .ImgSpriteSheet = tmpDIB
                    Else
                        
                        'When adding a new column to a DIB, we *leave* the DIB at its maximum row size
                        tmpDIB.CreateBlank .SpriteWidth * (targetColumn + 1), .ImgSpriteSheet.GetDIBHeight, 32, 0, 0
                        tmpDIB.SetInitialAlphaPremultiplicationState True
                        GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, .ImgSpriteSheet.GetDIBWidth, .ImgSpriteSheet.GetDIBHeight, .ImgSpriteSheet.GetDIBDC, 0, 0, vbSrcCopy
                        Set .ImgSpriteSheet = tmpDIB
                        
                    End If
                    
                End With
            
            End If
            
            'Paint the new DIB into place, and update all target references to reflect the correct index
            With m_ImageCache(targetIndex)
                GDI.BitBltWrapper .ImgSpriteSheet.GetDIBDC, targetColumn * .SpriteWidth, targetRow * .SpriteHeight, .SpriteWidth, .SpriteHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
                .NumImages = .NumImages + 1
                targetID = .NumImages
                .SpriteNames.AddString uniqueImageName
            End With
            
        Else
            'Duplicate entry found; that's okay - reuse it as-is!
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
            
            .SpriteWidth = targetWidth
            .SpriteHeight = targetHeight
            .NumImages = 1
            targetID = .NumImages
            
            'Create the first sprite sheet entry
            Set .ImgSpriteSheet = New pdDIB
            .ImgSpriteSheet.CreateFromExistingDIB srcDIB
            
            'Add this sprite's name to the collection
            Set .SpriteNames = New pdStringStack
            .SpriteNames.AddString uniqueImageName
            
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
    tmpDWord.WordOne = targetIndex
    tmpDWord.WordTwo = targetID
    
    CopyMemoryStrict VarPtr(AddImage), VarPtr(tmpDWord), 4
    
    'Finally, free the target sprite sheet from its DC; the DC will automatically be re-created as necessary
    m_ImageCache(targetIndex).ImgSpriteSheet.FreeFromDC
    
End Function

Public Function PaintCachedImage(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcImgID As Long) As Boolean

    'Resolve the image ID into a target index and image number
    Dim targetIndex As Long, imgNumber As Long, tmpDWord As FakeDWord
    CopyMemoryStrict VarPtr(tmpDWord), VarPtr(srcImgID), 4
    targetIndex = tmpDWord.WordOne
    imgNumber = tmpDWord.WordTwo - 1
    
    'Failsafe checks
    If (targetIndex > UBound(m_ImageCache)) Then Exit Function
    
    'Resolve the image number into a sprite row and column
    Dim targetRow As Long, targetColumn As Long
    GetNumRowsColumns targetIndex, imgNumber, targetRow, targetColumn
    
    'Paint the result!
    If (Not m_ImageCache(targetIndex).ImgSpriteSheet Is Nothing) Then
        With m_ImageCache(targetIndex)
            .ImgSpriteSheet.AlphaBlendToDCEx dstDC, dstX, dstY, .SpriteWidth, .SpriteHeight, targetColumn * .SpriteWidth, targetRow * .SpriteHeight, .SpriteWidth, .SpriteHeight
            .ImgSpriteSheet.FreeFromDC
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

'Fully reset the cache.  NOTE: this will invalidate all previously returned handles, so you *must*
' re-add any required images to the cache.
Public Sub ResetCache()
    ReDim m_ImageCache(0) As ImgCacheEntry
    m_NumOfCacheObjects = 0
End Sub

Public Sub TestCacheOnly()
    GDI.BitBltWrapper pdImages(g_CurrentImage).GetActiveDIB.GetDIBDC, 0, 0, m_ImageCache(0).ImgSpriteSheet.GetDIBWidth, m_ImageCache(0).ImgSpriteSheet.GetDIBHeight, m_ImageCache(0).ImgSpriteSheet.GetDIBDC, 0, 0, vbSrcCopy
    pdImages(g_CurrentImage).NotifyImageChanged UNDO_Everything
    ViewportEngine.Stage4_FlipBufferAndDrawUI pdImages(g_CurrentImage), FormMain.MainCanvas(0)
End Sub
