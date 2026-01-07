VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSpriteSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon spritesheet cache
'Copyright 2018-2026 by Tanner Helland
'Created: 13/July/18
'Last updated: 26/August/20
'Last update: refactor cache to allow for multiple sheets per size; sheet limits are now
'             calculated on the fly, using a max memory size (as opposed to a max number
'             of images per sheet).
'
'When storing large amounts of small images, it is more resource- and performance-friendly
' to stick the images in dedicated spritesheets.
'
'This class can be used as an arbitrary spritesheet manager for any group of source images.
' For best results, the images should all be identical sizes, but this class will still
' function correctly if images are different sizes.  Just note that separate spritesheets
' are created for each set of source sizes - so if all images are different sizes, you don't
' gain anything by storing them here.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This value can be set by the caller (it defaults to a multiple of sprite size and
' max memory per sheet), but it is VERY IMPORTANT to note that it SHOULD NOT BE CHANGED
' AFTER ONE OR MORE IMAGES HAVE BEEN LOADED.  Images are just indices into a sprite list,
' and changing the max column size wrecks existing indices - so don't attempt this.
Private m_MaxSpritesInColumn As Long
Private Const DEFAULT_MAX_SIZE_OF_SHEET As Long = 8000000  'Entirely arbitrary value; smaller may be better for some use-cases?
Private Const MAX_IMAGES_PER_SHEET As Long = 32     'Again, entirely arbitrary; designed to minimize excessive allocations

'Individual cache object.  This module manages a one-dimensional array of these headers.
Private Type ImgCacheEntry
    spriteWidth As Long
    spriteHeight As Long
    numImages As Long
    numSheets As Long
    imgSpriteSheets() As pdDIB
    spriteNames As pdStringStack
End Type

'Cheap way to "fake" integer access inside a long (and keep VB happy)
Private Type FakeDWord
    wordOne As Integer
    wordTwo As Integer
End Type

'The actual cache.  Resized dynamically as additional images are added.
Private m_ImageCache() As ImgCacheEntry
Private m_NumOfCacheObjects As Long

'The last sheet used for adding images and/or painting.  Whenever this value changes, we suspend the
' old DIB (e.g. suspend it to a compressed memory stream) to minimize resources.
Private m_LastSheetAccessed As Long

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
Friend Function AddImage(ByRef srcDIB As pdDIB, ByRef uniqueImageName As String) As Long

    'Failsafe checks
    If (srcDIB Is Nothing) Then Exit Function
    If (LenB(uniqueImageName) = 0) Then Exit Function
    
    Dim i As Long
    
    Dim targetWidth As Long, targetHeight As Long
    targetWidth = srcDIB.GetDIBWidth
    targetHeight = srcDIB.GetDIBHeight
    
    'If we haven't calculated a max spritesheet size, now is the time to do so
    If (m_MaxSpritesInColumn = 0) Then
        m_MaxSpritesInColumn = DEFAULT_MAX_SIZE_OF_SHEET \ (targetWidth * targetHeight * 4)
        If (m_MaxSpritesInColumn < 2) Then m_MaxSpritesInColumn = 2
        If (m_MaxSpritesInColumn > MAX_IMAGES_PER_SHEET) Then m_MaxSpritesInColumn = MAX_IMAGES_PER_SHEET
    End If
    
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
    
    '(Inside our target index, we may also need to find a target "sheet"; very large sheets
    ' may produce multiple smaller columns.)
    Dim targetSheet As Long
    targetSheet = 0
    
    'If we found a sprite sheet that matches our target size, we just need to append this
    ' new image to it.
    If (targetIndex >= 0) Then
        
        'Before adding this sprite, perform a quick check for duplicate IDs.  If one is found,
        ' return the existing sprite instead of adding it anew.
        targetID = m_ImageCache(targetIndex).spriteNames.ContainsString(uniqueImageName, True) + 1
        If (targetID = 0) Then
        
            'We have an existing sprite sheet with dimensions identical to this one, and it
            ' doesn't already contain a sprite with this name!  Figure out if we need to resize
            ' the target sprite sheet, or add another sheet entirely.
            Dim targetRow As Long, targetColumn As Long
            GetNumRowsColumns m_ImageCache(targetIndex).numImages, targetRow, targetColumn
            
            'Two circumstances require us to create a new DIB:
            ' 1) The current sheet is not tall enough to hold this DIB.  Sheets are resized
            '    using a standard power-of-2 strategy, until they exceed the maximum number
            '    of images allowed in a sheet.
            ' 2) This image belongs in a new DIB, and that DIB doesn't exist yet.
            With m_ImageCache(targetIndex)
                
                'First, see if a new sheet is required.  (This occurs when we've filled the current
                ' sheet to its max-sprites-allowed value.)
                If (targetColumn >= .numSheets) Then
                    ReDim Preserve .imgSpriteSheets(0 To .numSheets) As pdDIB
                    Set .imgSpriteSheets(.numSheets) = New pdDIB
                    .imgSpriteSheets(.numSheets).CreateBlank .spriteWidth, .spriteHeight, 32, 0, 0
                    .imgSpriteSheets(.numSheets).SetInitialAlphaPremultiplicationState True
                    .numSheets = .numSheets + 1
                    
                'If a sheet already exists for the target column, check to ensure the column is
                ' large enough to hold it.
                Else
                    
                    If (((targetRow + 1) * .spriteHeight) >= .imgSpriteSheets(targetColumn).GetDIBHeight) Then
                    
                        'The target column needs to be enlarged.  Figure out a new height, but do *not*
                        ' exceed that max height.
                        Dim newRowCount As Long
                        newRowCount = (targetRow * 2 - 1)
                        If (newRowCount > m_MaxSpritesInColumn - 1) Then newRowCount = m_MaxSpritesInColumn - 1
                        
                        Dim tmpDIB As pdDIB
                        Set tmpDIB = New pdDIB
                        tmpDIB.CreateBlank .spriteWidth, .spriteHeight * (newRowCount + 1), 32, 0, 0
                        tmpDIB.SetInitialAlphaPremultiplicationState True
                        GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, .spriteWidth, .imgSpriteSheets(targetColumn).GetDIBHeight, .imgSpriteSheets(targetColumn).GetDIBDC, 0, 0, vbSrcCopy
                        Set .imgSpriteSheets(targetColumn) = tmpDIB
                        
                    End If
                    
                End If
            
                'Paint the new DIB into place, and update all target references to reflect the correct index
                GDI.BitBltWrapper .imgSpriteSheets(targetColumn).GetDIBDC, 0, targetRow * .spriteHeight, .spriteWidth, .spriteHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
                .imgSpriteSheets(targetColumn).FreeFromDC
                srcDIB.FreeFromDC
                .numImages = .numImages + 1
                targetID = .numImages
                .spriteNames.AddString uniqueImageName
                
                'Keep memory usage down by suspending previous sheets
                If (targetColumn <> m_LastSheetAccessed) And (m_LastSheetAccessed >= 0) And (m_LastSheetAccessed <= UBound(.imgSpriteSheets)) Then .imgSpriteSheets(m_LastSheetAccessed).SuspendDIB cf_Lz4, False
                m_LastSheetAccessed = targetColumn
                
            End With
            
        Else
            
            'Duplicate entry found; that's okay - replace the existing version, then return the index as-is!
            Dim imgID As Long
            imgID = GetCombinedIndexAndID(targetIndex, targetID)
            
            'Failsafe check only
            If GetRowColumnPos(imgID, targetIndex, targetRow, targetColumn) Then
                
                With m_ImageCache(targetIndex)
                    GDI.BitBltWrapper .imgSpriteSheets(targetColumn).GetDIBDC, 0, targetRow * .spriteHeight, .spriteWidth, .spriteHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
                    .imgSpriteSheets(targetColumn).FreeFromDC
                    srcDIB.FreeFromDC
                End With
                
            End If
            
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
            .numSheets = 1
            targetSheet = 0
            
            'Create the first sprite sheet entry
            ReDim .imgSpriteSheets(0) As pdDIB
            Set .imgSpriteSheets(0) = New pdDIB
            .imgSpriteSheets(0).CreateFromExistingDIB srcDIB
            .imgSpriteSheets(0).FreeFromDC
            
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
    AddImage = GetCombinedIndexAndID(targetIndex, targetID)
    
    'Finally, free the target sprite sheet from its DC; the DC will automatically be re-created as necessary
    m_ImageCache(targetIndex).imgSpriteSheets(targetSheet).FreeFromDC
    
End Function

Friend Function DoesImageExist(ByRef srcImageName As String) As Boolean
    
    DoesImageExist = False
    
    If (m_NumOfCacheObjects > 0) Then
        
        'Look for a cache with matching dimensions
        Dim i As Long
        For i = 0 To m_NumOfCacheObjects - 1
            If (m_ImageCache(i).spriteNames.ContainsString(srcImageName, False) >= 0) Then
                DoesImageExist = True
                Exit Function
            End If
        Next i
        
    End If
    
End Function

Private Function GetCombinedIndexAndID(ByVal targetIndex As Long, ByVal targetID As Long) As Long
    Dim tmpDWord As FakeDWord
    tmpDWord.wordOne = targetIndex
    tmpDWord.wordTwo = targetID
    GetMem4 VarPtr(tmpDWord), GetCombinedIndexAndID
End Function

'BitBlt equivalent
Friend Function CopyCachedImage(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcImgID As Long) As Boolean

    'Resolve the image number into a sprite row and column
    Dim targetIndex As Long, targetRow As Long, targetColumn As Long
    If GetRowColumnPos(srcImgID, targetIndex, targetRow, targetColumn) Then
        
        'Paint the result!
        If (m_ImageCache(targetIndex).numSheets > 0) Then
            With m_ImageCache(targetIndex)
                GDI.BitBltWrapper dstDC, dstX, dstY, .spriteWidth, .spriteHeight, .imgSpriteSheets(targetColumn).GetDIBDC, 0, targetRow * .spriteHeight, vbSrcCopy
                .imgSpriteSheets(targetColumn).FreeFromDC
            End With
        Else
            PDDebug.LogAction "WARNING!  pdSpriteSheet.CopyCachedImage failed to paint image number " & srcImgID & " in spritesheet " & targetIndex
        End If
        
    End If
    
End Function

'Retrieve dimensions of a given sprite; not usually required by the client
Friend Function GetSpriteHeight(ByVal srcImgID As Long) As Long
    
    'Resolve the image number into a sprite row and column
    Dim targetIndex As Long, targetRow As Long, targetColumn As Long
    If GetRowColumnPos(srcImgID, targetIndex, targetRow, targetColumn) Then
        
        'Query the parent UDT for sprite dimensions
        If (m_ImageCache(targetIndex).numSheets > 0) Then
            GetSpriteHeight = m_ImageCache(targetIndex).spriteHeight
        Else
            PDDebug.LogAction "WARNING!  pdSpriteSheet.GetSpriteHeight failed to query image number " & srcImgID & " in spritesheet " & targetIndex
        End If
        
    End If
    
End Function

'Return the text name of a given sprite
Friend Function GetSpriteName(ByVal srcImgID As Long) As String
    
    Dim targetIndex As Long, targetRow As Long, targetColumn As Long
    If GetRowColumnPos(srcImgID, targetIndex, targetRow, targetColumn) Then
        If (m_ImageCache(targetIndex).numSheets > 0) Then
            GetSpriteName = m_ImageCache(targetIndex).spriteNames.GetString(targetColumn * m_MaxSpritesInColumn + targetRow)
        Else
            PDDebug.LogAction "WARNING!  pdSpriteSheet.GetSpriteName failed to query image number " & srcImgID & " in spritesheet " & targetIndex
        End If
    End If
    
End Function

'For detailed comments, see GetSpriteHeight() above
Friend Function GetSpriteWidth(ByVal srcImgID As Long) As Long
    
    Dim targetIndex As Long, targetRow As Long, targetColumn As Long
    If GetRowColumnPos(srcImgID, targetIndex, targetRow, targetColumn) Then
        If (m_ImageCache(targetIndex).numSheets > 0) Then
            GetSpriteWidth = m_ImageCache(targetIndex).spriteWidth
        Else
            PDDebug.LogAction "WARNING!  pdSpriteSheet.GetSpriteHeight failed to query image number " & srcImgID & " in spritesheet " & targetIndex
        End If
    End If
    
End Function

'AlphaBlend equivalent.
'
'The optional "allowMinimizeMemory " parameter will auto-suspend sheets to memory when the caller switches to
' a new sheet; this may introduce subtle delays to the rendering, but it can produce very large memory savings -
' so try to balance those competing needs at run-time.
Friend Function PaintCachedImage(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal srcImgID As Long, Optional ByVal customAlpha As Long = 255, Optional ByVal allowMinimizeMemory As Boolean = True) As Boolean

    'Resolve the image number into a sprite row and column
    Dim targetIndex As Long, targetRow As Long, targetColumn As Long
    If GetRowColumnPos(srcImgID, targetIndex, targetRow, targetColumn) Then
        
        'Paint the result!
        If (m_ImageCache(targetIndex).numSheets > 0) Then
            
            With m_ImageCache(targetIndex)
                .imgSpriteSheets(targetColumn).AlphaBlendToDCEx dstDC, dstX, dstY, .spriteWidth, .spriteHeight, 0, targetRow * .spriteHeight, .spriteWidth, .spriteHeight, customAlpha
                .imgSpriteSheets(targetColumn).FreeFromDC
            End With
            PaintCachedImage = True
            
            'Free past sheets as-we-go, if the user allows
            If allowMinimizeMemory Then
                If (m_LastSheetAccessed <> targetColumn) And (m_LastSheetAccessed >= 0) And (m_LastSheetAccessed <= UBound(m_ImageCache(targetIndex).imgSpriteSheets)) Then m_ImageCache(targetIndex).imgSpriteSheets(m_LastSheetAccessed).SuspendDIB cf_Lz4, False
            End If
            
            m_LastSheetAccessed = targetColumn
            
        Else
            PDDebug.LogAction "WARNING!  pdSpriteSheet.PaintCachedImage failed to paint image number " & srcImgID & " in spritesheet " & targetIndex
        End If
        
    End If
    
End Function

'GDI+ StretchBlt equivalent
Friend Function StretchBltCachedImage(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcImgID As Long, Optional ByVal newAlpha As Single = 1!, Optional ByVal interpolationType As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal useThisDestinationDCInstead As Long = 0, Optional ByVal disableEdgeFix As Boolean = False, Optional ByVal isZoomedIn As Boolean = False, Optional ByVal dstCopyIsOkay As Boolean = False) As Boolean
    
    'Resolve the image number into a sprite row and column
    Dim targetIndex As Long, targetRow As Long, targetColumn As Long
    If GetRowColumnPos(srcImgID, targetIndex, targetRow, targetColumn) Then
        
        'Paint the result!
        If (m_ImageCache(targetIndex).numSheets > 0) Then
            With m_ImageCache(targetIndex)
                GDI_Plus.GDIPlus_StretchBlt dstDIB, x1, y1, dstWidth, dstHeight, .imgSpriteSheets(targetColumn), 0!, targetRow * .spriteHeight, .spriteWidth, .spriteHeight, newAlpha, interpolationType, useThisDestinationDCInstead, disableEdgeFix, isZoomedIn, dstCopyIsOkay
                .imgSpriteSheets(targetColumn).FreeFromDC
            End With
        Else
            PDDebug.LogAction "WARNING!  pdSpriteSheet.StretchBltCachedImage failed to paint image number " & srcImgID & " in spritesheet " & targetIndex
        End If
        
    End If
    
End Function

'After painting (or adding, if you want) a sprite to this sprite sheet object,
' you can call this function to suspend the underlying sheet to a compressed stream
' of bytes.  This also frees all GDI objects associated with the image.
Friend Sub SuspendCachedImage(ByVal srcImgID As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Lz4, Optional ByVal autoKeepIfLarge As Boolean = True)

    'Resolve the image number into a sprite row and column, then suspend that image
    Dim targetIndex As Long, targetRow As Long, targetColumn As Long
    If GetRowColumnPos(srcImgID, targetIndex, targetRow, targetColumn) Then
        
        If (m_ImageCache(targetIndex).numSheets > 0) Then
            m_ImageCache(targetIndex).imgSpriteSheets(targetColumn).SuspendDIB cmpFormat, autoKeepIfLarge
        Else
            PDDebug.LogAction "WARNING!  pdSpriteSheet.SuspendCachedImage failed to suspend image number " & srcImgID & " in spritesheet " & targetIndex
        End If
        
    End If
    
End Sub

Private Function GetRowColumnPos(ByVal srcImgID As Long, ByRef dstIndex As Long, ByRef dstRow As Long, ByRef dstColumn As Long) As Boolean
    
    GetRowColumnPos = True
    
    Dim imgNumber As Long, tmpDWord As FakeDWord
    PutMem4 VarPtr(tmpDWord), srcImgID
    dstIndex = tmpDWord.wordOne
    imgNumber = tmpDWord.wordTwo - 1
    
    'Failsafe check
    If (dstIndex > UBound(m_ImageCache)) Then
        GetRowColumnPos = False
    
    'Resolve the image number into a sprite row and column
    Else
        GetNumRowsColumns imgNumber, dstRow, dstColumn
    End If
    
End Function

'Return the row and column location [0-based] of entry (n) in a target cache entry.
Private Sub GetNumRowsColumns(ByVal srcImageIndex As Long, ByRef dstRow As Long, ByRef dstColumn As Long)
    dstRow = srcImageIndex Mod m_MaxSpritesInColumn
    dstColumn = srcImageIndex \ m_MaxSpritesInColumn
End Sub

'Suspend all spritesheets to compressed memory streams; they will auto-recreate as necessary
Friend Sub MinimizeMemory(Optional ByVal cmpFormat As PD_CompressionFormat = cf_Lz4, Optional ByVal autoKeepIfLarge As Boolean = True)
    
    If (m_NumOfCacheObjects > 0) Then
        
        Dim i As Long, j As Long
        For i = 0 To m_NumOfCacheObjects - 1
            For j = 0 To m_ImageCache(i).numSheets - 1
                m_ImageCache(i).imgSpriteSheets(j).SuspendDIB cmpFormat, autoKeepIfLarge
            Next j
        Next i
        
    End If
    
End Sub

'Fully reset the cache.  NOTE: this will invalidate all previously returned handles, so you *must*
' re-add any required images to the cache.
Friend Sub ResetCache()
    ReDim m_ImageCache(0) As ImgCacheEntry
    m_NumOfCacheObjects = 0
End Sub

Friend Sub SetMaxSpritesInColumn(ByVal newMax As Long)
    If (m_NumOfCacheObjects = 0) Then
        m_MaxSpritesInColumn = newMax
    Else
        PDDebug.LogAction "WARNING!  You cannot change the maximum sprite count *after* caching images!"
    End If
End Sub

Private Sub Class_Initialize()
    m_MaxSpritesInColumn = 0
End Sub
