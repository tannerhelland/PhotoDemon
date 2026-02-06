Attribute VB_Name = "PDImages"
'***************************************************************************
'Image Canvas Handler (formerly Image Window Handler)
'Copyright 2002-2026 by Tanner Helland
'Created: 11/29/02
'Last updated: 10/March/25
'Last update: new "StrategicMemoryReduction" function to reduce memory when multiple images are loaded
'
'In "ye good ol' days", PhotoDemon exposed the collection of currently loaded user images as a bare array.
' This was a terrible idea (for too many reasons to count).
'
'These days, the open image collection is instead managed by this module.  This provides much more
' flexibility in how we manage the collection "behind-the-scenes", with improved options for tasks
' like memory management (e.g. suspending inactive images to disk).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The number of images PhotoDemon has loaded this session (always goes up, never down; starts at zero when
' the program is loaded).  This value correlates to the upper bound of the primary pdImages array.
' For performance reasons, that array is not dynamically resized when images are loaded - the array stays
' the same size, and entries are deactivated as needed.  Thus, WHENEVER YOU NEED TO ITERATE THROUGH ALL
' LOADED IMAGES, USE THIS VALUE INSTEAD OF m_OpenImageCount (which only reflects the number of images
' *currently* open).
Private m_ImagesLoadedThisSession As Long

'The ID number (e.g. index in the pdImages array) of image the user is currently interacting with (e.g. the currently active image
' window).  Whenever a function needs to access the current image, use PDImages.GetActiveImage.
Private m_CurrentImageID As Long

'Number of image windows CURRENTLY OPEN.  This value goes up and down as images are opened or closed.  Use it to test for no open
' images (e.g. If m_OpenImageCount = 0...).  Note that this value SHOULD NOT BE USED FOR ITERATING OPEN IMAGES.  Instead, use
' m_ImagesLoadedThisSession, which will always match the upper bound of the m_PDImages() array, and never decrements, even when images
' are unloaded.
Private m_OpenImageCount As Long

'This array is the heart and soul of a given PD session.  Every time an image is loaded, all of its relevant data is stored within
' a new entry in this array.
Private m_PDImages() As pdImage

'Add an already-created pdImage object to the centrak m_PDImages() collection.
' DO NOT PASS AN EMPTY OBJECT!
Public Function AddImageToCentralCollection(ByRef srcImage As pdImage) As Boolean
    
    If (Not srcImage Is Nothing) Then
        
        Set m_PDImages(m_ImagesLoadedThisSession) = srcImage
        
        'Activate the image and assign it a unique ID.  (IMPORTANT: at present, the ID always correlates to the
        ' image's position in the collection.  Do not change this behavior.)
        m_PDImages(m_ImagesLoadedThisSession).ChangeActiveState True
        m_PDImages(m_ImagesLoadedThisSession).imageID = m_ImagesLoadedThisSession
        
        'Newly loaded images are always auto-activated.
        m_CurrentImageID = m_ImagesLoadedThisSession
    
        'Track how many images we've loaded and/or currently have open
        m_ImagesLoadedThisSession = m_ImagesLoadedThisSession + 1
        m_OpenImageCount = m_OpenImageCount + 1
        
        If (m_ImagesLoadedThisSession > UBound(m_PDImages)) Then
            ReDim Preserve m_PDImages(0 To m_ImagesLoadedThisSession * 2 - 1) As pdImage
        End If
        
        AddImageToCentralCollection = True
        
    Else
        AddImageToCentralCollection = False
    End If
    
End Function

'Reference to the currently active image.  This is the PD equivalent of the GetActiveWindow function.
' NOTE: if no images are loaded, this will return NOTHING, by design.
Public Function GetActiveImage() As pdImage
    If (m_ImagesLoadedThisSession > 0) Then Set GetActiveImage = m_PDImages(m_CurrentImageID)
End Function

'Return the ID of the currently active image.  (This is a thin wrapper to m_CurrentImageID.)
'Returns: a value >= 0 if an image is active; -1 otherwise
Public Function GetActiveImageID() As Long
    GetActiveImageID = m_CurrentImageID
End Function

'Pass this function to obtain a default pdImage object, instantiated to match current UI settings and
' user preferences.  Note that this function *does not touch* the main pdImages object, and as such,
' the created image will not yet have an imageID value.  An ID value will be assigned when the object
' is added to the main m_PDImages() collection (via AddImageToCentralCollection(), above).
Public Sub GetDefaultPDImageObject(ByRef dstImage As pdImage)
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    dstImage.SetZoomIndex Zoom.GetZoom100Index
End Sub

Public Function GetImageByID(ByVal imgID As Long) As pdImage
    If (m_ImagesLoadedThisSession > 0) Then
        If ((imgID >= LBound(m_PDImages)) And (imgID <= UBound(m_PDImages))) Then Set GetImageByID = m_PDImages(imgID)
    End If
End Function

'Populate a pdStack object with a list of currently active image IDs (e.g. a list of all open user images).
' The destination stack will be forcibly reset if already populated.
'
'RETURNS: TRUE if at least one image is active; FALSE otherwise
Public Function GetListOfActiveImageIDs(ByRef dstStack As pdStack) As Boolean

    If PDImages.IsImageNonNull() Then
        
        Set dstStack = New pdStack
        
        Dim i As Long
        For i = LBound(m_PDImages) To UBound(m_PDImages)
            If (Not m_PDImages(i) Is Nothing) Then
                If m_PDImages(i).IsActive Then dstStack.AddInt m_PDImages(i).imageID
            End If
        Next i
        
        GetListOfActiveImageIDs = True
        
    Else
        Set dstStack = Nothing
        GetListOfActiveImageIDs = False
    End If

End Function

'Given an image ID, find the next ID in the collection (moving either forward or backward).  This is used for
' navigating between images using keyboard next/previous shortcuts.
'
'NOTE: returns -1 if no valid prev/next image can be found; this will occur if only one image is loaded.
Public Function GetNextImageID(ByVal curImageID As Long, Optional ByVal moveForward As Boolean = True) As Long
    
    GetNextImageID = -1
    
    'If one (or zero) images are loaded, ignore this option
    If (m_OpenImageCount <= 1) Then Exit Function
    
    Dim i As Long
    
    'Loop through all available images, and when we find one that is not this image, activate it and exit
    If moveForward Then
        i = m_CurrentImageID + 1
    Else
        i = m_CurrentImageID - 1
    End If
    
    Do While (i <> m_CurrentImageID)
            
        'Loop back to the start of the window collection
        If moveForward Then
            If (i > m_ImagesLoadedThisSession) Then i = 0
            If (i > UBound(m_PDImages)) Then i = 0
        Else
            If (i < 0) Then i = m_ImagesLoadedThisSession
            If (i > UBound(m_PDImages)) Then i = UBound(m_PDImages)
        End If
                
        If PDImages.IsImageActive(i) Then
            GetNextImageID = i
            Exit Function
        End If
                
        If moveForward Then
            i = i + 1
        Else
            i = i - 1
        End If
                
    Loop
    
End Function

'Return the number of user images currently open (e.g. currently loaded to the main window).
Public Function GetNumOpenImages() As Long
    GetNumOpenImages = m_OpenImageCount
End Function

'Return the number of user images ever opened during this session (including images that have been unloaded).
' At present, this number is used to generate unique image IDs.
Public Function GetNumSessionImages() As Long
    GetNumSessionImages = m_ImagesLoadedThisSession
End Function

'When loading an image file, there's a chance the image won't load correctly (i.e. a bad file).  Because of that,
' we always start with a "provisional" ID value for a given image.  If the image fails to load, we can reuse the
' ID value on a subsequent image.
Public Function GetProvisionalImageID() As Long
    GetProvisionalImageID = m_ImagesLoadedThisSession
End Function

'Return the upper bound of the current image collection; IDs above this value are invalid.  (This can be helpful
' when externally traversing the active image collection.)
Public Function GetImageCollectionSize() As Long
    GetImageCollectionSize = UBound(m_PDImages)
End Function

'Is a valid image loaded AND active?  This is a slightly more comprehensive wrapper to IsImageAvailable(), below.
Public Function IsImageActive(Optional ByVal imgID As Long = -1) As Boolean
    IsImageActive = False
    If (imgID < 0) Then imgID = m_CurrentImageID
    If IsImageNonNull(imgID) Then IsImageActive = m_PDImages(imgID).IsActive()
End Function

'Is a valid image loaded?  (Basically, this returns TRUE if image operations can be safely applied to the
' image ID in question; if no ID is passed, the check will default to "the currently active image, if any")
Public Function IsImageNonNull(Optional ByVal imgID As Long = -1) As Boolean
    
    IsImageNonNull = False
    
    If (m_ImagesLoadedThisSession > 0) Then
        If (imgID < 0) Then imgID = m_CurrentImageID
        On Error GoTo NoImagesAvailable
        IsImageNonNull = (Not m_PDImages(imgID) Is Nothing)
    End If
    
    Exit Function
    
NoImagesAvailable:
    IsImageNonNull = False
End Function

'The "Next Image" and "Previous Image" options simply wrap this function.
Public Sub MoveToNextImage(ByVal moveForward As Boolean)

    'If one (or zero) images are loaded, ignore this option
    If (PDImages.GetNumOpenImages() < 2) Then Exit Sub
    
    Dim newIndex As Long
    newIndex = PDImages.GetNextImageID(PDImages.GetActiveImageID(), moveForward)
    If (newIndex >= 0) Then CanvasManager.ActivatePDImage newIndex, "user requested next/previous image"
    
End Sub

'Forcibly release all associated PDImage resources; this may include additional resources besides just the
' m_PDImages() array.
Public Sub ReleaseAllPDImageResources()
    
    On Error GoTo PDImageResourcesFreed
    
    Dim i As Long
    For i = LBound(m_PDImages) To UBound(m_PDImages)
        If (Not m_PDImages(i) Is Nothing) Then
            m_PDImages(i).FreeAllImageResources
            Set m_PDImages(i) = Nothing
        End If
    Next i
    
PDImageResourcesFreed:
    
    'Finish by resetting the image array size
    ResetPDImageCollection
    
End Sub

'Remove an image from the collection.  Note that the formal implementation of this is not set in stone; at present,
' for performance reasons, we free some resources related to the image but do not resize the actual image collection.
' Also note that this function does NOT modify the m_CurrentImageID value, by design; callers are left to implement that
' in whatever fashion they desire.
'RETURNS: TRUE if the image was removed successfully; FALSE otherwise
Public Function RemovePDImageFromCollection(ByVal imgID As Long) As Boolean

    If PDImages.IsImageNonNull(imgID) Then
    
        'Decrease the open image count
        m_OpenImageCount = m_OpenImageCount - 1
        
        'Deactivate any resources associated with the target image
        m_PDImages(imgID).FreeAllImageResources
    
    End If

End Function

'Call once at start-up; unused otherwise
Public Sub ResetPDImageCollection()
    ReDim m_PDImages(0 To 3) As pdImage
    m_ImagesLoadedThisSession = 0
    m_CurrentImageID = 0
    m_OpenImageCount = 0
End Sub

'Set a new active image ID.  Note that this does NOT control any UI-related functions; that is left to the caller.
Public Sub SetActiveImageID(ByVal newID As Long)
    m_CurrentImageID = newID
End Sub

'If memory usage is getting worrisome (as defined by the caller), call this function to strategically off-load
' other loaded images to disk.  This can free up a *lot* of memory if multiple images are loaded, but like any
' other disk-access-operation, it's gonna require some disk space and CPU time to do so.
'
'(After calling this function, you don't need to do anything else - PD automatically un-suspends data as
' it accesses it.)
Public Sub StrategicMemoryReduction()
    
    'I don't like to dump un-compressed data to disk (it's wasteful!) but compression requires a temporary
    ' buffer to hold compression results, because we don't know how much space it'll take.
    '
    'As such, it's helpful to start with the smallest image and then work our way up from there.
    ' This (should) provide the maximum likelihood of success.
    Dim numImagesToSuspend As Long
    numImagesToSuspend = PDImages.GetNumOpenImages() - 1
    
    'If only one image is open, there's not much else we can do (gasp).
    If (numImagesToSuspend > 0) Then
        
        'Some images may already be suspended.  See how many we *actually* need to suspend.
        Dim i As Long
        For i = 0 To PDImages.GetImageCollectionSize()
            If PDImages.IsImageActive(i) Then
                If PDImages.GetImageByID(i).IsSuspended() Then
                    numImagesToSuspend = numImagesToSuspend - 1
                End If
            End If
        Next i
        
        'Now we know how many images we *actually* need to suspend.  Proceed, and dump as many as we can
        ' out to disk.
        Dim numImagesSuspended As Long
        Do While (numImagesSuspended < numImagesToSuspend)
            
            'Find the smallest remaining unsuspended image.
            Dim idTarget As Long, smallestYet As Long
            idTarget = -1
            smallestYet = LONG_MAX
            
            For i = 0 To PDImages.GetImageCollectionSize()
                If PDImages.IsImageActive(i) Then
                    If (Not PDImages.GetImageByID(i).IsSuspended()) Then
                        If (PDImages.GetImageByID(i).EstimateRAMUsage() < smallestYet) Then
                            smallestYet = PDImages.GetImageByID(i).EstimateRAMUsage()
                            idTarget = i
                        End If
                    End If
                End If
            Next i
            
            'Offload that image to disk
            If (idTarget >= 0) Then
            
                'Failsafe only
                If (Not PDImages.GetImageByID(idTarget) Is Nothing) Then
                    PDDebug.LogAction "Strategically suspending image #" & idTarget & " to disk to free up memory..."
                    PDImages.GetImageByID(idTarget).SuspendImage True
                    numImagesSuspended = numImagesSuspended + 1
                End If
            
            'Failsafe only; should never trigger
            Else
                Exit Do
            End If
            
        Loop
        
    End If
    
End Sub
