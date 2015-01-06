VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdICCProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon ICC (International Color Consortium) Profile Manager
'Copyright 2013-2015 by Tanner Helland
'Created: 05/November/13
'Last updated: 31/January/14
'Last update: retrieve the rendering intent from embedded ICC profiles, and use that for the first stage of color
'              transforms (if available).
'
'ICC profiles can be embedded in certain types of images (JPEG, PNG, and TIFF at the time of this writing).  These
' profiles can be used to convert an image to its true color space, taking into account any pecularities of the
' device that captured the image (typically a camera), and the device now being used to display the image
' (typically a monitor).
'
'ICC profile handling is broken into two parts: extracting the profile from an image, then applying that profile
' to the image.  The extraction step is currently handled via FreeImage or GDI+, while the application step is handled
' by Windows.  In the future I may look at adding ExifTool as a possibly mechanism for extracting the profile, as it
' provides better support for esoteric formats than FreeImage.
'
'Note that this class retains a copy of the ICC Profile itself, which is typically a ~1-5k byte array.  In the future
' I may add support for attaching a new profile to an image, or for exporting a profile to file.
'
'Also note that this class interacts heavily with the Color_Management module.
'
'This class would not be possible without this excellent test code from pro VB coder LaVolpe:
' http://www.vbforums.com/showthread.php?666143-RESOLVED-ICC-%28Color-Profiles%29
' Note that LaVolpe's code contains a number of errors, so if you're looking to build your own ICC implementation,
' I suggest basing it off my implementation instead.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Once we have converted the image to the sRGB working space, this will be set to TRUE.
Private iccProfileHasBeenApplied As Boolean

'If we are able to successfully load an ICC profile for this image, this will be set to TRUE.
Private iccDataPresent As Boolean

'The actual ICC Profile is stored in this array
Private iccData() As Byte

'Retrieve the source render intent from the contained ICC profile
Public Function getSourceRenderIntent() As Long

    'Make sure we have been given an ICC profile to work with!
    If UBound(iccData) = 0 Then
        getSourceRenderIntent = -1
        Exit Function
    End If
    
    'Render intent is stored in bytes 63-67 of the ICC profile, per the spec.  (http://www.color.org/specification/ICC1v43_2010-12.pdf)
    ' Technically only byte 67 matters, as render intent can only be a value of 0, 1, 2, or 3.
    Dim srcIntent As Long
    If UBound(iccData) >= 67 Then
        srcIntent = iccData(67)
    Else
        srcIntent = 0
    End If
    
    'Validate the intent; it must fall between 0 and 3.  If it does not, set it to 0 (Perceptual)
    If srcIntent < 0 Or srcIntent > 3 Then srcIntent = 0
    
    getSourceRenderIntent = srcIntent
    
End Function

'If a profile has been successfully applied to the parent image, this will return TRUE.
Public Function hasProfileBeenApplied() As Boolean
    hasProfileBeenApplied = iccProfileHasBeenApplied
End Function

'If external functions (such as our new CMYK conversion code) apply a profile independently, they can use this function
' to mark the image as successfully converted.  This saves us the extra work of checking again in the LoadFileAsNewImage function.
Public Sub markSuccessfulProfileApplication()
    iccProfileHasBeenApplied = True
End Sub

'The external CMYK transform code performs its own special transform using the stored ICC profile; as such, it requires a
' pointer to the stored ICC profile data.
Public Function getICCDataPointer() As Long
    getICCDataPointer = VarPtr(iccData(0))
End Function

Public Function getICCDataSize() As Long
    getICCDataSize = UBound(iccData) + 1
End Function

'If an ICC profile is present, apply it to this DIB.
Public Function applyICCtoSelf(ByRef copyOfThisInstance As pdDIB) As Boolean

    'Before doing anything else, make sure we actually have an ICC profile to apply!
    If Not iccDataPresent Then
        Message "ICC transform requested, but no data found.  Abandoning attempt."
        Exit Function
    End If
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Using embedded ICC profile to convert image to sRGB space for editing..."
    #End If
    
    'Use the Color_Management module to convert the raw ICC profile into an internal Windows profile handle.  Note that
    ' this function will also validate the profile for us.
    Dim srcProfile As Long
    srcProfile = loadICCProfileFromMemory(VarPtr(iccData(0)), UBound(iccData) + 1)
    
    'If we successfully opened and validated our source profile, continue on to the next step!
    If srcProfile <> 0 Then
    
        'Now it is time to determine our destination profile.  Because PhotoDemon operates on DIBs that default
        ' to the sRGB space, that's the profile we want to use for transformation.
            
        'Use the Color_Management module to request a standard sRGB profile.
        Dim dstProfile As Long
        dstProfile = loadStandardICCProfile(LCS_sRGB)
        
        'It's highly unlikely that a request for a standard ICC profile will fail, but just be safe, double-check the
        ' returned handle before continuing.
        If dstProfile <> 0 Then
            
            'Before proceeding, check to see if the source and destination profiles are identical.  Some dSLRs will embed
            ' sRGB transforms in their JPEGs, and applying another sRGB transform atop them is a waste of time and resources.
            ' Thanks to VB developer LaVolpe for this suggestion.
            If Not Color_Management.areColorProfilesEqual(srcProfile, dstProfile) Then
            
                'We can now use our profile matrix to generate a transformation object, which we will use to directly modify
                ' the DIB's RGB values.
                Dim iccTransformation As Long
                iccTransformation = requestProfileTransform(srcProfile, dstProfile, INTENT_PERCEPTUAL, getSourceRenderIntent)
                
                'If the transformation was generated successfully, carry on!
                If iccTransformation <> 0 Then
                    
                    'The only transformation function relevant to PD involves the use of BitmapBits, so we will provide
                    ' the API with direct access to our DIB bits.
                    
                    'Note that a color format must be explicitly specified - we vary this contingent on the parent image's
                    ' color depth.
                    Dim transformCheck As Boolean
                    transformCheck = applyColorTransformToDIB(iccTransformation, copyOfThisInstance)
                    
                    'If the transform was successful, pat ourselves on the back.
                    If transformCheck Then
                    
                        #If DEBUGMODE = 1 Then
                            pdDebug.LogAction "ICC profile transformation successful.  Image is now sRGB."
                        #End If
                        
                        iccProfileHasBeenApplied = True
                        applyICCtoSelf = True
                    Else
                        Message "ICC profile could not be applied.  Image remains in original profile."
                    End If
                    
                    'Release our transformation
                    releaseColorTransform iccTransformation
                                    
                Else
                    Message "Both ICC profiles loaded successfully, but transformation could not be created."
                    applyICCtoSelf = False
                End If
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "ICC transform is not required, because source and destination profiles are identical."
                #End If
                applyICCtoSelf = True
            End If
        
            releaseICCProfile dstProfile
        
        Else
            Message "Could not obtain standard sRGB color profile.  Color management has been disabled for this image."
            applyICCtoSelf = False
        End If
        
        releaseICCProfile srcProfile
    
    Else
        Message "Embedded ICC profile is invalid.  Color management has been disabled for this image."
        applyICCtoSelf = False
    End If
    
End Function

'At load-time, while a FreeImage handle is still active, use this function to extract an ICC profile from an image.
' This function assumes you have already checked to make sure the image contains an ICC profile (see the
' Plugin_FreeImage_Expanded module for details).
Public Function loadICCFromFreeImage(ByVal fi_Bitmap As Long) As Boolean

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Image contains ICC profile.  Copying profile now..."
    #End If
    
    'Check to make sure the supplied FreeImage bitmap actually has ICC data
    If FreeImage_GetICCProfileSize(fi_Bitmap) > 0 Then
    
        Dim fiProfileHeader As FIICCPROFILE
        fiProfileHeader = FreeImage_GetICCProfile(fi_Bitmap)
        
        'Use the SIZE value from the header to prepare a VB array to receive the profile data
        ReDim iccData(0 To fiProfileHeader.Size - 1) As Byte
        
        'Extract the ICC data into the array
        CopyMemory ByVal VarPtr(iccData(0)), ByVal fiProfileHeader.Data, fiProfileHeader.Size
        
        'The complete ICC profile will now be inside the iccData array.
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ICC profile copied successfully."
        #End If
            
        iccDataPresent = True
        loadICCFromFreeImage = True
        
    Else
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ICC profile could not be copied (unknown error)."
        #End If
        
        loadICCFromFreeImage = False
        
    End If
    
End Function

Public Function loadICCFromGDIPlus(ByVal profileSize As Long, ByVal ptrToProfile As Long) As Boolean

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Image contains ICC profile.  Copying profile now..."
    #End If
    
    'Use the SIZE parameter to prepare a VB array to receive the profile data
    ReDim iccData(0 To profileSize - 1) As Byte
    
    'Extract the ICC data using the supplied pointer
    CopyMemory ByVal VarPtr(iccData(0)), ByVal ptrToProfile, profileSize
        
    'The complete ICC profile will now be inside the iccData array.
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "ICC profile copied successfully."
    #End If
    
    iccDataPresent = True
    loadICCFromGDIPlus = True

End Function

Public Function hasICCData() As Boolean
    hasICCData = iccDataPresent
End Function

Private Sub Class_Initialize()
    
    iccDataPresent = False
    iccProfileHasBeenApplied = False
    
    ReDim iccData(0) As Byte
    
End Sub

