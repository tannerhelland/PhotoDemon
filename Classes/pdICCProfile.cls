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
'Copyright ©2012-2013 by Tanner Helland
'Created: 01/November/13
'Last updated: 01/November/13
'Last update: initial build
'
'ICC Profiles can be embedded in certain types of images (JPEG, PNG, and TIFF at the time of this writing - RAW
' may be added at a later date, but it's a bit more complicated as it will need to rely on ExifTool).  These
' profiles can be used to convert an image to its true color space, taking into account any pecularities of the
' device that captured the image (typically a camera), and the device now being used to display the image
' (typically a monitor).
'
'ICC Profile handling is broken into two parts: extracting the profile from an image, then applying that profile
' to the image.  The extraction step is currently handled via FreeImage, while the application step is handled by
' Windows.  In the future I may look at adding ExifTool as a possibly mechanism for extracting the profile, as it
' provides better support for RAW-format photos than FreeImage does.
'
'Note that this class retains a copy of the ICC Profile itself, which is typically a ~1k byte array.  In the future
' I may add support for attaching a new profile to an image, or for exporting a profile to file.
'
'This class would not be possible without this excellent test code from pro VB coder LaVolpe:
' http://www.vbforums.com/showthread.php?666143-RESOLVED-ICC-%28Color-Profiles%29
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'If we are able to successfully load an ICC profile for this image, this will be set to TRUE.
Private iccDataPresent As Boolean

'The actual ICC Profile is stored in this array
Private iccData() As Byte

'CopyMemory is used to shuttle around ICC data in memory
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDestination As Any, ByRef lpSource As Any, ByVal Length As Long)

'ICC Profile header; this stores basic information about a given profile, and is use to interact with various
' ICC-related API functions.
Private Type ICC_PROFILE
    dwType As Long
    pProfileData As Long
    cbDataSize As Long
End Type

'There are two possible dwType values for an ICC_PROFILE variable; we use PROFILE_MEMBUFFER
Private Const PROFILE_FILENAME As Long = 1&
Private Const PROFILE_MEMBUFFER As Long = 2&

'We only need to read ICC data, not write it
Private Const PROFILE_READ As Long = 1&

'Other functions are welcome to share the profile data
Private Const FILE_SHARE_READ As Long = 1&

'We want the function to fail if the profile cannot be opened; do not simply create a blank profile in its place
Private Const OPEN_EXISTING As Long = 3&

'Thanks to coder LaVolpe for pointing out that this value is actually "sRGB" in ASCII (big-endian)
Private Const LCS_sRGB As Long = &H73524742

'Profile transformation is Heisenbergian (sp?); for example, it is not always possible to preserve hue, saturation,
' and luminance - some must be sacrificed to preserve the other.  By default, PhotoDemon uses the standard intent
' for displays, which is IntentPerceptual (basically, stretch the image's luminance so that its full gamut is viewable).
Private Const INTENT_PERCEPTUAL As Long = 0&
Private Const INTENT_RELATIVECOLORIMETRIC As Long = 1&
Private Const INTENT_SATURATION As Long = 2&
Private Const INTENT_ABSOLUTECOLORIMETRIC As Long = 3&

'Windows provides different qualities for profile transformations (proof, normal, best).  We use BEST.
Private Const BEST_MODE As Long = 3&

'Because we only do ICC-to-ICC transforms, Windows can be instructed to use a 3rd-party CMS instead of its own
' internal one.  We don't care if it does this, and we tell it as much.
Private Const INDEX_DONT_CARE As Long = 0&

'When it comes time to actually apply the transformation to the image data, the transform needs to know the image's
' color depth.  PD only operates in 24 and 32bpp mode, so we only need two constants here.
Private Const BM_BGRTRIPLETS As Long = 4&
Private Const BM_xBGRQUADS As Long = 16&

'Various ICC-related APIs are needed to open profiles and transform data between them
Private Declare Function OpenColorProfile Lib "mscms" Alias "OpenColorProfileA" (ByRef pProfile As Any, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal dwCreationMode As Long) As Long
Private Declare Function IsColorProfileValid Lib "mscms" (ByVal hProfile As Long, ByRef pBool As Long) As Long
Private Declare Function GetStandardColorSpaceProfile Lib "mscms" Alias "GetStandardColorSpaceProfileA" (ByVal pcstr As String, ByVal dwProfileID As Long, ByVal pProfileName As Long, ByRef pdwSize As Long) As Long
Private Declare Function CreateMultiProfileTransform Lib "mscms" (ByRef pProfile As Any, ByVal nProfiles As Long, ByRef pIntents As Long, ByVal nIntents As Long, ByVal dwFlags As Long, ByVal indexPreferredCMM As Long) As Long
Private Declare Function TranslateBitmapBits Lib "mscms" (ByVal hTransform As Long, ByRef srcBits As Any, ByVal pBmInput As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal dwInputStride As Long, ByRef dstBits As Any, ByVal pBmOutput As Long, ByVal dwOutputStride As Long, ByRef pfnCallback As Long, ByVal ulCallbackData As Long) As Long
Private Declare Function CloseColorProfile Lib "mscms" (ByVal hProfile As Long) As Long
Private Declare Function DeleteColorTransform Lib "mscms" (ByVal hTransform As Long) As Long

'If an ICC profile is present, apply it to the parent pdImage object.  The parent must be explicitly passed, as this
' small class does not maintain a persistent link on its own.
Public Function applyICCtoParentImage(ByRef parentPDImage As pdImage) As Boolean

    'Before doing anything else, make sure we actually have an ICC profile to apply!
    If Not iccDataPresent Then
        Message "ICC transform requested, but no data found.  Abandoning attempt."
        Exit Function
    End If
    
    Message "Using embedded ICC profile to convert image to sRGB space for editing..."

    'Start by preparing an ICC_PROFILE header to use with the color management APIs
    Dim srcProfileHeader As ICC_PROFILE
    srcProfileHeader.dwType = PROFILE_MEMBUFFER
    srcProfileHeader.pProfileData = VarPtr(iccData(0))
    srcProfileHeader.cbDataSize = UBound(iccData) + 1
    
    'Use that header to open a reference to an internal Windows color profile (which is required by all ICC-related API)
    Dim srcProfile As Long
    srcProfile = OpenColorProfile(srcProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
    
    'If we opened the profile successfully, continue on to the next step!
    If srcProfile <> 0 Then
    
        'Validate the XML schema defined by the profile
        Dim tmpCheck As Long
        If IsColorProfileValid(srcProfile, tmpCheck) <> 0 Then
        
            'Our source profile is loaded and valid.  Awesome!
            
            'Now it's time to determine our destination profile.  Because PhotoDemon operates on DIBs that default
            ' to the sRGB space, that's the profile we want to use for transformation.
            
            'Start by preparing a header for the destination ICC profile
            Dim dstProfileHeader As ICC_PROFILE
            dstProfileHeader.dwType = PROFILE_FILENAME
            
            'We do not know the size of the sRGB profile in advance, so we must use a specialized call to
            ' GetStandardColorSpaceProfile, which will fill the last parameter with the size of the profile.
            GetStandardColorSpaceProfile vbNullString, LCS_sRGB, 0&, dstProfileHeader.cbDataSize
            
            '.cbDataSize now contains the size of the required sRGB profile.  Prepare a dummy array to hold
            ' the data.
            Dim dstICCData() As Byte
            ReDim dstICCData(0 To dstProfileHeader.cbDataSize - 1) As Byte
            dstProfileHeader.pProfileData = VarPtr(dstICCData(0))
            
            'Now that we have an array to contain the profile, we use GetStandardColorSpaceProfile to fill it
            GetStandardColorSpaceProfile vbNullString, LCS_sRGB, dstProfileHeader.pProfileData, dstProfileHeader.cbDataSize
            
            'With a fully populated header, it is finally time to open an internal Windows version of the data!
            Dim dstProfile As Long
            dstProfile = OpenColorProfile(dstProfileHeader, PROFILE_READ, FILE_SHARE_READ, OPEN_EXISTING)
            
            'We could validate the sRGB header here, but since it came from the system, I assume it's valid.
            
            'Next we need to prepare a transformation matrix to supply to CreateMultiProfileTransform.
            ' The format of this matrix is a bit odd; for details, see http://msdn.microsoft.com/en-us/library/windows/desktop/dd316840%28v=vs.85%29.aspx
            Dim profileMatrix(0 To 1, 0 To 1) As Long
            
            'The first row in the array contains the two profile pointers we've already acquired, in src/dest order
            profileMatrix(0, 0) = srcProfile
            profileMatrix(1, 0) = dstProfile
            
            'The second column in the array contains the render intents for the transformation.
            profileMatrix(0, 1) = INTENT_PERCEPTUAL
            profileMatrix(1, 1) = INTENT_PERCEPTUAL
            
            'We can now use our profile matrix to generate a transformation object, which we will use on the DIB itself
            Dim iccTransformation As Long
            iccTransformation = CreateMultiProfileTransform(profileMatrix(0, 0), 2&, profileMatrix(0, 1), 2&, BEST_MODE, INDEX_DONT_CARE)
            
            'If the transformation was generated successfully, carry on!
            If iccTransformation <> 0 Then
            
                'Create a temporary copy of the image to use as the source bits
                Dim tmpLayer As pdLayer
                Set tmpLayer = New pdLayer
                tmpLayer.createFromExistingLayer parentPDImage.mainLayer
                
                'The only transformation function relevant to PD involves the use of BitmapBits, so we will provide
                ' the API with direct access to our DIB bits.
                
                'Note that a color format must be explicitly specified - we vary this contingent on the parent image's
                ' color depth.
                Dim transformCheck As Long
                With parentPDImage.mainLayer
                
                    Dim bitDepthIdentifier As Long
                    If .getLayerColorDepth = 24 Then bitDepthIdentifier = BM_BGRTRIPLETS Else bitDepthIdentifier = BM_xBGRQUADS
                    
                    transformCheck = TranslateBitmapBits(iccTransformation, ByVal tmpLayer.getLayerDIBits, bitDepthIdentifier, .getLayerWidth, .getLayerHeight, 0&, ByVal .getLayerDIBits, bitDepthIdentifier, 0&, ByVal 0&, 0&)
                    
                End With
                
                'If the transform was successful, pat ourselves on the back.
                If transformCheck <> 0 Then
                    Message "ICC profile transformation successful.  Image is now sRGB."
                Else
                    Message "ICC profile could not be applied.  Image remains in original profile."
                End If
                    
                'Release our transformation
                DeleteColorTransform iccTransformation
                Erase profileMatrix
                
                'Erase our temporary image
                Set tmpLayer = Nothing
                
            Else
                Message "Both ICC profiles loaded successfully, but transformation could not be created."
                applyICCtoParentImage = False
            End If
            
            CloseColorProfile dstProfile
        
        Else
            Message "Source ICC profile loaded successfully, but invalid XML found."
            applyICCtoParentImage = False
        End If
        
        CloseColorProfile srcProfile
    
    Else
        Message "Could not open attached ICC profile."
        applyICCtoParentImage = False
    End If
    
End Function

'At load-time, while a FreeImage handle is still active, use this function to extract an ICC profile from an image.
' This function assumes you have already checked to make sure the image contains an ICC profile (see the
' Plugin_FreeImage_Expanded module for details).
Public Function loadICCFromFreeImage(ByVal fi_Bitmap As Long) As Boolean

    Message "Image contains ICC profile.  Copying profile now..."
    
    'Check to make sure the supplied FreeImage bitmap actually has ICC data
    If FreeImage_GetICCProfileSize(fi_Bitmap) > 0 Then
    
        Dim fiProfileHeader As FIICCPROFILE
        fiProfileHeader = FreeImage_GetICCProfile(fi_Bitmap)
        
        'Use the SIZE value from the header to prepare a VB array to receive the profile data
        ReDim iccData(0 To fiProfileHeader.Size - 1) As Byte
        
        'Extract the ICC data into the array
        CopyMemory ByVal VarPtr(iccData(0)), ByVal fiProfileHeader.Data, fiProfileHeader.Size
        
        'The complete ICC profile will now be inside the iccData array.
        Message "ICC profile copied successfully."
        Debug.Print "ICC profile size (in bytes): " & CStr(UBound(iccData) + 1)
        iccDataPresent = True
        loadICCFromFreeImage = True
        
    Else
        Message "ICC profile could not be copied (unknown error)."
        loadICCFromFreeImage = False
    End If
    
End Function

Public Function hasICCData() As Boolean
    hasICCData = iccDataPresent
End Function

Private Sub Class_Initialize()
    iccDataPresent = False
End Sub
