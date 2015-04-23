VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdTextRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon GDI+ Text Manager and Renderer
'Copyright 2015-2015 by Tanner Helland
'Created: 22/April/15
'Last updated: 22/April/15
'Last update: initial build
'
'While PhotoDemon already has a capable GDI font renderer in pdFont, text rendering onto actual images is a whole other
' can of worms.
'
'The biggest problem with GDI rendering is its total lack of support for 32-bpp rendering targets.  This is a deal-breaker
' for text layers, so I have no choice but to lean on GDI+ (short of writing some incredibly ugly functions for rasterizing
' GDI font render results, and extrapolating font data - ugh no).
'
'As such, this text renderer class was originally designed against GDI+ text capabilities.  As a failsafe against future
' improvements, I have included some framework code for variable backends.  Pango intrigues me as another rendering option,
' so I may look to it in the future if I can't wrangle GDI+ into supporting all the features I'd like.  (As it is, I'll
' be manually implementing some features, like line and character spacing.)
'
'Dependencies:
' - pdStringStack (used to report font lists as detected by GDI+)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Font collection interfaces
Private Declare Function GdipNewInstalledFontCollection Lib "gdiplus" (ByRef dstFontCollectionHandle As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal srcFontCollection As Long, ByRef dstNumFound As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal srcFontCollection As Long, ByVal sizeOfDstBuffer As Long, ByVal ptrToDstFontFamilyArray As Long, ByRef dstNumFound As Long) As Long
Private Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal ptrDstNameBuffer As Long, ByVal languageID As Integer) As Long
Private Const LF_FACESIZE As Long = 32      'Note: 32 chars, not bytes!
Private Const LANG_NEUTRAL As Integer = &H0

'Helper function for returning a string stack of currently installed, GDI+ compatible (e.g. TrueType) fonts
Public Function getListOfInstalledFonts(ByRef dstStringStack As pdStringStack) As Boolean
    
    'Create a new GDI+ font collection object
    Dim fontCollection As Long
    If GdipNewInstalledFontCollection(fontCollection) = 0 Then
    
        'Get the family count
        Dim fontCount As Long
        If GdipGetFontCollectionFamilyCount(fontCollection, fontCount) = 0 Then
        
            'Prep a Long-type array, which will receive the list of fonts installed on this machine
            Dim fontList() As Long
            If fontCount > 0 Then ReDim fontList(0 To fontCount - 1) As Long Else ReDim fontList(0) As Long
        
            'I don't know if it's possible for GDI+ to return a different amount of fonts than it originally reported,
            ' but since it takes a parameter for numFound, let's use it
            Dim fontsFound As Long
            If GdipGetFontCollectionFamilyList(fontCollection, fontCount, VarPtr(fontList(0)), fontsFound) = 0 Then
            
                'Populate our string stack with the names of this collection; also, since we know the approximate size of
                ' the stack in advance, we can accurately prep the stack's buffer.
                If dstStringStack Is Nothing Then Set dstStringStack = New pdStringStack
                dstStringStack.resetStack fontCount
                
                'Retrieve all font names
                Dim i As Long, thisFontName As String
                For i = 0 To fontsFound - 1
                    
                    'Retrieve the name for this entry
                    thisFontName = String$(LF_FACESIZE, 0)
                    If GdipGetFamilyName(fontList(i), StrPtr(thisFontName), LANG_NEUTRAL) = 0 Then
                        dstStringStack.AddString TrimNull(thisFontName)
                    End If
                    
                Next i
                
                'Return success
                getListOfInstalledFonts = True
            
            Else
                Debug.Print "WARNING! GDI+ refused to return a font collection list."
                getListOfInstalledFonts = False
            End If
        
        Else
            Debug.Print "WARNING! GDI+ refused to return a font collection count."
            getListOfInstalledFonts = False
        End If
    
    Else
        Debug.Print "WARNING! GDI+ refused to return a font collection object."
        getListOfInstalledFonts = False
    End If
    
End Function
