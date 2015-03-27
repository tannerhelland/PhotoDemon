Attribute VB_Name = "DIB_Handler"
'***************************************************************************
'DIB Support Functions
'Copyright 2012-2015 by Tanner Helland
'Created: 27/March/15
'Last updated: 27/March/16
'Last update: start migrating rare functions out of pdDIB and into this module.
'
'This module contains support functions for the pdDIB class.  In old versions of PD, these functions were provided by pdDIB,
' but there's no sense cluttering up that class with functions that are only used on rare occasions.  As such, I'm moving
' as many of those functions as I can to this module.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Check to see if a 32bpp DIB is really 32bpp. (Basically, scan all pixels in the alpha channel. If all values are set to
' 255 or all values are set to 0, the caller can opt to rebuild the DIB in 24bpp mode.)
Public Function verifyDIBAlphaChannel(ByRef srcDIB As pdDIB) As Boolean

    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If srcDIB.getDIBColorDepth <> 32 Then
        verifyDIBAlphaChannel = True
        Exit Function
    End If
    
    'This routine will fail if the width or height of a DIB is 0
    If srcDIB.getDIBWidth = 0 Or srcDIB.getDIBHeight = 0 Then
        verifyDIBAlphaChannel = True
        Exit Function
    End If

    'Start, as always, with a SafeArray
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    Dim checkAlpha As Boolean, initAlpha As Double
    checkAlpha = False
    
    'Determine the alpha value of the top-left pixel. This will be used as our baseline value.
    initAlpha = iData(3, 0)
    
    'If initAlpha is something other than 255 or 0, we don't need to check the image
    If (initAlpha <> 0) And (initAlpha <> 255) Then
        
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        Erase iData
        
        verifyDIBAlphaChannel = True
        Exit Function
        
    End If
        
    'Loop through the image, comparing colors as we go
    For x = 0 To srcDIB.getDIBWidth - 1
        QuickX = x * 4
    For y = 0 To srcDIB.getDIBHeight - 1
        
        'Compare the alpha data for this pixel to the initial pixel. If they DO NOT match, this is a valid alpha channel.
        If initAlpha <> iData(QuickX + 3, y) Then
            checkAlpha = True
            Exit For
        End If
        
    Next y
    
        'If the alpha channel has been verified, exit this loop
        If checkAlpha Then Exit For
        
    Next x
    
    'With our check complete, point iData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData

    'Return checkAlpha. If varying alpha values were found, this function returns TRUE. If all values were the same,
    ' this function returns FALSE.
    verifyDIBAlphaChannel = checkAlpha

End Function

'Does a given DIB have "binary" transparency, e.g. does it have alpha values of only 0 or 255?
' (This is used to determine how transparency is handled when exporting to file formats like GIF, which do not support variable alpha.)
Public Function isDIBAlphaBinary(ByRef srcDIB As pdDIB, Optional ByVal checkForZero As Boolean = True) As Boolean

    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function

    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If srcDIB.getDIBColorDepth = 32 Then

        'Make sure this DIB isn't empty
        If (srcDIB.getDIBDC <> 0) And (srcDIB.getDIBWidth <> 0) And (srcDIB.getDIBHeight <> 0) Then
    
            'Loop through the image and compare each alpha value against 0 and 255. If a value doesn't
            ' match either of this, this is a non-binary alpha channel and it must be handled specially.
            Dim iData() As Byte
            Dim tmpSA As SAFEARRAY2D
            prepSafeArray tmpSA, srcDIB
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
            Dim x As Long, y As Long, QuickX As Long
                
            'By default, assume that the image does not have a binary alpha channel. (This is the preferable
            ' default, as we will exit the loop IFF a non-0 or non-255 value is found.)
            Dim notBinary As Boolean
            notBinary = False
            
            Dim chkAlpha As Byte
                
            'Loop through the image, checking alphas as we go
            For x = 0 To srcDIB.getDIBWidth - 1
                QuickX = x * 4
            For y = 0 To srcDIB.getDIBHeight - 1
            
                'Retrieve the alpha value of the current pixel
                chkAlpha = iData(QuickX + 3, y)
                
                'For optimization reasons, this is stated as two IFs instead of an OR.
                If chkAlpha <> 255 Then
                
                    If checkForZero Then
                    
                        If chkAlpha <> 0 Then
                            notBinary = True
                            Exit For
                        End If
                        
                    Else
                        notBinary = True
                        Exit For
                    End If
                
                End If
                
            Next y
                If notBinary Then Exit For
            Next x
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            
            Erase iData
                
            'Exit
            isDIBAlphaBinary = Not notBinary
            Exit Function
                
        End If
        
    End If
    
    'If we made it to this line, something went horribly wrong or the user used this function incorrectly
    ' (e.g. calling it on a 24-bpp DIB).
    isDIBAlphaBinary = False

End Function

'Is a given DIB grayscale?  Determination is made by scanning each pixel and comparing RGB values to see if they match.
Public Function isDIBGrayscale(ByRef srcDIB As pdDIB) As Boolean
    
    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function
    
    'Make sure this DIB isn't empty
    If (srcDIB.getDIBDC <> 0) And (srcDIB.getDIBWidth <> 0) And (srcDIB.getDIBHeight <> 0) Then
    
        'Loop through the image and compare RGB values to determine grayscale or not.
        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
        
        Dim x As Long, y As Long, QuickX As Long
                        
        Dim r As Long, g As Long, b As Long
        
        Dim qvDepth As Long
        qvDepth = srcDIB.getDIBColorDepth \ 8
                        
        'Loop through the image, checking alphas as we go
        For x = 0 To srcDIB.getDIBWidth - 1
            QuickX = x * qvDepth
        For y = 0 To srcDIB.getDIBHeight - 1
            
            r = iData(QuickX + 2, y)
            g = iData(QuickX + 1, y)
            b = iData(QuickX, y)
            
            'For optimization reasons, this is stated as multiple IFs instead of an OR.
            If r <> g Then
                
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
        
                isDIBGrayscale = False
                Exit Function
                
            ElseIf g <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                isDIBGrayscale = False
                Exit Function
                
            ElseIf r <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                isDIBGrayscale = False
                Exit Function
                
            End If
                
        Next y
        Next x
    
        'With our alpha channel complete, point iData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        Erase iData
                        
        'If we scanned all pixels without exiting prematurely, the DIB is grayscale
        isDIBGrayscale = True
        Exit Function
        
    End If
    
    'If we made it to this line, the DIB is blank, so it doesn't matter what value we return
    isDIBGrayscale = False

End Function

