Attribute VB_Name = "Filters_Edge"
'***************************************************************************
'Filter (Edge) Interface
'Copyright ©2000-2014 by Tanner Helland
'Created: 12/June/01
'Last updated: 05/September/12
'Last update: rewrote and optimized all filters against the new DIB class.
'
'Runs all edge-related filters (edge detection, relief, etc.).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Redraw the image using a pencil sketch effect.
Public Sub FilterPencil()
    
    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("pencil sketch") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & "1|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|-1|-1|0|0|"
    tmpString = tmpString & "0|-1|6|-1|0|"
    tmpString = tmpString & "0|0|-1|-1|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString

End Sub

'A typical relief filter, that makes the image seem pseudo-3D.
Public Sub FilterRelief()

    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("relief") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & "0|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "2|40|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|2|-1|0|0|"
    tmpString = tmpString & "0|1|1|-1|0|"
    tmpString = tmpString & "0|0|1|-2|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString

End Sub

'A lighter version of a traditional sharpen filter; it's designed to bring out edge detail without the blowout typical of sharpening
Public Sub FilterEdgeEnhance()

    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("edge enhance") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & "0|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "4|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|0|-1|0|0|"
    tmpString = tmpString & "0|-1|8|-1|0|"
    tmpString = tmpString & "0|0|-1|0|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString

End Sub
