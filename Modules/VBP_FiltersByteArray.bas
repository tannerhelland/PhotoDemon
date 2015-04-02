Attribute VB_Name = "Filters_ByteArray"
'***************************************************************************
'Byte Array Filters Module
'Copyright 2014-2015 by Tanner Helland
'Created: 02/April/15
'Last updated: 02/April/15
'Last update: start assembling byte-array-specific filter collection
'
'After version 6.6 released, I started work on a number of modified PD filters.  Unlike most filters in the project
' - which explicitly operate on pdDIB objects - these filters are modified to run on single-channel 2D byte arrays.
' In some places in the project, byte arrays contain sufficient detail for things like bumpmaps, and they consume
' less memory and are faster to process than a three- or four-channel DIB.
'
'Going forward, it would be nice to further expand this function collection, but in the meantime, just know that
' the functions in this module cannot directly operate on images.  Also, they do not generally support progress
' bar reports (by design) as their emphasis is on raw performance.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Gaussian blur filter, using an IIR (Infininte Impulse Response) approach
'
'I developed this function with help from http://www.getreuer.info/home/gaussianiir
' Many thanks to Pascal Getreuer for his valuable reference.
'
'IIR provides many benefits over a naive Gaussian Blur implementation:
' - It's performed in-place, meaning a second array is not required.  (That said, the function requires floating-point data,
'    so an intermediate float-type array *is* currently needed.)
' - It approaches a true Gaussian over multiple iterations, but at low iterations, it provides a closer estimate than the
'    corresponding box blur filter would.
' - Floating-point radii are supported.
' - Most importantly, it's much faster to calculate!
'
'Note that the incoming arrayWidth and arrayHeight parameters are 1-based, so this function will automatically subtract 1 to arrive
' at an actual UBound value.
Public Function GaussianBlur_IIR_ByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long, ByVal radius As Double, ByVal numSteps As Long) As Boolean
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here.
    ' (In the future, it might be nice to allow the caller to specify blur lBounds, but for now, 0 is assumed.)
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim iWidth As Long, iHeight As Long
    iWidth = arrayWidth
    iHeight = arrayHeight
    
    'Prep some IIR-specific values next
    Dim g As Long
    Dim lambda As Double, dnu As Double
    Dim nu As Double, boundaryScale As Double, postScale As Double
    Dim i As Long, step As Long
    
    'Calculate sigma from the radius, using the same formula we do for PD's pure gaussian blur
    Dim sigma As Double
    sigma = Sqr(-(radius * radius) / (2 * Log(1# / 255#)))
    
    'Another possible sigma formula, per this link (http://stackoverflow.com/questions/21984405/relation-between-sigma-and-radius-on-the-gaussian-blur):
    'sigma = (radius + 1) / Sqr(2 * (Log(255) / Log(10)))
    
    'Make sure sigma and steps are valid
    If sigma <= 0 Then sigma = 0.01
    If numSteps <= 0 Then numSteps = 1
    
    'In the best paper I've read on this topic (http://dx.doi.org/10.5201/ipol.2013.87), an alternate lambda calculation
    ' is proposed.  This adjustment doesn't affect running time at all, and should reduce errors relative to a pure Gaussian,
    ' allowing for better results with fewer iterations.
    
    'This behavior could theroetically be toggled by the caller, but for now, I've hard-coded use of the modified formula.
    Dim useModifiedQ As Boolean, q As Single
    useModifiedQ = True
    
    If useModifiedQ Then
        q = sigma * (1# + (0.3165 * numSteps + 0.5695) / ((numSteps + 0.7818) * (numSteps + 0.7818)))
    Else
        q = sigma
    End If
    
    'Calculate IIR values
    lambda = (q * q) / (2 * numSteps)
    dnu = (1 + 2 * lambda - Sqr(1 + 4 * lambda)) / (2 * lambda)
    nu = dnu
    boundaryScale = (1 / (1 - dnu))
    postScale = ((dnu / lambda) ^ (2 * numSteps)) * 255
    
    'Intermediate float arrays are required for an IIR transform.
    Dim gFloat() As Single
    ReDim gFloat(initX To finalX, initY To finalY) As Single
    
    'Copy the contents of the incoming byte array into the float array
    For x = initX To finalX
    For y = initY To finalY
        g = srcArray(x, y)
        gFloat(x, y) = g / 255
    Next y
    Next x
    
    'Filter horizontally along each row
    For y = initY To finalY
    
        For step = 0 To numSteps - 1
            
            'Set initial values
            gFloat(initX, y) = gFloat(initX, y) * boundaryScale
            
            'Filter right
            For x = initX + 1 To finalX
                gFloat(x, y) = gFloat(x, y) + nu * gFloat(x - 1, y)
            Next x
            
            'Fix closing row
            gFloat(finalX, y) = gFloat(finalX, y) * boundaryScale
            
            'Filter left
            For x = finalX To 1 Step -1
                gFloat(x - 1, y) = gFloat(x - 1, y) + nu * gFloat(x, y)
            Next x
            
        Next step
        
    Next y
    
    'Now repeat all the above steps, but filtering vertically along each column, instead
    For x = initX To finalX
        
        For step = 0 To numSteps - 1
            
            'Set initial values
            gFloat(x, initY) = gFloat(x, initY) * boundaryScale
            
            'Filter down
            For y = initY + 1 To finalY
                gFloat(x, y) = gFloat(x, y) + nu * gFloat(x, y - 1)
            Next y
                
            'Fix closing column values
            gFloat(x, finalY) = gFloat(x, finalY) * boundaryScale
                
            'Filter up
            For y = finalY To 1 Step -1
                gFloat(x, y - 1) = gFloat(x, y - 1) + nu * gFloat(x, y)
            Next y
            
        Next step
        
    Next x
    
    'Apply final post-scaling
    For x = initX To finalX
    For y = initY To finalY
    
        'Apply post-scaling and perform failsafe clipping.  (Shouldn't technically be necessary, but better safe than sorry.)
        g = gFloat(x, y) * postScale
        If g > 255 Then g = 255
        
        'Store the final value back into the source array
        srcArray(x, y) = g
        
    Next y
    Next x
    
    GaussianBlur_IIR_ByteArray = True
    
End Function

'Given a 2D byte array, normalize the contents to guarantee a full stretch on the range [0, 255]
Public Function normalizeByteArray(ByRef srcArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long) As Boolean
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = arrayWidth - 1
    finalY = arrayHeight - 1
    
    Dim g As Long, minVal As Long, maxVal As Long
    minVal = 255
    maxVal = 0
    
    'Start by finding max/min values in the current array
    For x = initX To finalX
    For y = initY To finalY
                
        g = srcArray(x, y)
        If g < minVal Then
            minVal = g
        ElseIf g > maxVal Then
            maxVal = g
        End If
    
    Next y
    Next x
        
    'If the data isn't normalized, normalize it now
    If (minVal > 0) Or (maxVal < 255) Then
        
        Dim curRange As Long
        curRange = maxVal - minVal
        
        'Build a normalization lookup table
        Dim normalizedLookup() As Byte
        ReDim normalizedLookup(0 To 255) As Byte
        
        Dim newValue As Long
        
        For x = 0 To 255
        
            newValue = (CDbl(x - minVal) / curRange) * 255
            
            If newValue < 0 Then
                newValue = 0
            ElseIf newValue > 255 Then
                newValue = 255
            End If
            
            normalizedLookup(x) = newValue
            
        Next x
            
        'Apply normalization
        For x = initX To finalX
        For y = initY To finalY
            srcArray(x, y) = normalizedLookup(srcArray(x, y))
        Next y
        Next x
    
    End If
    
    normalizeByteArray = True
    
End Function

