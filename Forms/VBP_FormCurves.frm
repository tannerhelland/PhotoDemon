VERSION 5.00
Begin VB.Form FormCurves 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Curves"
   ClientHeight    =   9615
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   15135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   641
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1009
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset curve"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   8160
      Width           =   5520
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5160
      Left            =   6000
      ScaleHeight     =   342
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   462
      TabIndex        =   4
      Top             =   120
      Width           =   6960
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   12120
      TabIndex        =   0
      Top             =   9030
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   13590
      TabIndex        =   1
      Top             =   9030
      Width           =   1365
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "other options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   7680
      Width           =   1500
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -360
      TabIndex        =   2
      Top             =   8880
      Width           =   16095
   End
End
Attribute VB_Name = "FormCurves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Curves Adjustment Dialog
'Copyright ©2008-2013 by Tanner Helland
'Created: sometime 2008
'Last updated: 11/July/13
'Last update: merge the code from a standalone project into PhotoDemon
'
'Standard luminosity adjustment via curves.  This dialog is based heavily on similar tools in other photo editors.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Control points for the live preview box
Private Type fPoint
    pX As Double
    pY As Double
End Type

'This array will store new points added by the user
Private numOfPoints As Long
Private cPoints() As fPoint

'Track mouse status between MouseDown and MouseMove events
Private isMouseDown As Boolean

'Currently selected node in the workspace area
Private selectedPoint As Long

'How close to a node the user must click to select that node
Private Const mouseAccuracy As Byte = 6

'Four arrays are needed to generate the cubic spline used for the curve function
Private iX() As Double
Private iY() As Double
Private p() As Double
Private u() As Double

'The final curve is used to fill this array, which will contain the actual spline points for each location
' in the spline.  It will be dynamically resized to match the width of the curves picture box.
Private cResults() As Double

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()

    Me.Visible = False
    
    Dim paramString As String
    paramString = ""

    'Top-left
    'paramString = (m_nPoints(0).pX - m_oPoints(0).pX) * (iWidth / m_previewWidth)
    'paramString = paramString & "|" & (m_nPoints(0).pY - m_oPoints(0).pY) * (iHeight / m_previewHeight)
    
    'Top-right
    'paramString = paramString & "|" & (iWidth + ((m_nPoints(1).pX - m_oPoints(1).pX) * (iWidth / m_previewWidth)))
    'paramString = paramString & "|" & (m_nPoints(1).pY - m_oPoints(1).pY) * (iHeight / m_previewHeight)
    
    'Bottom-right
    'paramString = paramString & "|" & (iWidth + ((m_nPoints(2).pX - m_oPoints(2).pX) * (iWidth / m_previewWidth)))
    'paramString = paramString & "|" & (iHeight + (m_nPoints(2).pY - m_oPoints(2).pY) * (iHeight / m_previewHeight))
    
    'Bottom-left
    'paramString = paramString & "|" & ((m_nPoints(3).pX - m_oPoints(3).pX) * (iWidth / m_previewWidth))
    'paramString = paramString & "|" & (iHeight + (m_nPoints(3).pY - m_oPoints(3).pY) * (iHeight / m_previewHeight))
    
    'Edge handling
    'paramString = paramString & "|" & CLng(cmbEdges.ListIndex)
    
    'Resampling
    'paramString = paramString & "|" & OptInterpolate(0).Value
    
    'Based on the user's selection, submit the proper processor request
    Process "Perspective", , paramString
    
    Unload Me
    
End Sub

'Apply a curve to an image's luminance values
' Input: a list of 256 values, one for each luminance point in the image
Public Sub ApplyCurveToImage(ByVal listOfPoints As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If toPreview = False Then Message "Applying new curve to image luminance..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'Parse the incoming parameter string into individual (x, y) pairs
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(listOfPoints) > 0 Then cParams.setParamString listOfPoints
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, cParams.GetLong(9), cParams.GetBool(10), curLayerValues.maxX, curLayerValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
            
    'Store region width and height as floating-point
    Dim imgWidth As Double, imgHeight As Double
    imgWidth = finalX - initX
    imgHeight = finalY - initY
    
    'If this is a preview, we need to adjust the width and height values to match the size of the preview box
    Dim wModifier As Double, hModifier As Double
    wModifier = 1
    hModifier = 1
    
    'Scale quad coordinates to the size of the image
    Dim invWidth As Double, invHeight As Double
    invWidth = 1 / imgWidth
    invHeight = 1 / imgHeight
    
    'Copy the points given by the user (which are currently strings) into individual floating-point variables
    Dim x0 As Double, x1 As Double, x2 As Double, x3 As Double
    Dim y0 As Double, y1 As Double, y2 As Double, y3 As Double
    
    x0 = cParams.GetDouble(1)
    y0 = cParams.GetDouble(2)
    x1 = cParams.GetDouble(3)
    y1 = cParams.GetDouble(4)
    x2 = cParams.GetDouble(5)
    y2 = cParams.GetDouble(6)
    x3 = cParams.GetDouble(7)
    y3 = cParams.GetDouble(8)
    
    'Message x0 & "," & y0 & " | " & x1 & "," & y1 & " | " & x2 & "," & y2 & " | " & x3 & "," & y3
    
    If toPreview Then
        x0 = x0 * wModifier
        y0 = y0 * hModifier
        x1 = x1 * wModifier
        y1 = y1 * hModifier
        x2 = x2 * wModifier
        y2 = y2 * hModifier
        x3 = x3 * wModifier
        y3 = y3 * hModifier
    End If
    
    x0 = x0 * invWidth
    y0 = y0 * invHeight
    x1 = x1 * invWidth
    y1 = y1 * invHeight
    x2 = x2 * invWidth
    y2 = y2 * invHeight
    x3 = x3 * invWidth
    y3 = y3 * invHeight
    
    'First things first: we need to map the original image (in terms of the unit square)
    ' to the arbitrary quadrilateral defined by the user's parameters
    Dim dx1 As Double, dy1 As Double, dx2 As Double, dy2 As Double, dx3 As Double, dy3 As Double
    dx1 = x1 - x2
    dy1 = y1 - y2
    dx2 = x3 - x2
    dy2 = y3 - y2
    dx3 = x0 - x1 + x2 - x3
    dy3 = y0 - y1 + y2 - y3
    
    'Technically, these are points in a matrix - and they could be defined as an array.  But VB accesses
    ' individual data types more quickly than an array, so we declare them each separately.
    Dim h11 As Double, h21 As Double, h31 As Double
    Dim h12 As Double, h22 As Double, h32 As Double
    Dim h13 As Double, h23 As Double, h33 As Double
    
    'Certain values can lead to divide-by-zero problems - check those in advance and convert 0 to something like 0.000001
    Dim chkDenom As Double
    
    chkDenom = (dx1 * dy2 - dy1 * dx2)
    If chkDenom = 0 Then chkDenom = 0.000000001
    
    h13 = (dx3 * dy2 - dx2 * dy3) / chkDenom
    h23 = (dx1 * dy3 - dy1 * dx3) / chkDenom
    h11 = x1 - x0 + h13 * x1
    h21 = x3 - x0 + h23 * x3
    h31 = x0
    h12 = y1 - y0 + h13 * y1
    h22 = y3 - y0 + h23 * y3
    h32 = y0
    h33 = 1
    
    'Next, we need to calculate the key set of transformation parameters, using the reverse-map data we just generated.
    ' Again, these are technically just matrix entries, but we get better performance by declaring them individually.
    Dim hA As Double, hB As Double, hC As Double, hD As Double, hE As Double, hF As Double, hG As Double, hH As Double, hI As Double
    
    hA = h22 * h33 - h32 * h23
    hB = h31 * h23 - h21 * h33
    hC = h21 * h32 - h31 * h22
    hD = h32 * h13 - h12 * h33
    hE = h11 * h33 - h31 * h13
    hF = h31 * h12 - h11 * h32
    hG = h12 * h23 - h22 * h13
    hH = h21 * h13 - h11 * h23
    hI = h11 * h22 - h21 * h12
        
    'Message vA & "," & vB & "," & vC & "," & vD & "," & vE & "," & vF & "," & VG & "," & vH & "," & vI
        
    'Scale those values to match the size of the transformed image
    hA = hA * invWidth
    hD = hD * invWidth
    hG = hG * invWidth
    hB = hB * invHeight
    hE = hE * invHeight
    hH = hH * invHeight
            
    'With all that data calculated in advanced, the actual transform is quite simple.
            
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
                
        'Reverse-map the coordinates back onto the original image (to allow for resampling)
        chkDenom = (hG * x + hH * y + hI)
        If chkDenom = 0 Then chkDenom = 0.000000001
        
        srcX = imgWidth * (hA * x + hB * y + hC) / chkDenom
        srcY = imgHeight * (hD * x + hE * y + hF) / chkDenom
                
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

Private Sub cmdReset_Click()

    resetCurvePoints
    redrawPreviewBox
    updatePreview

End Sub

Private Sub Form_Activate()
        
    'Mark the mouse as not being down
    isMouseDown = False
    
    'Reset the curve
    resetCurvePoints
    
    'Redraw the curve box
    redrawPreviewBox
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
        
    'Create the preview
    'updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()

    Dim paramString As String
    paramString = ""

    'Top-left
    'paramString = (m_nPoints(0).pX - m_oPoints(0).pX) * (iWidth / m_previewWidth)
    'paramString = paramString & "|" & (m_nPoints(0).pY - m_oPoints(0).pY) * (iHeight / m_previewHeight)
    
    'Top-right
    'paramString = paramString & "|" & (iWidth + ((m_nPoints(1).pX - m_oPoints(1).pX) * (iWidth / m_previewWidth)))
    'paramString = paramString & "|" & (m_nPoints(1).pY - m_oPoints(1).pY) * (iHeight / m_previewHeight)
    
    'Bottom-right
    'paramString = paramString & "|" & (iWidth + ((m_nPoints(2).pX - m_oPoints(2).pX) * (iWidth / m_previewWidth)))
    'paramString = paramString & "|" & (iHeight + (m_nPoints(2).pY - m_oPoints(2).pY) * (iHeight / m_previewHeight))
    
    'Bottom-left
    'paramString = paramString & "|" & ((m_nPoints(3).pX - m_oPoints(3).pX) * (iWidth / m_previewWidth))
    'paramString = paramString & "|" & (iHeight + (m_nPoints(3).pY - m_oPoints(3).pY) * (iHeight / m_previewHeight))
    
    'Edge handling
    'paramString = paramString & "|" & CLng(cmbEdges.ListIndex)
    
    'Resampling
    'paramString = paramString & "|" & OptInterpolate(0).Value
    
    'PerspectiveImage paramString, True, fxPreview
    
End Sub

Private Sub redrawPreviewBox()

    picDraw.Picture = LoadPicture("")
    
    'Start by drawing a grid through the center of the image
    picDraw.DrawWidth = 1
    picDraw.ForeColor = RGB(160, 160, 160)
    picDraw.Line (0, picDraw.Height / 2)-(picDraw.Width, picDraw.Height / 2)
    picDraw.Line (picDraw.Width / 2, 0)-(picDraw.Width / 2, picDraw.Height)
    
    'Next, generate a list of points that correspond to the cubic spline used for the curve
    fillResultsArray
    
    'Use the newly created results array to draw the cubic spline onto picDraw, while using GDI+ for antialiasing
    Dim i As Long
    For i = 1 To picDraw.ScaleWidth
        GDIPlusDrawLineToDC picDraw.hDC, i, cResults(i), i - 1, cResults(i - 1) + 2, RGB(0, 0, 255)
    Next i
    
    picDraw.Picture = picDraw.Image
    
    Exit Sub
    
    'Next, draw connecting lines to form an image outline.  Use GDI+ for superior results (e.g. antialiasing).
    Dim oTransparency As Long
    oTransparency = 220
    
    picDraw.ForeColor = RGB(255, 0, 0)
    For i = 0 To 3
        If i < 3 Then
            'GDIPlusDrawLineToDC picDraw.hDC, m_nPoints(i).pX, m_nPoints(i).pY, m_nPoints(i + 1).pX, m_nPoints(i + 1).pY, picDraw.ForeColor, oTransparency
        Else
            'GDIPlusDrawLineToDC picDraw.hDC, m_nPoints(i).pX, m_nPoints(i).pY, m_nPoints(0).pX, m_nPoints(0).pY, picDraw.ForeColor, oTransparency
        End If
    Next i
    
    'Next, draw circles at the corners of the perspective area
    picDraw.ForeColor = RGB(0, 0, 255)
    
    For i = 0 To 3
        'GDIPlusDrawCircleToDC picDraw.hDC, m_nPoints(i).pX, m_nPoints(i).pY, 5, picDraw.ForeColor, oTransparency
    Next i
    
    'Finally, draw the center cross to help the user orient to the center point of the perspective effect
    'GDIPlusDrawLineToDC picDraw.hDC, m_nPoints(0).pX, m_nPoints(0).pY, m_nPoints(2).pX, m_nPoints(2).pY, picDraw.ForeColor, oTransparency
    'GDIPlusDrawLineToDC picDraw.hDC, m_nPoints(1).pX, m_nPoints(1).pY, m_nPoints(3).pX, m_nPoints(3).pY, picDraw.ForeColor, oTransparency
    
    picDraw.Refresh

End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    isMouseDown = True
    
    'If the mouse is over a point, mark it as the active point
    'm_selPoint = checkClick(x, y)
    
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If the mouse is not down, indicate to the user that points can be moved
    If Not isMouseDown Then
        
        'If the user is close to a knot, change the mousepointer to 'move'
        If checkClick(x, y) > -1 Then
            If picDraw.MousePointer <> 5 Then picDraw.MousePointer = 5
            
            Select Case checkClick(x, y)
                Case 0
                    picDraw.ToolTipText = g_Language.TranslateMessage("top-left")
                Case 1
                    picDraw.ToolTipText = g_Language.TranslateMessage("top-right")
                Case 2
                    picDraw.ToolTipText = g_Language.TranslateMessage("bottom-right")
                Case 3
                    picDraw.ToolTipText = g_Language.TranslateMessage("bottom-left")
                    
            End Select
            
        Else
            If picDraw.MousePointer <> 0 Then picDraw.MousePointer = 0
        End If
    
    'If the mouse is down, move the current point and redraw the preview
    Else
    
        'If m_selPoint >= 0 Then
        '    m_nPoints(m_selPoint).pX = x
        '    m_nPoints(m_selPoint).pY = y
            redrawPreviewBox
            updatePreview
        'End If
    
    End If

End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    isMouseDown = False
    'm_selPoint = -1
End Sub

'Simple distance routine to see if a location on the picture box is near an existing point
Private Function checkClick(ByVal x As Long, ByVal y As Long) As Long
    Dim dist As Double
    Dim i As Long
    For i = 0 To 3
        'dist = pDistance(x, y, m_nPoints(i).pX, m_nPoints(i).pY)
        'If we're close to an existing point, return the index of that point
        If dist < mouseAccuracy Then
            checkClick = i
            Exit Function
        End If
    Next i
    'Returning -1 says we're not close to an existing point
    checkClick = -1
End Function

'Simple distance formula here - we use this to calculate if the user has clicked on (or near) a point
Private Function pDistance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Double
    pDistance = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function

'Original required spline function:
Private Function getCurvePoint(ByVal i As Long, ByVal v As Single) As Single
    Dim t As Single
    'derived curve equation (which uses p's and u's for coefficients)
    t = (v - iX(i)) / u(i)
    getCurvePoint = t * iY(i + 1) + (1 - t) * iY(i) + u(i) * u(i) * (F(t) * p(i + 1) + F(1 - t) * p(i)) / 6#
End Function

'Original required spline function:
Private Function F(x As Single) As Single
        F = x * x * x - x
End Function

'Original required spline function:
Private Sub SetPandU()
    Dim i As Integer
    Dim d() As Single
    Dim w() As Single
    ReDim d(numOfPoints) As Single
    ReDim w(numOfPoints) As Single
'Routine to compute the parameters of our cubic spline.  Based on equations derived from some basic facts...
'Each segment must be a cubic polynomial.  Curve segments must have equal first and second derivatives
'at knots they share.  General algorithm taken from a book which has long since been lost.

'The math that derived this stuff is pretty messy...  expressions are isolated and put into
'arrays.  we're essentially trying to find the values of the second derivative of each polynomial
'at each knot within the curve.  That's why theres only N-2 p's (where N is # points).
'later, we use the p's and u's to calculate curve points...

    For i = 2 To numOfPoints - 1
        d(i) = 2 * (iX(i + 1) - iX(i - 1))
    Next
    For i = 1 To numOfPoints - 1
        u(i) = iX(i + 1) - iX(i)
    Next
    For i = 2 To numOfPoints - 1
        w(i) = 6# * ((iY(i + 1) - iY(i)) / u(i) - (iY(i) - iY(i - 1)) / u(i - 1))
    Next
    For i = 2 To numOfPoints - 2
        w(i + 1) = w(i + 1) - w(i) * u(i) / d(i)
        d(i + 1) = d(i + 1) - u(i) * u(i) / d(i)
    Next
    p(1) = 0#
    For i = numOfPoints - 1 To 2 Step -1
        p(i) = (w(i) - u(i) * p(i + 1)) / d(i)
    Next
    p(numOfPoints) = 0#
End Sub

'By default, three points are provided: one at each corner, and one in the middle
Private Sub resetCurvePoints()

    numOfPoints = 3
    ReDim cPoints(0 To numOfPoints - 1) As fPoint
    
    Dim i As Long
    For i = 0 To numOfPoints - 1
        cPoints(i).pX = i / picDraw.ScaleWidth
        cPoints(i).pY = picDraw.ScaleHeight - (i / picDraw.ScaleHeight)
    Next i

End Sub

'Generates a spline from the current set of control points, and fills the results array with the relevant values
Private Sub fillResultsArray()

    'ReDim iX(0 To numOfPoints) As Double
    'ReDim iY(0 To numOfPoints) As Double
    ReDim p(0 To numOfPoints) As Double
    ReDim u(0 To numOfPoints) As Double

    'Clear the results array and reset the max/min variables
    ReDim cResults(-1 To picDraw.ScaleWidth) As Double
    
    Dim i As Long
    For i = -1 To picDraw.ScaleWidth
        cResults(i) = -1
    Next i
    
    Dim minX As Double, maxX As Double
    minX = picDraw.ScaleWidth
    maxX = -1
    
    'Now run a loop through the knots, calculating spline values as we go
    SetPandU
    Dim xPos As Long, yPos As Single
    
    For i = 0 To numOfPoints - 2
        For xPos = cPoints(i).pX To cPoints(i + 1).pX
            yPos = getCurvePoint(i, xPos)
            If xPos < minX Then minX = xPos
            If xPos > maxX Then maxX = xPos
            If yPos > 255 Then yPos = 254       'Force values to be in the 1-254 range (0-255 also
            If yPos < 0 Then yPos = 1           ' works, but is harder to see on the picture box)
            cResults(xPos) = yPos
        Next xPos
    Next i

    'cResults() now contains the y-coordinate of the spline for every x-coordinate in picDraw

End Sub
