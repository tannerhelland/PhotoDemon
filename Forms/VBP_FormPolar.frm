VERSION 5.00
Begin VB.Form FormPolar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Polar Coordinate Conversion"
   ClientHeight    =   6540
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   12105
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbEdges 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3225
      Width           =   5700
   End
   Begin VB.ComboBox cboConvert 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   4860
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10590
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   10
      Top             =   4200
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Caption         =   "quality"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   1
      Left            =   7920
      TabIndex        =   11
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      Caption         =   "speed"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   100
      Value           =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "if pixels lie outside the image..."
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
      Index           =   5
      Left            =   6000
      TabIndex        =   9
      Top             =   2850
      Width           =   3315
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "radius (percentage):"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   1920
      Width           =   2145
   End
   Begin VB.Label lblInterpolation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "render emphasis:"
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
      Left            =   6000
      TabIndex        =   3
      Top             =   3810
      Width           =   1845
   End
   Begin VB.Label lblConvert 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "conversion technique:"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   960
      Width           =   2325
   End
End
Attribute VB_Name = "FormPolar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Polar Coordinate Conversion Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 14/January/13
'Last updated: 15/January/13
'Last update: added support for custom edge handling
'
'This tool allows the user to convert an image between rectangular and polar coordinates.  An optional polar
' inversion technique is also supplied (as this is used by Paint.NET).
'
'The transformation used by this tool is a modified version of a transformation originally written by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.  You may download his
' original version at the following link (good as of 07 January '13): http://www.jhlabs.com/ip/filters/index.html
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub cboConvert_Click()
    updatePreview
End Sub

Private Sub cmbEdges_Click()
    updatePreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Before rendering anything, check to make sure the text boxes have valid input
    If sltRadius.IsValid Then
        Me.Visible = False
        Process ConvertPolar, cboConvert.ListIndex, sltRadius.Value, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value
        Unload Me
    End If
    
End Sub

'Convert an image to/from polar coordinates.
' INPUT PARAMETERS FOR CONVERSION:
' 0) Convert rectangular to polar
' 1) Convert polar to rectangular
' 2) Polar inversion
Public Sub ConvertToPolar(ByVal conversionMethod As Long, ByVal polarRadius As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If toPreview = False Then Message "Performing polar coordinate conversion..."
    
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
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, curLayerValues.MaxX, curLayerValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
          
    'Polar conversion requires a number of specialized variables
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Rotation values
    Dim theta As Double, sRadius As Double, sRadius2 As Double, sDistance As Double
    Dim r As Double, t As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curLayerValues.Width
    tHeight = curLayerValues.Height
    sRadius = Sqr(tWidth * tWidth + tHeight * tHeight) / 2
              
    sRadius = sRadius * (polarRadius / 100)
    sRadius2 = sRadius * sRadius
        
    polarRadius = 1 / (polarRadius / 100)
        
    Dim iAspect As Double
    iAspect = tHeight / tWidth
              
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Each polar conversion requires a unique set of code
        Select Case conversionMethod
        
            'Rectangular to polar
            Case 0
                            
                'Remap the coordinates around a center point of (0, 0)
                nX = x - midX
                nY = y - midY
                
                'Calculate distance automatically
                sDistance = (nX * nX) + (nY * nY)
                
                If sDistance <= sRadius2 Then
                
                    'X is handled differently based on its relation to the center of the image
                    If x >= midX Then
                        nX = x - midX
                        If y > midY Then
                            theta = PI - Atn(nX / nY)
                            r = Sqr(sDistance)
                        ElseIf y < midY Then
                            theta = Atn(nX / (midY - y))
                            r = Sqr(nX * nX + (midY - y) * (midY - y))
                        Else
                            theta = PI_HALF
                            r = nX
                        End If
                    Else
                        nX = midX - x
                        If y > midY Then
                            theta = PI + Atn(nX / nY)
                            r = Sqr(sDistance)
                        ElseIf y < midY Then
                            theta = PI_DOUBLE - Atn(nX / (midY - y))
                            r = Sqr(nX * nX + (midY - y) * (midY - y))
                        Else
                            theta = PI * 1.5
                            r = nX
                        End If
                    End If
                    
                    srcX = (finalX) - (finalX / PI_DOUBLE * theta)
                    srcY = (finalY + 1) * r / sRadius
                    
                Else
                
                    srcX = x
                    srcY = y
                    
                End If
                
            'Polar to rectangular
            Case 1
            
                'Remap the coordinates around a center point of (0, 0)
                nX = x - midX
                nY = y - midY
                
                'Calculate distance automatically
                sDistance = (nX * nX) + (nY * nY)
            
                If sDistance <= sRadius2 Then
                
                    theta = x / (finalX + 1) * PI_DOUBLE
                    
                    If theta >= (PI * 1.5) Then
                        t = PI_DOUBLE - theta
                    ElseIf theta >= PI Then
                        t = theta - PI
                    ElseIf theta > PI_HALF Then
                        t = PI - theta
                    Else
                        t = theta
                    End If
                    
                    r = sRadius * (y / (finalY + 1))
                    
                    nX = -r * Sin(t)
                    nY = r * Cos(t)
                    
                    If theta >= 1.5 * PI Then
                        srcX = midX - nX
                        srcY = midY - nY
                    ElseIf theta >= PI Then
                        srcX = midX - nX
                        srcY = midY + nY
                    ElseIf theta >= PI_HALF Then
                        srcX = midX + nX
                        srcY = midY + nY
                    Else
                        srcX = midX + nX
                        srcY = midY - nY
                    End If
                    
                Else
                
                    srcX = x
                    srcY = y
                
                End If
                            
            'Polar inversion
            Case 2
            
                'Remap the coordinates around a center point of (0, 0)
                nX = x - midX
                nY = y - midY
                
                'Calculate distance automatically
                sDistance = (nX * nX) + (nY * nY)
                
                If sDistance <> 0 Then
                    srcX = midX + midX * midX * (nX / sDistance) * polarRadius
                    srcY = midY + midY * midY * (nY / sDistance) * polarRadius
                    srcX = Modulo(srcX, (finalX + 1))
                    srcY = Modulo(srcY, (finalY + 1))
                Else
                    srcX = x
                    srcY = y
                End If
            
        End Select
        
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
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

Private Sub Form_Activate()
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    popDistortEdgeBox cmbEdges, EDGE_ERASE
    
    'Populate the polar conversion technique drop-down
    cboConvert.AddItem "Rectangular to polar", 0
    cboConvert.AddItem "Polar to rectangular", 1
    cboConvert.AddItem "Polar inversion", 2
    cboConvert.ListIndex = 0
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me, m_ToolTip
        
    'Create the preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

'This is a modified module function; it handles negative values specially to ensure they work with our kaleidoscope function
Private Function Modulo(ByVal Quotient As Double, ByVal Divisor As Double) As Double
    Modulo = Quotient - Fix(Quotient / Divisor) * Divisor
    If Modulo < 0 Then Modulo = Modulo + Divisor
End Function

Private Sub sltRadius_Change()
    updatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()
    ConvertToPolar cboConvert.ListIndex, sltRadius.Value, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, True, fxPreview
End Sub
