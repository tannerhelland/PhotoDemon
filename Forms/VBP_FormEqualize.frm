VERSION 5.00
Begin VB.Form FormEqualize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Equalize Histogram"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
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
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLuminance 
      Appearance      =   0  'Flat
      Caption         =   "luminance"
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
      Height          =   345
      Left            =   4440
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CheckBox chkBlue 
      Appearance      =   0  'Flat
      Caption         =   "blue"
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
      Height          =   345
      Left            =   3120
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox chkGreen 
      Appearance      =   0  'Flat
      Caption         =   "green"
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
      Height          =   345
      Left            =   1680
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CheckBox chkRed 
      Appearance      =   0  'Flat
      Caption         =   "red"
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
      Height          =   345
      Left            =   480
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "after"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "before"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label lblEqualize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "equalize:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   945
   End
End
Attribute VB_Name = "FormEqualize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Histogram Equalization Interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 19/September/12
'Last updated: 19/September/12
'Last update: initial build.  Originally, the equalize functions were handled from menu entries on the main form, but this
'             was stupid - especially if you wanted to equalize multiple channels simultaneously.  The new form allows
'             more freedom.  Also, all Equalize routines are now condensed into one.
'
'Module for handling histogram equalization.  Any combination of red, green, blue, and luminance can be equalized, but if
' luminance is selected it will get precedent (e.g. it will be equalized first).
'
'***************************************************************************

Option Explicit

'Whenever a check box is changed, redraw the preview
Private Sub chkBlue_Click()
    EqualizeHistogram CBool(chkRed.Value), CBool(chkGreen.Value), CBool(chkBlue.Value), CBool(chkLuminance.Value), True, picEffect
End Sub

Private Sub chkGreen_Click()
    EqualizeHistogram CBool(chkRed.Value), CBool(chkGreen.Value), CBool(chkBlue.Value), CBool(chkLuminance.Value), True, picEffect
End Sub

Private Sub chkLuminance_Click()
    EqualizeHistogram CBool(chkRed.Value), CBool(chkGreen.Value), CBool(chkBlue.Value), CBool(chkLuminance.Value), True, picEffect
End Sub

Private Sub chkRed_Click()
    EqualizeHistogram CBool(chkRed.Value), CBool(chkGreen.Value), CBool(chkBlue.Value), CBool(chkLuminance.Value), True, picEffect
End Sub

'OK button
Private Sub cmdOK_Click()
    
    Me.Visible = False
        
    Process Equalize, CBool(chkRed.Value), CBool(chkGreen.Value), CBool(chkBlue.Value), CBool(chkLuminance.Value)
    
    Unload Me
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    'Create the previews
    DrawPreviewImage picPreview
    EqualizeHistogram CBool(chkRed.Value), CBool(chkGreen.Value), CBool(chkBlue.Value), CBool(chkLuminance.Value), True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'Equalize the red, green, blue, and/or Luminance channels of an image
' (Technically Luminance isn't a channel, but you know what I mean.)
Public Sub EqualizeHistogram(ByVal HandleR As Boolean, ByVal HandleG As Boolean, ByVal HandleB As Boolean, ByVal HandleL As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Analyzing image histogram..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
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
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    SetProgBarMax finalX * 2
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim R As Long, g As Long, b As Long
    Dim h As Single, S As Single, l As Single
    Dim lInt As Long
    
    'Histogram variables
    Dim rData(0 To 255) As Single, gData(0 To 255) As Single, bData(0 To 255) As Single
    Dim rDataInt(0 To 255) As Long, gDataInt(0 To 255) As Long, bDataInt(0 To 255) As Long
    Dim lData(0 To 255) As Single
    Dim lDataInt(0 To 255) As Long
        
    'Loop through each pixel in the image, converting values as we go.
    ' (This step is so fast that I calculate all channels, even those not being converted, with the exception of luminance.)
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        R = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Store those values in the histogram
        rDataInt(R) = rDataInt(R) + 1
        gDataInt(g) = gDataInt(g) + 1
        bDataInt(b) = bDataInt(b) + 1
        
        'Because luminance is slower to calculate, only calculate it if absolutely necessary
        If HandleL Then
            lInt = getLuminance(R, g, b)
            lDataInt(lInt) = lDataInt(lInt) + 1
        End If
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'Compute a scaling factor based on the number of pixels in the image
    Dim scaleFactor As Single
    scaleFactor = 255 / (curLayerValues.Width * curLayerValues.Height)
    
    'Compute red if requested
    If HandleR = True Then
        rData(0) = rDataInt(0) * scaleFactor
        For x = 1 To 255
            rData(x) = rData(x - 1) + (scaleFactor * rDataInt(x))
        Next x
    End If
    
    'Compute green if requested
    If HandleG = True Then
        gData(0) = gDataInt(0) * scaleFactor
        For x = 1 To 255
            gData(x) = gData(x - 1) + (scaleFactor * gDataInt(x))
        Next x
    End If
    
    'Compute blue if requested
    If HandleB = True Then
        bData(0) = bDataInt(0) * scaleFactor
        For x = 1 To 255
            bData(x) = bData(x - 1) + (scaleFactor * bDataInt(x))
        Next x
    End If
    
    'Compute luminance if requested
    If HandleL = True Then
        lData(0) = lDataInt(0) * scaleFactor
        For x = 1 To 255
            lData(x) = lData(x - 1) + (scaleFactor * lDataInt(x))
        Next x
    End If
    
    'Make sure all look-up values are in valid byte range (e.g. [0,255])
    For x = 0 To 255
        
        If rData(x) > 255 Then
            rDataInt(x) = 255
        Else
            rDataInt(x) = Int(rData(x))
        End If
        
        If gData(x) > 255 Then
            gDataInt(x) = 255
        Else
            gDataInt(x) = Int(gData(x))
        End If
        
        If bData(x) > 255 Then
            bDataInt(x) = 255
        Else
            bDataInt(x) = Int(bData(x))
        End If
        
        If lData(x) > 255 Then
            lDataInt(x) = 255
        Else
            lDataInt(x) = Int(lData(x))
        End If
        
    Next x
    
    'Apply the equalized values
    If toPreview = False Then Message "Equalizing image..."
    
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the RGB values
        R = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'If luminance has been requested, calculate it before messing with any of the color channels
        If HandleL Then
            tRGBToHSL R, g, b, h, S, l
            tHSLToRGB h, S, lDataInt(Int(l * 255)) / 255, R, g, b
        End If
        
        'Next, calculate new values for the color channels, based on what is being equalized
        If HandleR Then R = rDataInt(R)
        If HandleG Then g = gDataInt(g)
        If HandleB Then b = bDataInt(b)
        
        'Assign our new values back into the pixel array
        ImageData(QuickVal + 2, y) = R
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x + finalX
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub
