VERSION 5.00
Begin VB.Form FormShadowHighlight 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Shadow / Midtone / Highlight"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12120
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
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicColor 
      Appearance      =   0  'Flat
      BackColor       =   &H007F7F7F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   465
      ScaleWidth      =   5625
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2640
      Width           =   5655
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9180
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10650
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartCheckBox chkAutoThreshold 
      Height          =   480
      Left            =   6120
      TabIndex        =   8
      Top             =   3240
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   847
      Caption         =   "use the median midtone for this image"
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
   Begin PhotoDemon.sliderTextCombo sltShadow 
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   1770
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   0
      Max             =   30
      SigDigits       =   2
      Value           =   0.05
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
   Begin PhotoDemon.sliderTextCombo sltHighlight 
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   4170
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   0
      Max             =   30
      SigDigits       =   2
      Value           =   0.05
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
   Begin VB.Label lblMidtone 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "midtone target color (click box to change):"
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
      TabIndex        =   6
      Top             =   2280
      Width           =   4530
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "highlights:"
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
      TabIndex        =   5
      Top             =   3840
      Width           =   1125
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -120
      TabIndex        =   3
      Top             =   5760
      Width           =   12495
   End
   Begin VB.Label lblShadow 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "shadows:"
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
      Top             =   1440
      Width           =   1005
   End
End
Attribute VB_Name = "FormShadowHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Shadow / Midtone / Highlight Adjustment Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 17/February/13
'Last updated: 28/April/13
'Last update: greatly simplify code by relying on new slider/text custom control
'
'Shadow / Midtone / Highlight recovery and correction tool.
'
'This tool is based heavily on the logic on PhotoDemon's "white balance" tool.  The Shadow and Highlight parameters
' refer to the amount of pixels in the image which will be ignored at either end of the spectrum, prior to stretching
' the histogram.  By ignoring more pixels at the bottom, shadows are emphasized.  By ignoring more pixels at the
' top, highlights are emphasized.
'
'Midtones are a separate beast.  The new midtone color functions as the midpoint of the image's new histogram.
' Pixels will be spread so that half fall below the midtone, and half fall above it.  Midtones are calculated
' separately for each of red, green, and blue, so this tool can be used to apply a particular color cast to an image.
' (Though the results are difficult to predict, so use with caution.)
'
'The automatic midtone detection algorithm works by finding the actual midpoint of the original image's histogram, and
' centering the new histogram using that midpoint as (127, 127, 127). This results in a theoretically "perfect"
' exposure, but as with most "theoretically perfect" color algorithms(e.g. histogram equalization), it is unlikely to
' offer ideal results.  Rather, think of it as a starting point from which you can more easily find your ideal midtone
' point.
'
'***************************************************************************

Option Explicit

Private Sub chkAutoThreshold_Click()
    If CBool(chkAutoThreshold) Then
        CalculateOptimalMidtone
    Else
        PicColor.backColor = RGB(127, 127, 127)
    End If
    updatePreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'The scroll bar max and min values are used to check the gamma input for validity
    If sltShadow.IsValid And sltHighlight.IsValid Then
        Me.Visible = False
        Process ShadowHighlight, sltShadow, sltHighlight, CLng(PicColor.backColor)
        Unload Me
    End If
    
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    setHandCursor PicColor
    
    'Render a preview
    updatePreview
    
End Sub

'Correct white balance by stretching the histogram and ignoring pixels above or below the 0.05% threshold
Public Sub ApplyShadowHighlight(Optional ByVal shadowClipping As Double = 0.05, Optional ByVal highlightClipping As Double = 0.05, Optional ByVal targetMidtone As Long = &H7F7F7F, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Adjusting shadows, midtones, and highlights..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    AdjustLayerShadowHighlight shadowClipping, highlightClipping, targetMidtone, workingLayer, toPreview
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub PicColor_Click()
    'Use a common dialog box to select a new color.  (In the future, perhaps I'll design a better custom box.)
    Dim retColor As Long
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    retColor = PicColor.backColor
    
    If CD1.VBChooseColor(retColor, True, True, False, Me.hWnd) Then
        PicColor.backColor = retColor
        updatePreview
    End If
End Sub

Private Sub CalculateOptimalMidtone()

    'Create a local array and point it at the pixel data of the image
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
            
    prepImageData tmpSA
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
            
    'Color variables
    Dim r As Long, g As Long, b As Long
            
    'Histogram tables
    Dim rLookup(0 To 255) As Long, gLookup(0 To 255) As Long, bLookup(0 To 255) As Long
    Dim NumOfPixels As Long
                
    'Loop through each pixel in the image, tallying values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
            
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
                
        rLookup(r) = rLookup(r) + 1
        gLookup(g) = gLookup(g) + 1
        bLookup(b) = bLookup(b) + 1
        
        'Increment the pixel count
        NumOfPixels = NumOfPixels + 1
        
    Next y
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    workingLayer.eraseLayer
    Set workingLayer = Nothing
            
    'Divide the number of pixels by two
    NumOfPixels = NumOfPixels \ 2
                       
    Dim rCount As Long, gCount As Long, bCount As Long
    x = 0
                    
    'Find the median value for each color channel
    Do
        rCount = rCount + rLookup(x)
        x = x + 1
    Loop While rCount < NumOfPixels
    
    rCount = x - 1
    
    x = 0
    
    Do
        gCount = gCount + gLookup(x)
        x = x + 1
    Loop While gCount < NumOfPixels
    
    gCount = x - 1
    
    x = 0
    
    Do
        bCount = bCount + bLookup(x)
        x = x + 1
    Loop While bCount < NumOfPixels
    
    bCount = x - 1
    
    PicColor.backColor = RGB(255 - rCount, 255 - gCount, 255 - bCount)
        
End Sub

Private Sub sltHighlight_Change()
    updatePreview
End Sub

Private Sub sltShadow_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    ApplyShadowHighlight sltShadow, sltHighlight, CLng(PicColor.backColor), True, fxPreview
End Sub

