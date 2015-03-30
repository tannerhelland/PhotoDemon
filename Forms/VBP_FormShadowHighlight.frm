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
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Index           =   0
      Left            =   5880
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   4
      Top             =   120
      Width           =   6135
      Begin PhotoDemon.sliderTextCombo sltShadow 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   873
         Min             =   -50
         Max             =   50
         SigDigits       =   1
      End
      Begin PhotoDemon.sliderTextCombo sltHighlight 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2970
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   873
         Min             =   -50
         Max             =   50
         SigDigits       =   1
      End
      Begin PhotoDemon.smartCheckBox chkAutoThreshold 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   582
         Caption         =   "use the median midtone for this image"
      End
      Begin PhotoDemon.colorSelector colorPicker 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   873
         curColor        =   8421504
      End
      Begin VB.Label lblMidtone 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "midtone target color (right-click preview to change):"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   5550
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
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1005
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
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   1125
      End
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14802140
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      ColorSelection  =   -1  'True
   End
   Begin PhotoDemon.buttonStrip btsOptions 
      Height          =   600
      Left            =   6240
      TabIndex        =   2
      Top             =   5040
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   1058
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Index           =   1
      Left            =   5880
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   5
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "options:"
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
      Index           =   6
      Left            =   6000
      TabIndex        =   3
      Top             =   4680
      Width           =   870
   End
End
Attribute VB_Name = "FormShadowHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Shadow / Midtone / Highlight Adjustment Tool
'Copyright 2013-2015 by Tanner Helland
'Created: 17/February/13
'Last updated: 28/March/15
'Last update: total overhaul of the shadow/highlight adjustment strategy
'
'This tool is based heavily on the logic on PhotoDemon's Curves tool.  The Shadow and Highlight parameters control
' an auto-generated S-curve, which allows the function to adjust regions of the image intelligently, while still
' maintaining a smooth transition between shadow and highlight regions.
'
'Midtones work similarly.  The midtone color selector is used to calculate a midpoint for the image's new luminance
' curve.  When automatic midtone calculation is active, pixels will be roughly spread so that half fall below the
' midtone, and half fall above it.  Midtones are calculated separately for each channel, so this tool is now capable
' of adjusting shadows and/or highlights without disturbing an image's color distribution.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_Tooltip As clsToolTip

Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    picContainer(buttonIndex).Visible = True
    picContainer(1 - buttonIndex).Visible = False
End Sub

Private Sub chkAutoThreshold_Click()
    If CBool(chkAutoThreshold) Then
        CalculateOptimalMidtone
    Else
        colorPicker.Color = RGB(127, 127, 127)
    End If
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Shadows and highlights", , buildParams(sltShadow, sltHighlight, CLng(colorPicker.Color)), UNDO_LAYER
End Sub

Private Sub cmdBar_RandomizeClick()
    chkAutoThreshold.Value = vbUnchecked
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    If CBool(chkAutoThreshold) Then CalculateOptimalMidtone
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    chkAutoThreshold.Value = vbUnchecked
    colorPicker.Color = RGB(127, 127, 127)
End Sub

Private Sub colorPicker_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_Tooltip = New clsToolTip
    makeFormPretty Me, m_Tooltip
    
    'Render a preview
    updatePreview
    
End Sub

'Correct white balance by stretching the histogram and ignoring pixels above or below the 0.05% threshold
Public Sub ApplyShadowHighlight(Optional ByVal shadowClipping As Double = 0.05, Optional ByVal highlightClipping As Double = 0.05, Optional ByVal targetMidtone As Long = &H7F7F7F, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Adjusting shadows, midtones, and highlights..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    AdjustDIBShadowHighlight shadowClipping, highlightClipping, targetMidtone, workingDIB, toPreview
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Load()
    
    'Set up the basic/advanced panels
    btsOptions.AddItem "basic", 0
    btsOptions.AddItem "advanced", 1
    btsOptions.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub fxPreview_ColorSelected()
    colorPicker.Color = fxPreview.SelectedColor
    updatePreview
End Sub

Private Sub CalculateOptimalMidtone()

    'Create a local array and point it at the pixel data of the image
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
            
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
                
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
                    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
            
    'Color variables
    Dim r As Long, g As Long, b As Long
            
    'Histogram tables
    Dim rLookup(0 To 255) As Long, gLookUp(0 To 255) As Long, bLookup(0 To 255) As Long
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
        gLookUp(g) = gLookUp(g) + 1
        bLookup(b) = bLookup(b) + 1
        
        'Increment the pixel count
        NumOfPixels = NumOfPixels + 1
        
    Next y
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    workingDIB.eraseDIB
    Set workingDIB = Nothing
            
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
        gCount = gCount + gLookUp(x)
        x = x + 1
    Loop While gCount < NumOfPixels
    
    gCount = x - 1
    
    x = 0
    
    Do
        bCount = bCount + bLookup(x)
        x = x + 1
    Loop While bCount < NumOfPixels
    
    bCount = x - 1
    
    colorPicker.Color = RGB(255 - rCount, 255 - gCount, 255 - bCount)
        
End Sub

Private Sub sltHighlight_Change()
    updatePreview
End Sub

Private Sub sltShadow_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyShadowHighlight sltShadow, sltHighlight, CLng(colorPicker.Color), True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

