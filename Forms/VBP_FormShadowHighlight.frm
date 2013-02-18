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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   5655
   End
   Begin VB.TextBox txtHighlight 
      Alignment       =   2  'Center
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
      Left            =   11040
      TabIndex        =   8
      Text            =   "0.05"
      Top             =   4155
      Width           =   735
   End
   Begin VB.HScrollBar hsHighlight 
      Height          =   255
      Left            =   6120
      Max             =   3000
      Min             =   1
      TabIndex        =   7
      Top             =   4200
      Value           =   5
      Width           =   4815
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
   Begin VB.HScrollBar hsShadow 
      Height          =   255
      Left            =   6120
      Max             =   3000
      Min             =   1
      TabIndex        =   2
      Top             =   1800
      Value           =   5
      Width           =   4815
   End
   Begin VB.TextBox txtShadow 
      Alignment       =   2  'Center
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
      Left            =   11040
      TabIndex        =   3
      Text            =   "0.05"
      Top             =   1755
      Width           =   735
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
   Begin PhotoDemon.smartCheckBox chkAutoThreshold 
      Height          =   480
      Left            =   6120
      TabIndex        =   12
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   3840
      Width           =   1125
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -120
      TabIndex        =   5
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
      TabIndex        =   4
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
'Last updated: 17/February/13
'Last update: initial build
'
'White balance handler.  Unlike other programs, which shove this under the Levels dialog as an "auto levels"
' function, I consider it worthy of its own interface.  The reason is - white balance is an important function.
' It's arguably more useful than the Levels dialog, especially to a casual user, because it automatically
' calculates levels according to a reliable, often-accurate algorithm.  Rather than forcing the user through the
' Levels dialog (because really, how many people know that Auto Levels is actually White Balance in photography
' parlance?), PhotoDemon provides a full implementation of custom white balance handling.
' The value box on the form is the percentage of pixels ignored at the top and bottom of the histogram.
' 0.05 is the recommended default.  I've specified 1.5 as the maximum, but there's no reason it couldn't be set
' higher... just be forewarned that higher values (obviously) blow out the picture with increasing strength.
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
    If Not EntryValid(txtShadow, hsShadow.Min / 100, hsShadow.Max / 100) Then
        AutoSelectText txtShadow
        Exit Sub
    End If
    
    If Not EntryValid(txtHighlight, hsHighlight.Min / 100, hsHighlight.Max / 100) Then
        AutoSelectText txtHighlight
        Exit Sub
    End If
    
    Me.Visible = False
    Process ShadowHighlight, CSng(hsShadow / 100), CSng(hsHighlight / 100), CLng(PicColor.backColor)
    Unload Me
    
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

'When the horizontal scroll bar is moved, change the text box to match
Private Sub hsShadow_Change()
    copyToTextBoxF CSng(hsShadow) / 100, txtShadow
    updatePreview
End Sub

Private Sub hsShadow_Scroll()
    copyToTextBoxF CSng(hsShadow) / 100, txtShadow
    updatePreview
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

Private Sub txtShadow_GotFocus()
    AutoSelectText txtShadow
End Sub

'If the user changes the text box value by hand, check it for numerical correctness, then change the horizontal scroll bar to match
Private Sub txtShadow_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtShadow, , True
    If EntryValid(txtShadow, hsShadow.Min / 100, hsShadow.Max / 100, False, False) Then hsShadow.Value = Val(txtShadow) * 100
End Sub

Private Sub hsHighlight_Change()
    copyToTextBoxF CSng(hsHighlight) / 100, txtHighlight
    updatePreview
End Sub

Private Sub hsHighlight_Scroll()
    copyToTextBoxF CSng(hsHighlight) / 100, txtHighlight
    updatePreview
End Sub

Private Sub txtHighlight_GotFocus()
    AutoSelectText txtHighlight
End Sub

Private Sub txtHighlight_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtHighlight, , True
    If EntryValid(txtHighlight, hsHighlight.Min / 100, hsHighlight.Max / 100, False, False) Then hsHighlight.Value = Val(txtHighlight) * 100
End Sub

Private Sub updatePreview()
    ApplyShadowHighlight CSng(hsShadow / 100), CSng(hsHighlight / 100), CLng(PicColor.backColor), True, fxPreview
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
    'PicColor.backColor = RGB(rCount, gCount, bCount)
        
End Sub
