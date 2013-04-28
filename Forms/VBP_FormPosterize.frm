VERSION 5.00
Begin VB.Form FormPosterize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Posterize"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11970
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
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10470
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
   Begin PhotoDemon.sliderTextCombo sltBits 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   2640
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Min             =   1
      Max             =   8
      Value           =   2
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
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   5760
      Width           =   12015
   End
   Begin VB.Label lblBits 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "posterize to this many bits:"
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
      Height          =   405
      Left            =   6000
      TabIndex        =   2
      Top             =   2280
      Width           =   2880
   End
End
Attribute VB_Name = "FormPosterize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Posterizing Effect Handler
'Copyright ©2001-2013 by Tanner Helland
'Created: 4/15/01
'Last updated: 28/April/13
'Last update: simplify code by relying on new slider/text custom control
'
'Updated posterizing interface; it has been optimized for speed and
'  ease-of-implementation.  If only VB had bit-shift operators....
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    If sltBits.IsValid Then
        Me.Visible = False
        Process Posterize, sltBits.Value
        Unload Me
    End If
End Sub

'Subroutine for reducing the representative bits in an image
Public Sub PosterizeImage(ByVal NumOfBits As Byte, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Posterizing image..."
    
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
    progBarCheck = findBestProgBarValue()
    
    'pStep is the distance between values that X number of bits allows
    Dim pStep As Double
    pStep = 255 / (2 ^ CLng(NumOfBits) - 1)
    
    'Look-up tables make this far more efficient
    Dim LookUp(0 To 255) As Byte
    For x = 0 To 255
        'Add 0.5 so that values are rounded, not truncated (slightly better results)
        LookUp(x) = CByte(Int(Int(CDbl(x) / pStep + 0.5) * pStep))
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'The posterize can be implemented in a single look-up per color
        ImageData(QuickVal + 2, y) = LookUp(ImageData(QuickVal + 2, y))
        ImageData(QuickVal + 1, y) = LookUp(ImageData(QuickVal + 1, y))
        ImageData(QuickVal, y) = LookUp(ImageData(QuickVal, y))
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
     
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Create a preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBits_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    PosterizeImage sltBits.Value, True, fxPreview
End Sub
