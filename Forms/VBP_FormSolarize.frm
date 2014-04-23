VERSION 5.00
Begin VB.Form FormSolarize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Solarize"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12075
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   12075
      _ExtentX        =   21299
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
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltThreshold 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   2520
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   873
      Min             =   1
      Max             =   254
      Value           =   127
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
   Begin VB.Label lblThreshold 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "solarize threshold:"
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
      TabIndex        =   1
      Top             =   2160
      Width           =   1950
   End
End
Attribute VB_Name = "FormSolarize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Solarizing Effect Handler
'Copyright ©2001-2014 by Tanner Helland
'Created: 4/14/01
'Last updated: 24/August/13
'Last update: added command bar
'
'Updated solarizing interface; it has been optimized for speed and ease-of-implementation.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Subroutine for "solarizing" an image
' Inputs: solarize threshold [0,255], optional previewing information
Public Sub SolarizeImage(ByVal Threshold As Byte, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Solarizing image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
            
    'Because solarize values are constant, we can use a look-up table to calculate them.  Very fast.
    Dim sLookup(0 To 255) As Byte
    For x = 0 To 255
        If x > Threshold Then sLookup(x) = 255 - x Else sLookup(x) = x
    Next x
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Perform the solarize in a single line, thanks to our pre-built look-up table
        ImageData(quickVal + 2, y) = sLookup(ImageData(quickVal + 2, y))
        ImageData(quickVal + 1, y) = sLookup(ImageData(quickVal + 1, y))
        ImageData(quickVal, y) = sLookup(ImageData(quickVal, y))
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Solarize", , Str(sltThreshold.Value)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltThreshold.Value = 127
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Render a preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltThreshold_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then SolarizeImage sltThreshold.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

