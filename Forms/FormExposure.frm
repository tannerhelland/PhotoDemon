VERSION 5.00
Begin VB.Form FormExposure 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exposure"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picChart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   8280
      ScaleHeight     =   159
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   6
      Top             =   480
      Width           =   3495
   End
   Begin PhotoDemon.sliderTextCombo sltExposure 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   3480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   5
      SigDigits       =   2
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9030
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10500
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
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "new exposure curve:"
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
      Height          =   1005
      Index           =   2
      Left            =   5880
      TabIndex        =   7
      Top             =   1530
      Width           =   2280
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "exposure (EV):"
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
      Top             =   3120
      Width           =   1590
   End
End
Attribute VB_Name = "FormExposure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Exposure Form
'Copyright ©2013 by audioglider
'Created: 13/July/13
'Last updated: 13/July/13
'Last update: Initial build
'
'Simple exposure adjustment
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()

    If sltExposure.IsValid Then
        Me.Visible = False
        Process "Exposure", , buildParams(sltExposure)
        Unload Me
    End If
    
End Sub

Public Sub Exposure(ByVal exposureAdjust As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Adjusting exposure..."
    
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
    
    Dim r As Long, g As Long, b As Long
    
    'exposure can be easily applied using a look-up table
    Dim gLookup(0 To 2, 0 To 255) As Byte
    Dim tmpVal As Double
    
    For y = 0 To 2
    For x = 0 To 255
        tmpVal = x / 255
        Select Case y
            Case 0
                tmpVal = (1 - Exp(-tmpVal * exposureAdjust))
            Case 1
                tmpVal = (1 - Exp(-tmpVal * exposureAdjust))
            Case 2
                tmpVal = (1 - Exp(-tmpVal * exposureAdjust))
        End Select
        tmpVal = tmpVal * 255
        
        If tmpVal > 255 Then tmpVal = 255
        If tmpVal < 0 Then tmpVal = 0
        
        gLookup(y, x) = tmpVal
    Next x
    Next y

    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        ImageData(QuickVal + 2, y) = gLookup(0, r)
        ImageData(QuickVal + 1, y) = gLookup(1, g)
        ImageData(QuickVal, y) = gLookup(2, b)
        
    Next y
        If Not toPreview Then
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

Private Sub Form_Activate()

    'Draw a preview of the effect
    updatePreview
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sltExposure_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    
    Dim prevX As Double, prevY As Double
    Dim curX As Double, curY As Double
    Dim x As Long, y As Long
    
    Dim xWidth As Long, yHeight As Long
    xWidth = picChart.ScaleWidth
    yHeight = picChart.ScaleHeight
        
    'Clear out the old chart and draw a gray line across the diagonal for reference
    picChart.Picture = LoadPicture("")
    picChart.ForeColor = RGB(127, 127, 127)
    GDIPlusDrawLineToDC picChart.hDC, 0, yHeight, xWidth, 0, RGB(127, 127, 127)
    
    Dim expVal As Double, tmpVal As Double
      
    expVal = sltExposure
            
    picChart.ForeColor = RGB(0, 0, 255)
        
    prevX = 0
    prevY = yHeight
    curX = 0
    curY = yHeight
    
    If expVal > 0 Then
        'Draw the curve
        For x = 0 To xWidth
            tmpVal = x / xWidth
            tmpVal = (1 - Exp(-tmpVal * expVal))
            tmpVal = yHeight - (tmpVal * yHeight)
            curY = tmpVal
            curX = x
            GDIPlusDrawLineToDC picChart.hDC, prevX, prevY, curX, curY, picChart.ForeColor
            prevX = curX
            prevY = curY
        Next x
    End If
    
    picChart.Picture = picChart.Image
    picChart.Refresh

    Exposure sltExposure, True, fxPreview
End Sub
