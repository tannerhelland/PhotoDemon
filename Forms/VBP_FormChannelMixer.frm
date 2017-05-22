VERSION 5.00
Begin VB.Form FormChannelMixer 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Channel Mixer"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
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
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbChannel 
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
      ItemData        =   "VBP_FormChannelMixer.frx":0000
      Left            =   7920
      List            =   "VBP_FormChannelMixer.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   360
      Width           =   3900
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
   Begin RasterWave.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin RasterWave.sliderTextCombo sltRed 
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -200
      Max             =   200
      Value           =   -70
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
   Begin RasterWave.sliderTextCombo sltGreen 
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -200
      Max             =   200
      Value           =   200
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
   Begin RasterWave.sliderTextCombo sltBlue 
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   2880
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -200
      Max             =   200
      Value           =   -30
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
   Begin RasterWave.smartCheckBox chkMonochrome 
      Height          =   570
      Left            =   6120
      TabIndex        =   10
      Top             =   4800
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1005
      Caption         =   "monochrome"
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RasterWave.sliderTextCombo sltConstant 
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   3840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -200
      Max             =   200
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
   Begin RasterWave.smartCheckBox chkOverlay 
      Height          =   570
      Left            =   8640
      TabIndex        =   13
      Top             =   4800
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1005
      Caption         =   "overlay"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      Caption         =   "output channel:"
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
      Left            =   6120
      TabIndex        =   15
      Top             =   360
      Width           =   1665
   End
   Begin VB.Label lblConstant 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "constant"
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
      Left            =   6360
      TabIndex        =   12
      Top             =   4320
      Width           =   885
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label lblBlue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   285
      Left            =   6360
      TabIndex        =   4
      Top             =   3360
      Width           =   450
   End
   Begin VB.Label lblGreen 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label lblRed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   285
      Left            =   6360
      TabIndex        =   2
      Top             =   1440
      Width           =   345
   End
End
Attribute VB_Name = "FormChannelMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Channel Mixer Form
'Copyright ©2013 by Tanner Helland & audioglider
'Created: 08/June/13
'Last updated: 08/June/13
'Last update: Initial build
'
'Fairly simple and standard channel mixer form.  Layout and feature set derived from comparable tool
' in Photoshop.
'
'***************************************************************************

Option Explicit

Public Enum OutputType
    RedOutput = 0
    GreenOutput = 1
    BlueOutput = 2
    GrayOutput = 3
End Enum

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub chkMonochrome_Click()
    
    If chkMonochrome.Value = 1 Then
        cmbChannel.ListIndex = 3
    Else
        cmbChannel.ListIndex = 0
    End If
    
    updatePreview
    
End Sub

Private Sub chkOverlay_Click()
    updatePreview
End Sub

Private Sub cmbChannel_Change()
    updatePreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Validate all textbox entries
    If sltRed.IsValid And sltGreen.IsValid And sltBlue.IsValid And sltConstant.IsValid Then
        Me.Visible = False
        Process "Channel mixer", , buildParams(sltRed, sltGreen, sltBlue, sltConstant, CBool(chkMonochrome), CBool(chkOverlay), cmbChannel.ListIndex)
        Unload Me
    End If
    
End Sub

'Apply a new channel mixer to the image
' Input: offset for each of red, green, and blue
Public Sub ApplyChannelMixer(ByVal rVal As Single, ByVal gVal As Single, ByVal bVal As Single, ByVal cVal As Single, ByVal bMono As Boolean, ByVal bOverlay As Boolean, tOutput As OutputType, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Mixing color channels..."
    
    'Because these modifiers are constant throughout the image, we can build look-up tables for them
    Dim rLookup(0 To 255) As Byte, gLookup(0 To 255) As Byte, bLookup(0 To 255) As Byte
    
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
    
        'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim A As Long, Col As Long
    Dim newR As Long, newG As Long, newB As Long
    
    'Color look-up table
    Dim cLookup(0 To 255, 0 To 255) As Byte
    
    If bOverlay Then
        For x = 0 To 255
            For y = 0 To 255
                A = x + y
                If A > 255 Then A = 255
                cLookup(x, y) = A
            Next y
        Next x
    End If
    
    'Monochrome output
    If bMono Then
        tOutput = GrayOutput
    End If
    
    rVal = rVal / 100
    gVal = gVal / 100
    bVal = bVal / 100
    cVal = cVal / 100
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        Col = rVal * r + gVal * g + bVal * b
        
        Col = Col + cVal * 255
        If Col > 255 Then Col = 255
        If Col < 0 Then Col = 0
        
        If bOverlay Then
            Select Case tOutput
                Case 0
                    newR = cLookup(r, Col)
                Case 1
                    newG = cLookup(g, Col)
                Case 2
                    newB = cLookup(b, Col)
                Case 3
                    newR = cLookup(r, Col)
                    newG = cLookup(g, Col)
                    newB = cLookup(b, Col)
            End Select
        Else
            Select Case tOutput
                Case 0
                    newR = Col
                Case 1
                    newG = Col
                Case 2
                    newB = Col
                Case 3
                    newR = Col
                    newG = Col
                    newB = Col
            End Select
        End If
                                
        If newR < 0 Then newR = 0
        If newG < 0 Then newG = 0
        If newB < 0 Then newB = 0
        
        If newR > 255 Then newR = 255
        If newG > 255 Then newG = 255
        If newB > 255 Then newB = 255
                
        ImageData(QuickVal + 2, y) = newR
        ImageData(QuickVal + 1, y) = newG
        ImageData(QuickVal, y) = newB
                
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

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    cmbChannel.ListIndex = 3
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBlue_Change()
    updatePreview
End Sub

Private Sub sltConstant_Change()
    updatePreview
End Sub

Private Sub sltGreen_Change()
    updatePreview
End Sub

Private Sub sltRed_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    ApplyChannelMixer CSng(sltRed), CSng(sltGreen), CSng(sltBlue), CSng(sltConstant), CBool(chkMonochrome), CBool(chkOverlay), cmbChannel.ListIndex, True, fxPreview
End Sub

