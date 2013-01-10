VERSION 5.00
Begin VB.Form FormGamma 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Gamma Correction"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12060
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
   ScaleWidth      =   804
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
   Begin VB.HScrollBar hsGamma 
      Height          =   255
      Left            =   6120
      Max             =   200
      Min             =   1
      TabIndex        =   3
      Top             =   3240
      Value           =   100
      Width           =   4935
   End
   Begin VB.TextBox txtGamma 
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
      Left            =   11160
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "1.00"
      Top             =   3180
      Width           =   615
   End
   Begin VB.ComboBox CboChannel 
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
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblBackground 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "strength:"
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
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "channel:"
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
      Top             =   1800
      Width           =   900
   End
End
Attribute VB_Name = "FormGamma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gamma Correction Handler
'Copyright ©2000-2013 by Tanner Helland
'Created: 12/May/01
'Last updated: 09/September/12
'Last update: rewrote all code against the new layer class
'
'Updated version of the gamma handler; fully optimized, it uses a look-up
' table and can correct any color channel.
'
'***************************************************************************

Option Explicit

'Update the preview when the user changes the channel combo box
Private Sub CboChannel_Click()
    GammaCorrect CSng(Val(txtGamma)), CByte(CboChannel.ListIndex), True, fxPreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    'The scroll bar max and min values are used to check the gamma input for validity
    If EntryValid(txtGamma, hsGamma.Min / 100, hsGamma.Max / 100) Then
        Me.Visible = False
        Process GammaCorrection, CSng(Val(txtGamma)), CByte(CboChannel.ListIndex)
        Unload Me
    Else
        AutoSelectText txtGamma
    End If
End Sub

Private Sub Form_Activate()
    
    'Populate the channels that gamma correction can operate on
    CboChannel.AddItem "RGB", 0
    CboChannel.AddItem "Red", 1
    CboChannel.AddItem "Green", 2
    CboChannel.AddItem "Blue", 3
    CboChannel.ListIndex = 0
    DoEvents
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Finally, render a preview
    GammaCorrect CSng(Val(txtGamma)), CByte(CboChannel.ListIndex), True, fxPreview
    
End Sub

'Basic gamma correction.  It's a simple function - use an exponent to adjust R/G/B values.
' Inputs: new gamma level, which channels to adjust (r/g/b/all), and optional preview information
Public Sub GammaCorrect(ByVal Gamma As Single, ByVal Method As Byte, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
     
    If toPreview = False Then Message "Adjusting gamma values..."
    
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
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Gamma can be easily applied using a look-up table
    Dim gLookup(0 To 255) As Byte
    Dim tmpVal As Single
    
    For x = 0 To 255
        tmpVal = x / 255
        tmpVal = tmpVal ^ (1 / Gamma)
        tmpVal = tmpVal * 255
        
        If tmpVal > 255 Then tmpVal = 255
        If tmpVal < 0 Then tmpVal = 0
        
        gLookup(x) = tmpVal
    Next x
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Correct the gamma values according to the channel requested by the user
        If Method = 0 Then
            r = gLookup(r)
            g = gLookup(g)
            b = gLookup(b)
        ElseIf Method = 1 Then
            r = gLookup(r)
        ElseIf Method = 2 Then
            g = gLookup(g)
        ElseIf Method = 3 Then
            b = gLookup(b)
        End If
        
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
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

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When the horizontal scroll bar is moved, change the text box to match
Private Sub hsGamma_Change()
    txtGamma.Text = Format(CSng(hsGamma.Value) / 100, "0.00")
    txtGamma.Refresh
    GammaCorrect CSng(Val(txtGamma)), CByte(CboChannel.ListIndex), True, fxPreview
End Sub

Private Sub hsGamma_Scroll()
    txtGamma.Text = Format(CSng(hsGamma.Value) / 100, "0.00")
    txtGamma.Refresh
    GammaCorrect CSng(Val(txtGamma)), CByte(CboChannel.ListIndex), True, fxPreview
End Sub

Private Sub txtGamma_GotFocus()
    AutoSelectText txtGamma
End Sub

'If the user changes the gamma value by hand, check it for numerical correctness, then change the horizontal scroll bar to match
Private Sub txtGamma_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtGamma, , True
    If EntryValid(txtGamma, hsGamma.Min / 100, hsGamma.Max / 100, False, False) Then hsGamma.Value = Val(txtGamma) * 100
End Sub
