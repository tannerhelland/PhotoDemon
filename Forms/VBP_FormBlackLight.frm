VERSION 5.00
Begin VB.Form FormBlackLight 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Black Light Options"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6285
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
   ScaleHeight     =   355
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3330
      TabIndex        =   0
      Top             =   4710
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   4710
      Width           =   1365
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsIntensity 
      Height          =   255
      Left            =   360
      Max             =   10
      Min             =   1
      TabIndex        =   2
      Top             =   3840
      Value           =   2
      Width           =   4935
   End
   Begin VB.TextBox txtIntensity 
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
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "2"
      Top             =   3780
      Width           =   615
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -720
      TabIndex        =   9
      Top             =   4560
      Width           =   7095
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "intensity:"
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
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "FormBlackLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Blacklight Form
'Copyright ©2001-2013 by Tanner Helland
'Created: some time 2001
'Last updated: 08/September/12
'Last update: rewrote effect against new layer class, merged previewing into core effect
'
'I found this effect on accident, and it has gradually become one of my favorite effects.
' Visually stunning on many photographs.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    If EntryValid(txtIntensity, hsIntensity.Min, hsIntensity.Max) Then
        Me.Visible = False
        Process BlackLight, hsIntensity.Value
        Unload Me
    Else
        AutoSelectText txtIntensity
    End If
    
End Sub

'Perform a blacklight filter
'Input: strength of the filter (min 1, no real max - but above 7 it becomes increasingly blown-out)
Public Sub fxBlackLight(Optional ByVal Weight As Long = 2, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Illuminating image with imaginary blacklight..."
    
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
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate the gray value using the look-up table
        grayVal = grayLookUp(r + g + b)
        
        'Perform the blacklight conversion
        r = Abs(r - grayVal) * Weight
        g = Abs(g - grayVal) * Weight
        b = Abs(b - grayVal) * Weight
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal + 2, y) = b
        
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

    'Create a copy of the original image to the preview picture box
    DrawPreviewImage picPreview
    
    'Draw a preview of the effect on the neighboring picture box
    fxBlackLight hsIntensity.Value, True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The next three routines keep the scroll bar and text box values in sync
Private Sub hsIntensity_Change()
    fxBlackLight hsIntensity.Value, True, picEffect
    copyToTextBoxI txtIntensity, hsIntensity.Value
End Sub

Private Sub hsIntensity_Scroll()
    fxBlackLight hsIntensity.Value, True, picEffect
    copyToTextBoxI txtIntensity, hsIntensity.Value
End Sub

Private Sub txtIntensity_GotFocus()
    AutoSelectText txtIntensity
End Sub

Private Sub txtIntensity_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtIntensity
    If EntryValid(txtIntensity, hsIntensity.Min, hsIntensity.Max, False, False) Then
        hsIntensity.Value = Val(txtIntensity)
    End If
End Sub
