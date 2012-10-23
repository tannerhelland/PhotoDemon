VERSION 5.00
Begin VB.Form FormSolarize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Solarize"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6255
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
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   6
      Top             =   120
      Width           =   2895
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
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsThreshold 
      Height          =   255
      Left            =   360
      Max             =   254
      Min             =   1
      TabIndex        =   1
      Top             =   3840
      Value           =   127
      Width           =   4935
   End
   Begin VB.TextBox txtThreshold 
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
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "127"
      Top             =   3795
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   4560
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
      Left            =   240
      TabIndex        =   4
      Top             =   3480
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
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/14/01
'Last updated: 10/September/12
'Last update: rewrote against new layer class; also, optimized with look-up tables
'
'Updated solarizing interface; it has been optimized for speed and ease-of-implementation.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max) Then
        Me.Visible = False
        Process Solarize, hsThreshold.Value
        Unload Me
    Else
        AutoSelectText txtThreshold
    End If
End Sub

'Subroutine for "solarizing" an image
' Inputs: solarize threshold [0,255], optional previewing information
Public Sub SolarizeImage(ByVal Threshold As Byte, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Solarizing image..."
    
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
            
    'Because solarize values are constant, we can use a look-up table to calculate them.  Very fast.
    Dim sLookup(0 To 255) As Byte
    For x = 0 To 255
        If x > Threshold Then sLookup(x) = 255 - x Else sLookup(x) = x
    Next x
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Perform the solarize in a single line, thanks to our pre-built look-up table
        ImageData(QuickVal + 2, y) = sLookup(ImageData(QuickVal + 2, y))
        ImageData(QuickVal + 1, y) = sLookup(ImageData(QuickVal + 1, y))
        ImageData(QuickVal, y) = sLookup(ImageData(QuickVal, y))
        
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
    
    'Create the previews
    DrawPreviewImage picPreview
    SolarizeImage hsThreshold.Value, True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'When the horizontal scroll bar is moved, update the preview and text box to match
Private Sub hsThreshold_Change()
    copyToTextBoxI txtThreshold, hsThreshold.Value
    SolarizeImage hsThreshold.Value, True, picEffect
End Sub

Private Sub hsThreshold_Scroll()
    copyToTextBoxI txtThreshold, hsThreshold.Value
    SolarizeImage hsThreshold.Value, True, picEffect
End Sub

'When the text box is changed, update the preview and text box to match (assuming the text box value is valid)
Private Sub txtThreshold_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtThreshold
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max, False, False) Then hsThreshold.Value = Val(txtThreshold)
End Sub

Private Sub txtThreshold_GotFocus()
    AutoSelectText txtThreshold
End Sub

