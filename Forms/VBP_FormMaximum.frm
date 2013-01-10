VERSION 5.00
Begin VB.Form FormRank 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Custom Rank Filter"
   ClientHeight    =   6540
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
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
   Begin VB.HScrollBar hsRadius 
      Height          =   255
      Left            =   6120
      Max             =   25
      Min             =   1
      TabIndex        =   4
      Top             =   3120
      Value           =   1
      Width           =   4935
   End
   Begin VB.TextBox txtRadius 
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
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "1"
      Top             =   3075
      Width           =   615
   End
   Begin VB.ComboBox cboRank 
      Appearance      =   0  'Flat
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
      Width           =   3135
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
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label lblRank 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "rank method:"
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
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Label lblRadius 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "radius:"
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
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "FormRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Rank (a.k.a. High/Low Pass, Dilate/Erode) Filter Interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 26/October/06
'Last update: Image preview and additional optimizations. Image previewing
'             was a beast to add to this function o_O...
'Still needs: replace gotos with text labels
'
'Optimized but non-processable rank filters.  Max, min, and the all-new,
'all-original extreme version.  Very cool.
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify the strength to be representative of the final filter.  This means dividing by the
' original image width in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

Private Sub cboRank_Click()
    CustomRankFilter hsRadius.Value, cboRank.ListIndex, True, fxPreview
End Sub

Private Sub cboRank_KeyDown(KeyCode As Integer, Shift As Integer)
    CustomRankFilter hsRadius.Value, cboRank.ListIndex, True, fxPreview
End Sub

'OK Button
Private Sub cmdOK_Click()
    If EntryValid(txtRadius, hsRadius.Min, hsRadius.Max) Then
        Me.Visible = False
        Process CustomRank, hsRadius.Value, cboRank.ListIndex
        Unload Me
    Else
        AutoSelectText txtRadius
    End If
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'A powerful routine for any kind of rank filter at any radius
Public Sub CustomRankFilter(ByVal Radius As Long, ByVal RankType As Byte, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
        
    If toPreview = False Then
        Select Case RankType
            Case 0
                Message "Dilating image via maximum (high-pass) rank filter..."
            Case 1
                Message "Eroding image via minimum (low-pass) rank filter..."
            Case 2
                Message "Redrawing image via extreme rank filter..."
        End Select
    End If
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-mosaic'ed pixels from affecting the results of later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, c As Long, d As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    'If this is a preview, we need to adjust the xDiffuse and yDiffuse values to match the size of the preview box
    If toPreview Then
        Radius = (Radius / iWidth) * curLayerValues.Width
        If Radius = 0 Then Radius = 1
    End If
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, QuickValDst As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Rank calculations require a lot of variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long, grayValOriginal As Long
    Dim MaxX As Long, MaxY As Long
    Dim MaxTotal As Long
        
    'Because gray values are constant, we can use a look-up table to calculate them.
    ' Note that we divide each gray value by two to minimize the the effect of the unfade.
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'The total number needs to be set based on the type of rank analysis we're performing
        Select Case RankType
            Case 0
                MaxTotal = -1
            Case 1
                MaxTotal = 256
            Case 2
                MaxTotal = -1
        End Select
        
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
            
        grayValOriginal = gLookup(r + g + b)
        
        For c = x - Radius To x + Radius
            QuickValInner = c * qvDepth
        For d = y - Radius To y + Radius
        
            If c < 0 Then GoTo NextRankPixel
            If c > finalX Then GoTo NextRankPixel
            If d < 0 Then GoTo NextRankPixel
            If d > finalY Then GoTo NextRankPixel
        
            r = srcImageData(QuickValInner + 2, d)
            g = srcImageData(QuickValInner + 1, d)
            b = srcImageData(QuickValInner, d)
            
            grayVal = gLookup(r + g + b)
            
            Select Case RankType
                Case 0
                    If grayVal > MaxTotal Then
                        MaxTotal = grayVal
                        MaxX = c
                        MaxY = d
                    End If
                Case 1
                    If grayVal < MaxTotal Then
                        MaxTotal = grayVal
                        MaxX = c
                        MaxY = d
                    End If
                Case 2
                    grayVal = Abs(grayValOriginal - grayVal)
                    If grayVal > MaxTotal Then
                        MaxTotal = grayVal
                        MaxX = c
                        MaxY = d
                    End If
            End Select

NextRankPixel:
        Next d
        Next c
    
        QuickValDst = MaxX * qvDepth
        
        'Assign that ranked value to each color channel
        dstImageData(QuickVal + 2, y) = srcImageData(QuickValDst + 2, MaxY)
        dstImageData(QuickVal + 1, y) = srcImageData(QuickValDst + 1, MaxY)
        dstImageData(QuickVal, y) = srcImageData(QuickValDst, MaxY)
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic

End Sub

Private Sub Form_Activate()
   
    'Note the current image's width and height, which will be needed to adjust the preview effect
    iWidth = pdImages(CurrentImage).Width
    iHeight = pdImages(CurrentImage).Height
    
    'Possible methods of calculating rank filters:
    cboRank.AddItem "Maximum (Dilate)", 0
    cboRank.AddItem "Minimum (Erode)", 1
    cboRank.AddItem "Extreme (Furthest value)", 2
    
    'Make "Maximum" the default value
    cboRank.ListIndex = 0
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Create a preview
    CustomRankFilter hsRadius.Value, cboRank.ListIndex, True, fxPreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub hsRadius_Change()
    copyToTextBoxI txtRadius, hsRadius.Value
    CustomRankFilter hsRadius.Value, cboRank.ListIndex, True, fxPreview
End Sub

Private Sub hsRadius_Scroll()
    copyToTextBoxI txtRadius, hsRadius.Value
    CustomRankFilter hsRadius.Value, cboRank.ListIndex, True, fxPreview
End Sub

Private Sub txtRadius_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtRadius
    If EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, False, False) Then hsRadius.Value = Val(txtRadius)
End Sub

Private Sub txtRadius_GotFocus()
    AutoSelectText txtRadius
End Sub
