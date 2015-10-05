VERSION 5.00
Begin VB.Form FormRechannel 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Rechannel"
   ClientHeight    =   6540
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   9450
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
   ScaleWidth      =   630
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   5
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "red"
      Value           =   -1  'True
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
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   6
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "green"
   End
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   2
      Left            =   6840
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "blue"
   End
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   3
      Left            =   6840
      TabIndex        =   8
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "cyan"
   End
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   4
      Left            =   6840
      TabIndex        =   9
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "magenta"
   End
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   10
      Top             =   2880
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "yellow"
   End
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   6
      Left            =   6840
      TabIndex        =   11
      Top             =   3840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "cyan"
   End
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   7
      Left            =   6840
      TabIndex        =   12
      Top             =   4200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "magenta"
   End
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   8
      Left            =   6840
      TabIndex        =   13
      Top             =   4560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "yellow"
   End
   Begin PhotoDemon.smartOptionButton optChannel 
      Height          =   375
      Index           =   9
      Left            =   6840
      TabIndex        =   14
      Top             =   4920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Caption         =   "key (black)"
   End
   Begin VB.Label lblCMYK 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMYK channels"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   3480
      Width           =   1605
   End
   Begin VB.Label lblCMY 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMY channels"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   1800
      Width           =   1470
   End
   Begin VB.Label lblRGB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RGB channels"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "FormRechannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Rechannel Interface
'Copyright 2000-2015 by Tanner Helland
'Created: original rechannel algorithm - sometimes 2001, this form 28/September/12
'Last updated: 24/August/13
'Last update: added command bar
'
'Rechannel (or "channel isolation") tool.  This allows the user to isolate a single color channel from
' the RGB and CMY/CMYK color spaces.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub cmdBar_OKClick()
    
    Dim i As Long
    For i = 0 To optChannel.Count - 1
        If optChannel(i) Then Process "Rechannel", , buildParams(i), UNDO_LAYER
    Next i
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Render a preview
    updatePreview
    
End Sub


'Rechannel an image (red, green, blue, cyan, magenta, yellow)
Public Sub RechannelImage(ByVal rType As Byte, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    'Based on the channel the user has selected, display a user-friendly description of this filter
    Dim cName As String
    Select Case rType
        Case 0
            cName = g_Language.TranslateMessage("red")
        Case 1
            cName = g_Language.TranslateMessage("green")
        Case 2
            cName = g_Language.TranslateMessage("blue")
        Case 3
            cName = g_Language.TranslateMessage("cyan")
        Case 4
            cName = g_Language.TranslateMessage("magenta")
        Case 5
            cName = g_Language.TranslateMessage("yellow")
        Case 6
            cName = g_Language.TranslateMessage("cyan")
        Case 7
            cName = g_Language.TranslateMessage("magenta")
        Case 8
            cName = g_Language.TranslateMessage("yellow")
        Case 9
            cName = g_Language.TranslateMessage("black")
    End Select
        
    If toPreview = False Then Message "Isolating the %1 channel...", cName
    
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
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    Dim cK As Double, MK As Double, yK As Double, bK As Double, invBK As Double
    
    'After all that work, the Rechannel code itself is relatively small and unexciting!
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        Select Case rType
            'Rechannel red
            Case 0
                ImageData(QuickVal, y) = 0
                ImageData(QuickVal + 1, y) = 0
            'Rechannel green
            Case 1
                ImageData(QuickVal, y) = 0
                ImageData(QuickVal + 2, y) = 0
            'Rechannel blue
            Case 2
                ImageData(QuickVal + 1, y) = 0
                ImageData(QuickVal + 2, y) = 0
            'Rechannel cyan
            Case 3
                ImageData(QuickVal, y) = 255
                ImageData(QuickVal + 1, y) = 255
            'Rechannel magenta
            Case 4
                ImageData(QuickVal, y) = 255
                ImageData(QuickVal + 2, y) = 255
            'Rechannel yellow
            Case 5
                ImageData(QuickVal + 1, y) = 255
                ImageData(QuickVal + 2, y) = 255
            
            'Rechannel CMYK
            Case Else
                cK = 255 - ImageData(QuickVal + 2, y)
                MK = 255 - ImageData(QuickVal + 1, y)
                yK = 255 - ImageData(QuickVal, y)
                
                cK = cK / 255
                MK = MK / 255
                yK = yK / 255
                
                bK = Min3Float(cK, MK, yK)
    
                invBK = 1 - bK
                If invBK = 0 Then invBK = 0.0001
                
                If rType = 6 Then
                    cK = ((cK - bK) / invBK) * 255
                    ImageData(QuickVal + 2, y) = 255 - cK
                    ImageData(QuickVal + 1, y) = 255
                    ImageData(QuickVal, y) = 255
                End If
                
                If rType = 7 Then
                    MK = ((MK - bK) / invBK) * 255
                    ImageData(QuickVal + 2, y) = 255
                    ImageData(QuickVal + 1, y) = 255 - MK
                    ImageData(QuickVal, y) = 255
                End If
                
                If rType = 8 Then
                    yK = ((yK - bK) / invBK) * 255
                    ImageData(QuickVal + 2, y) = 255
                    ImageData(QuickVal + 1, y) = 255
                    ImageData(QuickVal, y) = 255 - yK
                End If
                
                If rType = 9 Then
                    ImageData(QuickVal + 2, y) = invBK * 255
                    ImageData(QuickVal + 1, y) = invBK * 255
                    ImageData(QuickVal, y) = invBK * 255
                End If
                
        End Select
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

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub optChannel_Click(Index As Integer)
    updatePreview
End Sub

Private Sub updatePreview()
    
    If cmdBar.previewsAllowed Then
    
        Dim i As Long
        For i = 0 To optChannel.Count - 1
            If optChannel(i) Then RechannelImage i, True, fxPreview
        Next i
        
    End If
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

