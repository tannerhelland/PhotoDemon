VERSION 5.00
Begin VB.Form FormEmbossEngrave 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Emboss/Engrave"
   ClientHeight    =   6165
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
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   5550
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4710
      TabIndex        =   1
      Top             =   5550
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.OptionButton OptEmboss 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "emboss"
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
      Height          =   345
      Left            =   1560
      TabIndex        =   2
      Top             =   3420
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton OptEngrave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "engrave"
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
      Height          =   345
      Left            =   3480
      TabIndex        =   3
      Top             =   3420
      Width           =   1335
   End
   Begin VB.PictureBox PicColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      ScaleHeight     =   465
      ScaleWidth      =   5505
      TabIndex        =   5
      Top             =   4560
      Width           =   5535
   End
   Begin VB.CheckBox ChkToColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "use custom background color (click colored box to change)..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   5655
   End
   Begin VB.Label lblBackground 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   -840
      TabIndex        =   10
      Top             =   5400
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   2880
      Width           =   480
   End
End
Attribute VB_Name = "FormEmbossEngrave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Emboss/Engrave Filter Interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 3/6/03
'Last updated: 09/September/12
'Last update: rewrite emboss/engrave against new layer class
'
'Module for handling all emboss and engrave filters.  It's basically just an
'interfacing layer to the 4 main filters: Emboss/EmbossToColor and Engrave/EngraveToColor
'
'***************************************************************************

Option Explicit

Private Sub ChkToColor_Click()
    If OptEmboss.Value = True Then
        If ChkToColor.Value = vbChecked Then FilterEmbossColor PicColor.backColor, True, picEffect Else FilterEmbossColor RGB(127, 127, 127), True, picEffect
    Else
        If ChkToColor.Value = vbChecked Then FilterEngraveColor PicColor.backColor, True, picEffect Else FilterEngraveColor RGB(127, 127, 127), True, picEffect
    End If
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    
    'Used to remember the last color used for embossing
    g_EmbossEngraveColor = PicColor.backColor
    Me.Visible = False
    
    'Dependent: filter to grey OR to a background color
    If OptEmboss.Value = True Then
        If ChkToColor.Value = vbChecked Then Process EmbossToColor, PicColor.backColor Else Process EmbossToColor, RGB(127, 127, 127)
    Else
        If ChkToColor.Value = vbChecked Then Process EngraveToColor, PicColor.backColor Else Process EngraveToColor, RGB(127, 127, 127)
    End If
    
    Unload Me
End Sub

Private Sub Form_Activate()
    
    'Remember the last emboss/engrave color selection
    PicColor.backColor = g_EmbossEngraveColor
    
    'Draw a preview of the current image to the left box
    DrawPreviewImage picPreview
    
    'Draw a preview of the emboss/engrave effect to the right box
    If OptEmboss.Value = True Then
        If ChkToColor.Value = vbChecked Then FilterEmbossColor PicColor.backColor, True, picEffect Else FilterEmbossColor RGB(127, 127, 127), True, picEffect
    Else
        If ChkToColor.Value = vbChecked Then FilterEngraveColor PicColor.backColor, True, picEffect Else FilterEngraveColor RGB(127, 127, 127), True, picEffect
    End If
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    setHandCursor PicColor
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When the emboss/engrave options are clicked, redraw the preview
Private Sub OptEmboss_Click()
    If OptEmboss.Value = True Then
        If ChkToColor.Value = vbChecked Then FilterEmbossColor PicColor.backColor, True, picEffect Else FilterEmbossColor RGB(127, 127, 127), True, picEffect
    Else
        If ChkToColor.Value = vbChecked Then FilterEngraveColor PicColor.backColor, True, picEffect Else FilterEngraveColor RGB(127, 127, 127), True, picEffect
    End If
End Sub

Private Sub OptEngrave_Click()
    If OptEmboss.Value = True Then
        If ChkToColor.Value = vbChecked Then FilterEmbossColor PicColor.backColor, True, picEffect Else FilterEmbossColor RGB(127, 127, 127), True, picEffect
    Else
        If ChkToColor.Value = vbChecked Then FilterEngraveColor PicColor.backColor, True, picEffect Else FilterEngraveColor RGB(127, 127, 127), True, picEffect
    End If
End Sub

'Clicking on the picture box allows the user to select a new color
Private Sub PicColor_Click()

    'Use a common dialog box to select a new color.  (In the future, perhaps I'll design a better custom box.)
    Dim retColor As Long
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    retColor = PicColor.backColor
    
    If CD1.VBChooseColor(retColor, True, True, False, Me.hWnd) Then
        PicColor.backColor = retColor
        ChkToColor.Value = vbChecked
        If OptEmboss.Value = True Then
            If ChkToColor.Value = vbChecked Then FilterEmbossColor PicColor.backColor, True, picEffect Else FilterEmbossColor RGB(127, 127, 127), True, picEffect
        Else
            If ChkToColor.Value = vbChecked Then FilterEngraveColor PicColor.backColor, True, picEffect Else FilterEngraveColor RGB(127, 127, 127), True, picEffect
        End If
    End If
    
End Sub

'Emboss an image
' Inputs: color to emboss to, and whether or not this is a preview (plus the destination picture box if it IS a preview)
Public Sub FilterEmbossColor(ByVal cColor As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Embossing image..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already embossed pixels from screwing up our results for later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right - 1
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValRight As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tB As Long, tG As Long

    'Extract the red, green, and blue values from the color we've been passed
    tR = ExtractR(cColor)
    tG = ExtractG(cColor)
    tB = ExtractB(cColor)
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
        QuickValRight = (x + 1) * qvDepth
    For y = initY To finalY
    
        'This line is the emboss code.  Very simple, very fast.
        r = Abs(CLng(srcImageData(QuickVal + 2, y)) - CLng(srcImageData(QuickValRight + 2, y)) + tR)
        g = Abs(CLng(srcImageData(QuickVal + 1, y)) - CLng(srcImageData(QuickValRight + 1, y)) + tG)
        b = Abs(CLng(srcImageData(QuickVal, y)) - CLng(srcImageData(QuickValRight, y)) + tB)
        
        If r > 255 Then
            r = 255
        ElseIf r < 0 Then
            r = 0
        End If
        
        If g > 255 Then
            g = 255
        ElseIf g < 0 Then
            g = 0
        End If
        
        If b > 255 Then
            b = 255
        ElseIf b < 0 Then
            b = 0
        End If

        dstImageData(QuickVal + 2, y) = r
        dstImageData(QuickVal + 1, y) = g
        dstImageData(QuickVal, y) = b
        
        'The right-most line of pixels will always be missed, so manually check for and correct that
        If x = finalX Then
            dstImageData(QuickValRight + 2, y) = r
            dstImageData(QuickValRight + 1, y) = g
            dstImageData(QuickValRight, y) = b
        End If
        
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

'Engrave an image
' Inputs: color to emboss to, and whether or not this is a preview (plus the destination picture box if it IS a preview)
Public Sub FilterEngraveColor(ByVal cColor As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Embossing image..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already engraved pixels from screwing up our results for later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right - 1
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValRight As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tB As Long, tG As Long
    
    'Extract the red, green, and blue values from the color we've been passed
    tR = ExtractR(cColor)
    tG = ExtractG(cColor)
    tB = ExtractB(cColor)
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
        QuickValRight = (x + 1) * qvDepth
    For y = initY To finalY
    
        'This line is the emboss code.  Very simple, very fast.
        r = Abs(CLng(srcImageData(QuickValRight + 2, y)) - CLng(srcImageData(QuickVal + 2, y)) + tR)
        g = Abs(CLng(srcImageData(QuickValRight + 1, y)) - CLng(srcImageData(QuickVal + 1, y)) + tG)
        b = Abs(CLng(srcImageData(QuickValRight, y)) - CLng(srcImageData(QuickVal, y)) + tB)
        
        If r > 255 Then
            r = 255
        ElseIf r < 0 Then
            r = 0
        End If
        
        If g > 255 Then
            g = 255
        ElseIf g < 0 Then
            g = 0
        End If
        
        If b > 255 Then
            b = 255
        ElseIf b < 0 Then
            b = 0
        End If

        dstImageData(QuickVal + 2, y) = r
        dstImageData(QuickVal + 1, y) = g
        dstImageData(QuickVal, y) = b
        
        'The right-most line of pixels will always be missed, so manually check for and correct that
        If x = finalX Then
            dstImageData(QuickValRight + 2, y) = r
            dstImageData(QuickValRight + 1, y) = g
            dstImageData(QuickValRight, y) = b
        End If
        
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
