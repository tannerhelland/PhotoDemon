VERSION 5.00
Begin VB.Form FormRechannel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Rechannel"
   ClientHeight    =   6630
   ClientLeft      =   -15
   ClientTop       =   225
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
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
      Caption         =   "key (black)"
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
      Height          =   300
      Index           =   9
      Left            =   4440
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
   End
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
      Caption         =   "yellow"
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
      Height          =   300
      Index           =   8
      Left            =   4440
      TabIndex        =   9
      Top             =   4800
      Width           =   1695
   End
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
      Caption         =   "magenta"
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
      Height          =   300
      Index           =   7
      Left            =   4440
      TabIndex        =   8
      Top             =   4320
      Width           =   1695
   End
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
      Caption         =   "cyan"
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
      Height          =   300
      Index           =   6
      Left            =   4440
      TabIndex        =   7
      Top             =   3840
      Width           =   1695
   End
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
      Caption         =   "yellow"
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
      Height          =   300
      Index           =   5
      Left            =   2280
      TabIndex        =   6
      Top             =   4800
      Width           =   1815
   End
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
      Caption         =   "magenta"
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
      Height          =   300
      Index           =   4
      Left            =   2280
      TabIndex        =   5
      Top             =   4320
      Width           =   1695
   End
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
      Caption         =   "cyan"
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
      Height          =   300
      Index           =   3
      Left            =   2280
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   120
      Width           =   2895
   End
   Begin VB.OptionButton OptChannel 
      Appearance      =   0  'Flat
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
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   6000
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   6000
      Width           =   1245
   End
   Begin VB.Label lblCMYK 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMYK channels:"
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
      Left            =   4320
      TabIndex        =   18
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblCMY 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CMY channels:"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   3360
      Width           =   1560
   End
   Begin VB.Label lblRGB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RGB channels:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   1530
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   2880
      Width           =   480
   End
End
Attribute VB_Name = "FormRechannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Rechannel Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: original rechannel algorithm - sometimes 2001, this form 28/September/12
'Last updated: 28/September/12
'Last update: built a dedicated form for rechanneling, added CMY options
'
'Rechannel (or "channel isolation") tool.  This allows the user to isolate a single color channel from
' the RGB and CMY color spaces.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    Me.Visible = False
    
    Dim rechannelMethod As Long
    
    If OptChannel(0) Then rechannelMethod = 0
    If OptChannel(1) Then rechannelMethod = 1
    If OptChannel(2) Then rechannelMethod = 2
    If OptChannel(3) Then rechannelMethod = 3
    If OptChannel(4) Then rechannelMethod = 4
    If OptChannel(5) Then rechannelMethod = 5
    If OptChannel(6) Then rechannelMethod = 6
    If OptChannel(7) Then rechannelMethod = 7
    If OptChannel(8) Then rechannelMethod = 8
    If OptChannel(9) Then rechannelMethod = 9
    
    Process Rechannel, rechannelMethod
    
    Unload Me
End Sub

Private Sub Form_Activate()
    
    'Create the image previews
    DrawPreviewImage picPreview
    RechannelImage 0, True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub


'Rechannel an image (red, green, blue, cyan, magenta, yellow)
Public Sub RechannelImage(ByVal rType As Byte, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    'Based on the channel the user has selected, display a user-friendly description of this filter
    Dim cName As String
    Select Case rType
        Case 0
            cName = "red"
        Case 1
            cName = "green"
        Case 2
            cName = "blue"
        Case 3
            cName = "cyan"
        Case 4
            cName = "magenta"
        Case 5
            cName = "yellow"
        Case 6
            cName = "cyan"
        Case 7
            cName = "magenta"
        Case 8
            cName = "yellow"
        Case 9
            cName = "black"
    End Select
    
    If toPreview = False Then Message "Isolating the " & cName & " channel..."
    
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
    
    Dim cK As Single, mK As Single, yK As Single, bK As Single, invBK As Single
    
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
                mK = 255 - ImageData(QuickVal + 1, y)
                yK = 255 - ImageData(QuickVal, y)
                
                cK = cK / 255
                mK = mK / 255
                yK = yK / 255
                
                bK = Minimum(cK, mK, yK)
    
                invBK = 1 - bK
                If invBK = 0 Then invBK = 0.0001
                
                If rType = 6 Then
                    cK = ((cK - bK) / invBK) * 255
                    ImageData(QuickVal + 2, y) = 255 - cK
                    ImageData(QuickVal + 1, y) = 255
                    ImageData(QuickVal, y) = 255
                End If
                
                If rType = 7 Then
                    mK = ((mK - bK) / invBK) * 255
                    ImageData(QuickVal + 2, y) = 255
                    ImageData(QuickVal + 1, y) = 255 - mK
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

Private Sub optChannel_Click(Index As Integer)
    RechannelImage Index, True, picEffect
End Sub
