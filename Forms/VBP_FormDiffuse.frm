VERSION 5.00
Begin VB.Form FormDiffuse 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Custom Diffuse"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
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
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsY 
      Height          =   255
      Left            =   240
      Max             =   10
      MouseIcon       =   "VBP_FormDiffuse.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4320
      Value           =   5
      Width           =   4575
   End
   Begin VB.HScrollBar hsX 
      Height          =   255
      Left            =   240
      Max             =   10
      MouseIcon       =   "VBP_FormDiffuse.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3240
      Value           =   5
      Width           =   4575
   End
   Begin VB.TextBox txtX 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Text            =   "0"
      Top             =   2730
      Width           =   495
   End
   Begin VB.TextBox txtY 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Text            =   "0"
      Top             =   3810
      Width           =   495
   End
   Begin VB.PictureBox PicEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   2175
      Left            =   2640
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox PicPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   2175
      Left            =   240
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox ChkWrap 
      Appearance      =   0  'Flat
      Caption         =   " Wrap edge values"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1680
      MouseIcon       =   "VBP_FormDiffuse.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      MouseIcon       =   "VBP_FormDiffuse.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5400
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      MouseIcon       =   "VBP_FormDiffuse.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   5400
      Width           =   1125
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  Before                                           After"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2310
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max Y Distance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   1080
      TabIndex        =   8
      Top             =   3840
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max X Distance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
      Width           =   1290
   End
End
Attribute VB_Name = "FormDiffuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Diffuse Filter Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 8/14/01
'Last updated: 5/May/07
'Last update: Fixed grabbing of off-image pixels when previewing and erroneous
'             values of preview when the form is first loaded
'
'Module for handling the diffusion-style filters.  Automates both saturated
'and wrapped diffusion.
'
'***************************************************************************

Option Explicit

'Arrays necessary for running diffusion filters properly
Dim TmpBitmapArray() As Long
Dim CalcArray() As Long

Private Sub ChkWrap_Click()
    DrawPreviewDiffuse (hsX.Value / PicWidthL) * PreviewWidth, (hsY.Value / PicHeightL) * PreviewHeight
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Public Sub Diffuse()
    BuildTmpArray
    Randomize Timer
    Message "Diffusing image..."
    SetProgBarMax PicWidthL
    Dim DiffuseX As Integer, DiffuseY As Integer
    For x = 0 To PicWidthL
        For y = 0 To PicHeightL
            DiffuseX = Rnd * 3 - 1.5
            DiffuseY = Rnd * 3 - 1.5
            If DiffuseX + x < 0 Then DiffuseX = 0
            If DiffuseY + y < 0 Then DiffuseY = 0
            If DiffuseX + x > PicWidthL - 1 Then DiffuseX = 0
            If DiffuseY + y > PicHeightL - 1 Then DiffuseY = 0
            TmpBitmapArray(x, y) = CalcArray(DiffuseX + x, DiffuseY + y)
        Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetTmpArray
    Message "Image diffused."
End Sub

Public Sub DiffuseMore()
    BuildTmpArray
    Randomize Timer
    Message "Diffusing image..."
    SetProgBarMax PicWidthL
    Dim DiffuseX As Integer, DiffuseY As Integer
    For x = 0 To PicWidthL
        For y = 0 To PicHeightL
            DiffuseX = Rnd * 6 - 3
            DiffuseY = Rnd * 6 - 3
            If DiffuseX + x < 0 Then DiffuseX = 0
            If DiffuseY + y < 0 Then DiffuseY = 0
            If DiffuseX + x > PicWidthL - 1 Then DiffuseX = 0
            If DiffuseY + y > PicHeightL - 1 Then DiffuseY = 0
            TmpBitmapArray(x, y) = CalcArray(DiffuseX + x, DiffuseY + y)
        Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetTmpArray
    Message "Image diffused."
End Sub

'OK button
Private Sub CmdOK_Click()
    'The max and min values of the scroll bars are used to validate the range of the text box
    If EntryValid(txtX, hsX.Min, hsX.Max) And EntryValid(txtY, hsY.Min, hsY.Max) Then
        FormDiffuse.Visible = False
        Process CustomDiffuse, hsX.Value, hsY.Value, , , False
        Unload Me
    End If
End Sub

Public Sub DiffuseCustom(ByVal xDiffuse As Long, ByVal yDiffuse As Long)
    GetImageData
    BuildTmpArray
    Randomize Timer
    Message "Diffusing image..."
    SetProgBarMax PicWidthL
    Dim DiffuseX As Integer, DiffuseY As Integer
    Dim dx As Long, dy As Long
    Dim HDX As Single, HDY As Single
    dx = xDiffuse
    dy = yDiffuse
    HDX = dx / 2
    HDY = dy / 2
    Dim dstX As Long, dstY As Long
    Dim tPicWI As Long, tPicHi As Long
    tPicWI = PicWidthL - 1
    tPicHi = PicHeightL - 1
    'Diffusion for unwrapped data
    If ChkWrap.Value = vbUnchecked Then
        For x = 0 To PicWidthL
            For y = 0 To PicHeightL
                DiffuseX = Rnd * dx - HDX
                DiffuseY = Rnd * dy - HDY
                dstX = DiffuseX + x
                dstY = DiffuseY + y
                If dstX < 0 Then dstX = 0
                If dstY < 0 Then dstY = 0
                If dstX > tPicWI Then dstX = tPicWI
                If dstY > tPicHi Then dstY = tPicHi
                TmpBitmapArray(x, y) = CalcArray(dstX, dstY)
            Next y
            If x Mod 20 = 0 Then SetProgBarVal x
        Next x
    Else
    
    'Diffusion for wrapped data
        For x = 0 To PicWidthL
            For y = 0 To PicHeightL
                DiffuseX = Rnd * dx - HDX
                DiffuseY = Rnd * dy - HDY
                dstX = DiffuseX + x
                dstY = DiffuseY + y
                If dstX < 0 Then dstX = PicWidthL + dstX
                If dstY < 0 Then dstY = PicHeightL + dstY
                If dstX > tPicWI Then dstX = dstX - PicWidthL
                If dstY > tPicHi Then dstY = dstY - PicHeightL
                TmpBitmapArray(x, y) = CalcArray(dstX, dstY)
            Next y
            If x Mod 20 = 0 Then SetProgBarVal x
        Next x
    End If
    SetTmpArray
    Message "Image diffused."
End Sub

Private Sub BuildTmpArray()
    'Copy the data into an easier format
    Message "Gathering image information..."
    ReDim CalcArray(0 To PicWidthL, 0 To PicHeightL) As Long
    ReDim TmpBitmapArray(0 To PicWidthL, 0 To PicHeightL) As Long
    Dim tX As Long
    For x = 0 To PicWidthL
        tX = x * 3
    For y = 0 To PicHeightL
        TmpBitmapArray(x, y) = RGB(ImageData(tX + 2, y), ImageData(tX + 1, y), ImageData(tX, y))
        CalcArray(x, y) = TmpBitmapArray(x, y)
    Next y
    Next x
End Sub

Public Sub SetTmpArray()
    'Copy the information from the temporary arrays back into the main one
    SetProgBarVal cProgBar.Max
    Dim tX As Long
    Dim TV As Long
    For x = 0 To PicWidthL
        tX = x * 3
    For y = 0 To PicHeightL
        TV = TmpBitmapArray(x, y)
        ImageData(tX + 2, y) = CByte(ExtractR(TV))
        ImageData(tX + 1, y) = CByte(ExtractG(TV))
        ImageData(tX, y) = CByte(ExtractB(TV))
    Next y
    Next x
    SetImageData
End Sub

Private Sub Form_Load()
    
    GetImageData
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    hsX.Max = PicWidthL
    hsY.Max = PicHeightL
    hsX.Value = PicWidthL \ 2
    hsY.Value = PicHeightL \ 2
    DoEvents
    
    'Draw preview effect
    DrawPreviewDiffuse (hsX.Value / PicWidthL) * PreviewWidth, (hsY.Value / PicHeightL) * PreviewHeight
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'Diffuse preview
Private Sub DrawPreviewDiffuse(ByVal xDiffuse As Long, ByVal yDiffuse As Long)
    GetPreviewData PicPreview

    ReDim CalcArray(0 To PreviewWidth + PreviewX * 2, 0 To PreviewHeight + PreviewY * 2) As Long
    ReDim TmpBitmapArray(0 To PreviewWidth + PreviewX * 2, 0 To PreviewHeight + PreviewY * 2) As Long
    Dim tX As Long
    For x = PreviewX To PreviewX + PreviewWidth
        tX = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        TmpBitmapArray(x, y) = RGB(ImageData(tX + 2, y), ImageData(tX + 1, y), ImageData(tX, y))
        CalcArray(x, y) = TmpBitmapArray(x, y)
    Next y
    Next x

    Randomize Timer
    
    Dim DiffuseX As Integer, DiffuseY As Integer
    Dim dx As Long, dy As Long
    Dim HDX As Single, HDY As Single
    dx = xDiffuse
    dy = yDiffuse
    HDX = dx / 2
    HDY = dy / 2
    Dim dstX As Long, dstY As Long
    Dim tPicWI As Long, tPicHi As Long
    tPicWI = PreviewX + PreviewWidth
    tPicHi = PreviewY + PreviewHeight
    'Diffusion for unwrapped data
    If ChkWrap.Value = vbUnchecked Then
        For x = PreviewX To PreviewX + PreviewWidth
            For y = PreviewY To PreviewY + PreviewHeight
                DiffuseX = Rnd * dx - HDX
                DiffuseY = Rnd * dy - HDY
                dstX = DiffuseX + x
                dstY = DiffuseY + y
                If dstX < PreviewX Then dstX = PreviewX
                If dstY < PreviewY Then dstY = PreviewY
                If dstX > tPicWI Then dstX = PreviewX + PreviewWidth
                If dstY > tPicHi Then dstY = PreviewY + PreviewHeight
                TmpBitmapArray(x, y) = CalcArray(dstX, dstY)
            Next y
        Next x
    Else
    
    'Diffusion for wrapped data
        For x = PreviewX To PreviewX + PreviewWidth
            For y = PreviewY To PreviewY + PreviewHeight
                DiffuseX = Rnd * dx - HDX
                DiffuseY = Rnd * dy - HDY
                dstX = DiffuseX + x
                dstY = DiffuseY + y
                If dstX < PreviewX Then dstX = PreviewWidth + PreviewX + dstX
                If dstY < PreviewY Then dstY = PreviewHeight + PreviewY + dstY
                If dstX > tPicWI Then dstX = dstX - PreviewX - PreviewWidth
                If dstY > tPicHi Then dstY = dstY - PreviewY - PreviewHeight
                TmpBitmapArray(x, y) = CalcArray(dstX, dstY)
            Next y
        Next x
    End If
    
    Dim TV As Long
    
    For x = PreviewX To PreviewX + PreviewWidth
        tX = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
        TV = TmpBitmapArray(x, y)
        ImageData(tX + 2, y) = (TV And 255)
        ImageData(tX + 1, y) = (TV \ 256) And 255
        ImageData(tX, y) = (TV \ 65536) And 255
    Next y
    Next x
    SetPreviewData PicEffect
    
End Sub

'Everything below this line relates to mirroring the input of the textboxes across the scrollbars (and vice versa)
Private Sub hsX_Change()
    txtX.Text = hsX.Value
    DrawPreviewDiffuse (hsX.Value / PicWidthL) * PreviewWidth, (hsY.Value / PicHeightL) * PreviewHeight
End Sub

Private Sub hsX_Scroll()
    txtX.Text = hsX.Value
    DrawPreviewDiffuse (hsX.Value / PicWidthL) * PreviewWidth, (hsY.Value / PicHeightL) * PreviewHeight
End Sub

Private Sub hsY_Change()
    txtY.Text = hsY.Value
    DrawPreviewDiffuse (hsX.Value / PicWidthL) * PreviewWidth, (hsY.Value / PicHeightL) * PreviewHeight
End Sub

Private Sub hsY_Scroll()
    txtY.Text = hsY.Value
    DrawPreviewDiffuse (hsX.Value / PicWidthL) * PreviewWidth, (hsY.Value / PicHeightL) * PreviewHeight
End Sub

Private Sub txtX_Change()
    If EntryValid(txtX, hsX.Min, hsX.Max, False, False) Then hsX.Value = val(txtX)
End Sub

Private Sub txtX_GotFocus()
    AutoSelectText txtX
End Sub

Private Sub txtY_Change()
    If EntryValid(txtY, hsY.Min, hsY.Max, False, False) Then hsY.Value = val(txtY)
End Sub

Private Sub txtY_GotFocus()
    AutoSelectText txtY
End Sub
