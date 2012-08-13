VERSION 5.00
Begin VB.Form FormMosaic 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Mosaic Options"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5100
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
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.HScrollBar hsHeight 
      Height          =   255
      Left            =   240
      Max             =   64
      Min             =   2
      TabIndex        =   3
      Top             =   4320
      Value           =   2
      Width           =   4575
   End
   Begin VB.HScrollBar hsWidth 
      Height          =   255
      Left            =   240
      Max             =   64
      Min             =   2
      TabIndex        =   1
      Top             =   3240
      Value           =   2
      Width           =   4575
   End
   Begin VB.TextBox txtHeight 
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
      Text            =   "2"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtWidth 
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
      Text            =   "2"
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox PicPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   2175
      Left            =   240
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox PicEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   2175
      Left            =   2640
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3720
      TabIndex        =   5
      Top             =   4920
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   2520
      TabIndex        =   4
      Top             =   4920
      Width           =   1125
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Block Width:"
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
      Left            =   1320
      TabIndex        =   10
      Top             =   2790
      Width           =   1035
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Block Height:"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   3870
      Width           =   1080
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
      TabIndex        =   8
      Top             =   2310
      Width           =   4575
   End
End
Attribute VB_Name = "FormMosaic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Mosaic filter interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 8/5/00
'Last updated: 4/November/07
'Last update: preview no longer gets "Divide by 0" errors for small images
'
'Form for handling all the mosaic image transform code.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    If EntryValid(txtWidth, hsWidth.Min, hsWidth.Max) And EntryValid(txtHeight, hsHeight.Min, hsHeight.Max) Then
        Me.Visible = False
        Process Mosaic, hsWidth.Value, hsHeight.Value
        Unload Me
    End If
End Sub

Public Sub MosaicFilter(ByVal BlockSizeX As Byte, ByVal BlockSizeY As Byte)
    
    'Used for the for..next loops
    Dim XLoop As Long, YLoop As Long
    Dim DstXLoop As Long, DstYLoop As Long
    Dim InitXLoop As Long, InitYLoop As Long
    
    'How many pixels must be averaged
    Dim NumOfPixels As Long
    
    'Holds the RGB data for the mosaic
    Dim tR As Long, tG As Long, tB As Long
    
    GetImageData True
    
    'Calculate how many mosaic tiles will fit on the current image's size
    XLoop = Int(PicWidthL \ BlockSizeX) + 1
    YLoop = Int(PicHeightL \ BlockSizeY) + 1
    
    'Store the pixel data into an array for faster accessing
    Dim PicArray() As Byte
    ReDim PicArray(0 To (PicWidthL + 1) * 3, 0 To PicHeightL + 1) As Byte
    
    Dim a As Long, b As Long
    
    Message "Preparing transfer array..."
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    For z = 0 To 2
        PicArray(QuickVal + z, y) = ImageData(QuickVal + z, y)
    Next z
    Next y
    Next x
    
    Message "Generating mosaic image..."
    SetProgBarMax XLoop
    
    'Begin the main mosaic loop
    For x = 0 To XLoop
    For y = 0 To YLoop
    
        'This sub loop is to gather all of the data for the current mosaic tile
        InitXLoop = x * BlockSizeX
        InitYLoop = y * BlockSizeY
        DstXLoop = (x + 1) * BlockSizeX - 1
        DstYLoop = (y + 1) * BlockSizeY - 1
        For a = InitXLoop To DstXLoop
            QuickVal = a * 3
        For b = InitYLoop To DstYLoop
        
            'If this particular pixel is off of the image, don't bother counting it
            If a > PicWidthL Or b > PicHeightL Then GoTo NextMosiacPixel1
            
            'Total up all the red, green, and blue values for the pixels within this
            'mosiac tile
            tR = tR + PicArray(QuickVal + 2, b)
            tG = tG + PicArray(QuickVal + 1, b)
            tB = tB + PicArray(QuickVal, b)
            
            'Count this as a valid pixel
            NumOfPixels = NumOfPixels + 1
            
NextMosiacPixel1:
        
        Next b
        Next a
        
        'If this tile is completely off of the image, don't worry about it and go to the next one
        If NumOfPixels = 0 Then GoTo NextMosaicPixel3
        
        'Take the average red, green, and blue values of all the pixles within this tile
        tR = tR \ NumOfPixels
        tG = tG \ NumOfPixels
        tB = tB \ NumOfPixels
        
        'Now run a loop through the same pixels you just analyzed, only this time you're gonna
        'draw the averaged color over the top of them
        For a = InitXLoop To DstXLoop
            QuickVal = a * 3
        For b = InitYLoop To DstYLoop
        
            'Same thing as above - if it's off the image, ignore it
            If a > PicWidthL Or b > PicHeightL Then GoTo NextMosiacPixel2
            
            'Set the pixel
            ImageData(QuickVal + 2, b) = tR
            ImageData(QuickVal + 1, b) = tG
            ImageData(QuickVal, b) = tB
            
NextMosiacPixel2:

        Next b
        Next a

NextMosaicPixel3:

        'Clear all the variables and go to the next pixel
        tR = 0
        tG = 0
        tB = 0
        NumOfPixels = 0
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    
    SetImageData True
    
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Create the previews
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    
    hsWidth.Max = PicWidthL
    hsHeight.Max = PicHeightL
    PreviewMosaicFilter (CSng(hsWidth.Value) / hsWidth.Max) * PreviewWidth, (CSng(hsHeight.Value) / hsHeight.Max) * PreviewHeight
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

Private Sub PreviewMosaicFilter(ByVal BlockSizeX As Byte, ByVal BlockSizeY As Byte)
    
    'Preview only
    GetPreviewData PicPreview
    
    'Used for the for..next loops
    Dim XLoop As Long, YLoop As Long
    Dim DstXLoop As Long, DstYLoop As Long
    Dim InitXLoop As Long, InitYLoop As Long
    'How many pixels must be averaged
    Dim NumOfPixels As Long
    'Holds the RGB data for the mosaic
    Dim tR As Long, tG As Long, tB As Long
    
    'Fix the BlockSizeX & Y variables to never be less than 1
    If BlockSizeX < 1 Then BlockSizeX = 1
    If BlockSizeY < 1 Then BlockSizeY = 1
    
    'Calculate how many mosaic tiles will fit on the current image's size
    XLoop = Int(PreviewWidth \ BlockSizeX) + 1
    YLoop = Int(PreviewHeight \ BlockSizeY) + 1
    'Store the pixel data into an array for faster accessing
    Dim PicArray() As Byte
    ReDim PicArray(0 To UBound(ImageData, 1), 0 To UBound(ImageData, 2)) As Byte
    Dim a As Long, b As Long
    Dim QuickVal As Long
    For x = PreviewX To PreviewX + PreviewWidth
        QuickVal = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
    For z = 0 To 2
        PicArray(QuickVal + z, y) = ImageData(QuickVal + z, y)
    Next z
    Next y
    Next x
    'Begin the main mosaic loop
    For x = 0 To XLoop
    For y = 0 To YLoop
        'This sub loop is to gather all of the data for the current mosaic tile
        InitXLoop = PreviewX + x * BlockSizeX
        InitYLoop = PreviewY + y * BlockSizeY
        DstXLoop = PreviewX + (x + 1) * BlockSizeX - 1
        DstYLoop = PreviewY + (y + 1) * BlockSizeY - 1
        For a = InitXLoop To DstXLoop
            QuickVal = a * 3
        For b = InitYLoop To DstYLoop
            'If this particular pixel is off of the image, don't bother counting it
            If a > (PreviewX + PreviewWidth) Or b > (PreviewY + PreviewHeight) Then GoTo 10011
            'total up all of the red, green, and blue values for the pixels within this
            'mosiac tile
            tR = tR + PicArray(QuickVal + 2, b)
            tG = tG + PicArray(QuickVal + 1, b)
            tB = tB + PicArray(QuickVal, b)
            'Count this as a valid pixel
            NumOfPixels = NumOfPixels + 1
10011   Next b
        Next a
        'If this tile is completely off of the image, don't worry about it and go to the next one
        If NumOfPixels = 0 Then GoTo 30011
        'Take the average red, green, and blue values of all the pixles within this tile
        tR = tR \ NumOfPixels
        tG = tG \ NumOfPixels
        tB = tB \ NumOfPixels
        'Now run a loop through the same pixels you just analyzed, only this time you're gonna
        'draw the averaged color over the top of them
        For a = InitXLoop To DstXLoop
            QuickVal = a * 3
        For b = InitYLoop To DstYLoop
            'Same thing as above - if it's off the image, ignore it
            If a > (PreviewX + PreviewWidth) Or b > (PreviewY + PreviewHeight) Then GoTo 20031
            'Set the pixel
            ImageData(QuickVal + 2, b) = tR
            ImageData(QuickVal + 1, b) = tG
            ImageData(QuickVal, b) = tB
20031   Next b
        Next a
        'Clear all the variables and go to the next pixel
30011   tR = 0
        tG = 0
        tB = 0
        NumOfPixels = 0
    Next y
    Next x
    
    SetPreviewData PicEffect

End Sub

Private Sub hsHeight_Change()
    txtHeight.Text = hsHeight.Value
    PreviewMosaicFilter (CSng(hsWidth.Value) / hsWidth.Max) * PreviewWidth, (CSng(hsHeight.Value) / hsHeight.Max) * PreviewHeight
End Sub

Private Sub hsWidth_Change()
    txtWidth.Text = hsWidth.Value
    PreviewMosaicFilter (CSng(hsWidth.Value) / hsWidth.Max) * PreviewWidth, (CSng(hsHeight.Value) / hsHeight.Max) * PreviewHeight
End Sub

Private Sub hsHeight_Scroll()
    txtHeight.Text = hsHeight.Value
    PreviewMosaicFilter (CSng(hsWidth.Value) / hsWidth.Max) * PreviewWidth, (CSng(hsHeight.Value) / hsHeight.Max) * PreviewHeight
End Sub

Private Sub hsWidth_Scroll()
    txtWidth.Text = hsWidth.Value
    PreviewMosaicFilter (CSng(hsWidth.Value) / hsWidth.Max) * PreviewWidth, (CSng(hsHeight.Value) / hsHeight.Max) * PreviewHeight
End Sub

Private Sub txtHeight_Change()
    If EntryValid(txtHeight, hsHeight.Min, hsHeight.Max, False, False) Then hsHeight.Value = val(txtHeight)
End Sub

Private Sub txtHeight_GotFocus()
    AutoSelectText txtHeight
End Sub

Private Sub txtWidth_Change()
    If EntryValid(txtWidth, hsWidth.Min, hsWidth.Max, False, False) Then hsWidth.Value = val(txtWidth)
End Sub

Private Sub txtWidth_GotFocus()
    AutoSelectText txtWidth
End Sub
