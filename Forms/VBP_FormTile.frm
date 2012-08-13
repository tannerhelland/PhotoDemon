VERSION 5.00
Begin VB.Form FormTile 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Generate Twins"
   ClientHeight    =   3855
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   2640
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox PicPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   240
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.OptionButton OptVertical 
      Appearance      =   0  'Flat
      Caption         =   "Vertical"
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
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.OptionButton OptHorizontal 
      Appearance      =   0  'Flat
      Caption         =   "Horizontal"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   2760
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3360
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Label lblPreview 
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
      TabIndex        =   6
      Top             =   2310
      Width           =   4575
   End
End
Attribute VB_Name = "FormTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Twin" Filter Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 13/June/12
'Last update: fixed RGB misalignment when performing vertical flipping
'Need to update: Needs optimization. Only run the loop through half the image,
'                and double-store values (they are mirrored, after all - duh!!)
'
'Unoptimized "twin" generator.  Simple 50% alpha blending combined with a flip.
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
    If OptVertical.Value = True Then
        Process Tile, 0
    Else
        Process Tile, 1
    End If
    Unload Me
End Sub

'This routine mirrors and alphablends an image, making it "tilable" or symmetrical
Public Sub GenerateTwins(ByVal TType As Byte)
    
    GetImageData
    
    'Temporary colors
    Dim Color1 As Long, Color2 As Long
    
    'Temporary array to store the image information (preventing bad overlaps)
    Dim Ta() As Byte
    Dim tWidth As Long
    tWidth = (PicWidthL * 3) - 1
    tWidth = tWidth + (PicWidthL Mod 4)
    
    ReDim Ta(0 To tWidth, 0 To PicHeightL) As Byte
    
    Message "Creating twin image..."
    
    'First, copy our array into TA()
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        Ta(QuickVal + 2, y) = ImageData(QuickVal + 2, y)
        Ta(QuickVal + 1, y) = ImageData(QuickVal + 1, y)
        Ta(QuickVal, y) = ImageData(QuickVal, y)
    Next y
    Next x
    
    SetProgBarMax PicWidthL
    
    Dim PicBitsX As Long
    Dim NewColor As Long
    
    PicBitsX = PicWidthL * 3
    
    'This loop will actually generate the twins
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    For z = 0 To 2
        'Get the value of the "first" pixel
        Color1 = Ta(QuickVal + z, y)
        'Get the value of the "second" pixel, depending on the method
        If TType = 0 Then
            Color2 = Ta(QuickVal + z, PicHeightL - y)
        Else
            Color2 = Ta(PicBitsX - QuickVal + z, y)
        End If
        'Simple alpha-blend, kids
        NewColor = (Color1 + Color2) \ 2
        'Remember this value and continue
        ImageData(QuickVal + z, y) = NewColor
    Next z
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Create the image previews
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    PreviewTwins 1
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

Private Sub OptHorizontal_Click()
    PreviewTwins 1
End Sub

Private Sub OptVertical_Click()
    PreviewTwins 0
End Sub

'Same routine as above, but only for previewing
Private Sub PreviewTwins(ByVal TType As Byte)
    GetPreviewData PicPreview
    Dim Color1 As Long, Color2 As Long
    Dim Ta() As Byte
    Dim tWidth As Long
    PicWidthL = PicPreview.ScaleWidth
    PicHeightL = PicPreview.ScaleHeight
    tWidth = (PicWidthL * 3) + 2
    tWidth = tWidth + (PicWidthL Mod 4)
    ReDim Ta(0 To tWidth, 0 To PicHeightL + 1) As Byte
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        Ta(QuickVal + 2, y) = ImageData(QuickVal + 2, y)
        Ta(QuickVal + 1, y) = ImageData(QuickVal + 1, y)
        Ta(QuickVal, y) = ImageData(QuickVal, y)
    Next y
    Next x
    Dim PicBitsX As Long
    Dim NewColor As Long
    PicBitsX = PicWidthL * 3
    For x = PreviewX To PreviewX + PreviewWidth
        QuickVal = x * 3
    For y = PreviewY To PreviewY + PreviewHeight
    For z = 0 To 2
        Color1 = Ta(QuickVal + z, y)
        If TType = 0 Then
            Color2 = Ta(QuickVal + z, PicHeightL - y)
        Else
            Color2 = Ta(PicBitsX - QuickVal + z, y)
        End If
        NewColor = (Color1 + Color2) \ 2
        ImageData(QuickVal + z, y) = NewColor
    Next z
    Next y
    Next x
    SetPreviewData PicEffect
End Sub
