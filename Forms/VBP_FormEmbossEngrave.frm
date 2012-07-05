VERSION 5.00
Begin VB.Form FormEmbossEngrave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Emboss/Engrave"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OptEmboss 
      Appearance      =   0  'Flat
      Caption         =   "Emboss"
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
      Height          =   255
      Left            =   1200
      MouseIcon       =   "VBP_FormEmbossEngrave.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2880
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton OptEngrave 
      Appearance      =   0  'Flat
      Caption         =   "Engrave"
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
      Height          =   255
      Left            =   2880
      MouseIcon       =   "VBP_FormEmbossEngrave.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2880
      Width           =   975
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
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
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
      TabIndex        =   6
      Top             =   120
      Width           =   2175
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
      MouseIcon       =   "VBP_FormEmbossEngrave.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4440
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
      MouseIcon       =   "VBP_FormEmbossEngrave.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4440
      Width           =   1125
   End
   Begin VB.PictureBox PicColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3600
      MouseIcon       =   "VBP_FormEmbossEngrave.frx":0548
      MousePointer    =   99  'Custom
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox ChkToColor 
      Appearance      =   0  'Flat
      Caption         =   "To Color (click colored box to change)..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      MouseIcon       =   "VBP_FormEmbossEngrave.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3615
      Width           =   3255
   End
   Begin VB.Label Label2 
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
Attribute VB_Name = "FormEmbossEngrave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Emboss/Engrave Filter Interface
'©2000-2012 Tanner Helland
'Created: 3/6/03
'Last updated: 5/June/07
'Last update: previews! Yay for previews!
'
'Module for handling all emboss and engrave filters.  It's basically just an
'interfacing layer to the 4 main filters: Emboss/EmbossToColor and Engrave/EngraveToColor
'
'***************************************************************************

Option Explicit

Private Sub ChkToColor_Click()
    DrawPreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Used to remember the last color used for embossing
    EmbossEngraveColor = PicColor.BackColor
    Me.Visible = False
    
    'Dependent: filter to grey OR to a background color
    If OptEmboss.Value = True Then
        If ChkToColor.Value = vbChecked Then Process EmbossToColor, PicColor.BackColor Else Process EmbossToColor, RGB(127, 127, 127)
    Else
        If ChkToColor.Value = vbChecked Then Process EngraveToColor, PicColor.BackColor Else Process EngraveToColor, RGB(127, 127, 127)
    End If
    
    Unload Me
End Sub

'LOAD the form
Private Sub Form_Load()
    'Remember the last emboss/engrave color selection
    PicColor.BackColor = EmbossEngraveColor
    
    'Preview stuff
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    DrawPreview
    
End Sub

Private Sub OptEmboss_Click()
    DrawPreview
End Sub

Private Sub OptEngrave_Click()
    DrawPreview
End Sub

'Clicking on the picture box allows the user to select a new color
Private Sub PicColor_Click()
    Dim retColor As Long
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    retColor = PicColor.BackColor
    CD1.VBChooseColor retColor, True, True, False, Me.hWnd
    If retColor > 0 Then
        PicColor.BackColor = retColor
        ChkToColor.Value = vbChecked
    End If
    DrawPreview
End Sub

Public Sub FilterEmbossColor(ByVal cColor As Long)

    Dim r As Long, g As Long, B As Long
    Dim tR As Long, tB As Long, tG As Long
    
    Message "Engraving image..."
    
    SetProgBarMax PicWidthL
    
    tR = ExtractR(cColor)
    tG = ExtractG(cColor)
    tB = ExtractB(cColor)
    
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1) As Byte
    
    Dim QuickX As Long, QuickXRight As Long
    
    For x = 0 To PicWidthL - 1
        QuickX = x * 3
        QuickXRight = (x + 1) * 3
    For y = 0 To PicHeightL
        
        r = Abs(CLng(ImageData(QuickX + 2, y + 1)) - CLng(ImageData(QuickXRight + 2, y)) + tR)
        g = Abs(CLng(ImageData(QuickX + 1, y + 1)) - CLng(ImageData(QuickXRight + 1, y)) + tG)
        B = Abs(CLng(ImageData(QuickX, y + 1)) - CLng(ImageData(QuickXRight, y)) + tB)
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If B > 255 Then B = 255
        If B < 0 Then B = 0

        tData(QuickX + 2, y) = r
        tData(QuickX + 1, y) = g
        tData(QuickX, y) = B
        
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    TransferImageData
    
    SetImageData
    
End Sub

Public Sub FilterEngraveColor(ByVal cColor As Long)

    Dim r As Long, g As Long, B As Long
    Dim tR As Long, tB As Long, tG As Long
    
    Message "Engraving image..."
    
    SetProgBarMax PicWidthL
    
    tR = ExtractR(cColor)
    tG = ExtractG(cColor)
    tB = ExtractB(cColor)
    
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1) As Byte
    
    Dim QuickX As Long, QuickXRight As Long
    
    For x = 0 To PicWidthL - 1
        QuickX = x * 3
        QuickXRight = (x + 1) * 3
    For y = 0 To PicHeightL
        
        r = Abs(CLng(ImageData(QuickXRight + 2, y + 1)) - CLng(ImageData(QuickX + 2, y)) + tR)
        g = Abs(CLng(ImageData(QuickXRight + 1, y + 1)) - CLng(ImageData(QuickX + 1, y)) + tG)
        B = Abs(CLng(ImageData(QuickXRight, y + 1)) - CLng(ImageData(QuickX, y)) + tB)
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If B > 255 Then B = 255
        If B < 0 Then B = 0

        tData(QuickX + 2, y) = r
        tData(QuickX + 1, y) = g
        tData(QuickX, y) = B
        
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    TransferImageData
    
    SetImageData
    
End Sub

Private Sub DrawPreview()

    Dim cColor As Long
    If ChkToColor.Value = vbChecked Then
        cColor = PicColor.BackColor
    Else
        cColor = RGB(127, 127, 127)
    End If
    Dim toEmboss As Boolean
    toEmboss = OptEmboss.Value

    GetPreviewData PicPreview

    ReDim tData(0 To (PreviewWidth + PreviewX * 2) * 3, 0 To PreviewHeight + PreviewY * 2)
    
    If toEmboss = False Then
        Dim r As Integer, g As Integer, B As Integer
        Dim tR As Integer, tB As Integer, tG As Integer
        tR = ExtractR(cColor)
        tG = ExtractG(cColor)
        tB = ExtractB(cColor)
        For x = PreviewX To PreviewX + PreviewWidth - 1
        For y = PreviewY To PreviewY + PreviewHeight - 1
            r = Abs(CInt(ImageData((x + 1) * 3 + 2, y + 1)) - CInt(ImageData(x * 3 + 2, y)) + tR)
            g = Abs(CInt(ImageData((x + 1) * 3 + 1, y + 1)) - CInt(ImageData(x * 3 + 1, y)) + tG)
            B = Abs(CInt(ImageData((x + 1) * 3, y + 1)) - CInt(ImageData(x * 3, y)) + tB)
            ByteMe r
            ByteMe g
            ByteMe B
            tData(x * 3 + 2, y) = r
            tData(x * 3 + 1, y) = g
            tData(x * 3, y) = B
        Next y
        Next x
    Else
        Dim dR As Integer, dG As Integer, dB As Integer
        dR = ExtractR(cColor)
        dG = ExtractG(cColor)
        dB = ExtractB(cColor)
        For x = PreviewX To PreviewX + PreviewWidth - 1
        For y = PreviewY To PreviewY + PreviewHeight - 1
            tR = Abs(CInt(ImageData(x * 3 + 2, y)) - CInt(ImageData((x + 1) * 3 + 2, y)) + dR)
            tG = Abs(CInt(ImageData(x * 3 + 1, y)) - CInt(ImageData((x + 1) * 3 + 1, y)) + dG)
            tB = Abs(CInt(ImageData(x * 3, y)) - CInt(ImageData((x + 1) * 3, y)) + dB)
            ByteMe tR
            ByteMe tG
            ByteMe tB
            tData(x * 3 + 2, y) = tR
            tData(x * 3 + 1, y) = tG
            tData(x * 3, y) = tB
        Next y
        Next x

    End If
    
    Dim QuickVal As Long
    For x = PreviewX To PreviewX + PreviewWidth - 1
        QuickVal = x * 3
    For y = PreviewY To PreviewY + PreviewHeight - 1
        For z = 0 To 2
            ImageData(QuickVal + z, y) = tData(QuickVal + z, y)
        Next z
    Next y
    Next x
    SetPreviewData PicEffect
    
End Sub
