VERSION 5.00
Begin VB.Form FormRank 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Custom Rank Filter"
   ClientHeight    =   5085
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
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsRadius 
      Height          =   255
      Left            =   240
      Max             =   25
      Min             =   1
      MouseIcon       =   "VBP_FormMaximum.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3960
      Value           =   1
      Width           =   4575
   End
   Begin VB.TextBox txtRadius 
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
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Text            =   "1"
      Top             =   3480
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
   Begin VB.ComboBox cboRank 
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   1560
      MouseIcon       =   "VBP_FormMaximum.frx":0152
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      MouseIcon       =   "VBP_FormMaximum.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4560
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      MouseIcon       =   "VBP_FormMaximum.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4560
      Width           =   1125
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rank Method:"
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
      Left            =   360
      TabIndex        =   9
      Top             =   2820
      Width           =   1140
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
   Begin VB.Label lblRadius 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Radius:"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   3495
      Width           =   570
   End
End
Attribute VB_Name = "FormRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Rank (a.k.a. High/Low Pass, Dilate/Erode) Filter Interface
'©2000-2012 Tanner Helland
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

Private Sub cboRank_Click()
    UpdatePreview
End Sub

Private Sub cboRank_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdatePreview
End Sub

'OK Button
Private Sub CmdOK_Click()
    If EntryValid(txtRadius, hsRadius.Min, hsRadius.Max) Then
        Me.Visible = False
        Process CustomRank, val(hsRadius.Value), cboRank.ListIndex
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
Public Sub CustomRankFilter(ByVal Radius As Long, ByVal RankType As Byte)
    Dim tR As Long, tG As Long, tB As Long
    Dim MaxX As Long, MaxY As Long
    Dim MaxTotal As Long
    Dim c As Long, d As Long
    SetProgBarMax PicWidthL
    Dim FTransfer() As Byte
    Dim tTransfer() As Integer
    Dim TempColor As Long
    Dim tWidth As Long
    tWidth = ((PicWidthL + 1) * 3) - 1
    tWidth = tWidth + ((PicWidthL + 1) Mod 4)
    ReDim FTransfer(0 To tWidth, 0 To PicHeightL) As Byte
    ReDim tTransfer(0 To tWidth, 0 To PicHeightL) As Integer
    Dim QuickVal As Long
    Dim QuickVal2 As Long
    Message "Preparing rank data..."
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TempColor = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        tTransfer(x, y) = TempColor
    Next y
    Next x
    Message "Running custom rank filter..."
    For x = 0 To PicWidthL
        QuickVal2 = x * 3
    For y = 0 To PicHeightL
        If RankType = 0 Then
            MaxTotal = -1
        ElseIf RankType = 1 Then
            MaxTotal = 256
        Else
            MaxTotal = -1
        End If
        For c = x - Radius To x + Radius
        For d = y - Radius To y + Radius
            If c < 0 Then GoTo 303
            If c > (PicWidthL - 1) Then GoTo 303
            If d < 0 Then GoTo 303
            If d > (PicHeightL - 1) Then GoTo 303
            If RankType = 0 Then
                If tTransfer(c, d) > MaxTotal Then
                    MaxTotal = tTransfer(c, d)
                    MaxX = c
                    MaxY = d
                End If
            ElseIf RankType = 1 Then
                If tTransfer(c, d) < MaxTotal Then
                    MaxTotal = tTransfer(c, d)
                    MaxX = c
                    MaxY = d
                End If
            Else
                TempColor = Abs(tTransfer(x, y) - tTransfer(c, d))
                If TempColor > MaxTotal Then
                    MaxTotal = TempColor
                    MaxX = c
                    MaxY = d
                End If
            End If
303     Next d
        Next c
        QuickVal = MaxX * 3
        FTransfer(QuickVal2 + 2, y) = ImageData(QuickVal + 2, MaxY)
        FTransfer(QuickVal2 + 1, y) = ImageData(QuickVal + 1, MaxY)
        FTransfer(QuickVal2, y) = ImageData(QuickVal, MaxY)
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    For x = 0 To PicWidthL
    QuickVal = x * 3
    For y = 0 To PicHeightL
    For z = 0 To 2
        ImageData(QuickVal + z, y) = FTransfer(QuickVal + z, y)
    Next z
    Next y
    Next x
    SetProgBarVal cProgBar.Max
    Message "Custom rank filter finished.  Creating correct image array..."
    SetImageData
    Unload Me
End Sub

'Maximum rank/Dilate/High Pass
Public Sub rMaximize()
    Dim tR As Long, tG As Long, tB As Long
    Dim MaxX As Integer, MaxY As Integer
    Dim MaxTotal As Integer
    Dim c As Integer, d As Integer
    SetProgBarMax PicWidthL
    Dim FTransfer() As Byte
    Dim tTransfer() As Byte
    Dim TempColor As Long
    ReDim FTransfer(0 To (PicWidthL + 1) * 3, 0 To PicHeightL) As Byte
    ReDim tTransfer(0 To PicWidthL, 0 To PicHeightL) As Byte
    Message "Preparing rank data..."
    GetImageData
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TempColor = Int((tR + tG + tB) \ 3)
        tTransfer(x, y) = TempColor
    Next y
    Next x
    Message "Dilating image..."
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        MaxTotal = -1
        For c = x - 1 To x + 1
        For d = y - 1 To y + 1
            If (c < 0) Or (c > PicWidthL) Or (d < 0) Or (d > PicHeightL) Then GoTo 3030
            If tTransfer(c, d) > MaxTotal Then
                MaxTotal = tTransfer(c, d)
                MaxX = c
                MaxY = d
            End If
3030     Next d
        Next c
        FTransfer(QuickVal + 2, y) = ImageData(MaxX * 3 + 2, MaxY)
        FTransfer(QuickVal + 1, y) = ImageData(MaxX * 3 + 1, MaxY)
        FTransfer(QuickVal, y) = ImageData(MaxX * 3, MaxY)
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    For z = 0 To 2
        ImageData(QuickVal + z, y) = FTransfer(QuickVal + z, y)
    Next z
    Next y
    Next x
    SetProgBarVal cProgBar.Max
    Message "Image dilated successfully.  Generating final data..."
    SetImageData
End Sub

'Minimum Rank/Erode/Low Pass
Public Sub rMinimize()
    Dim tR As Long, tG As Long, tB As Long
    Dim MaxX As Integer, MaxY As Integer
    Dim MaxTotal As Integer
    Dim c As Integer, d As Integer
    SetProgBarMax PicWidthL
    Dim FTransfer() As Byte
    Dim tTransfer() As Byte
    Dim TempColor As Long
    ReDim FTransfer(0 To (PicWidthL + 1) * 3, 0 To PicHeightL) As Byte
    ReDim tTransfer(0 To PicWidthL, 0 To PicHeightL) As Byte
    Message "Preparing rank data..."
    GetImageData
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TempColor = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        'TempColor = Int((TR + TG + TB) \ 3)
        tTransfer(x, y) = TempColor
    Next y
    Next x
    Message "Eroding image..."
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        MaxTotal = 256
        For c = x - 1 To x + 1
        For d = y - 1 To y + 1
            If (c < 0) Or (c > PicWidthL) Or (d < 0) Or (d > PicHeightL) Then GoTo 3031
            If tTransfer(c, d) < MaxTotal Then
                MaxTotal = tTransfer(c, d)
                MaxX = c
                MaxY = d
            End If
3031     Next d
        Next c
        FTransfer(QuickVal + 2, y) = ImageData(MaxX * 3 + 2, MaxY)
        FTransfer(QuickVal + 1, y) = ImageData(MaxX * 3 + 1, MaxY)
        FTransfer(QuickVal, y) = ImageData(MaxX * 3, MaxY)
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    For z = 0 To 2
        ImageData(QuickVal + z, y) = FTransfer(QuickVal + z, y)
    Next z
    Next y
    Next x
    SetProgBarVal cProgBar.Max
    Message "Image eroded successfully.  Generating final data..."
    SetImageData
End Sub

'My own original combination of maximum and minimum: extreme!  :)
Public Sub rExtreme()
    Dim tR As Long, tG As Long, tB As Long
    Dim MaxX As Integer, MaxY As Integer
    Dim MaxTotal As Integer
    Dim c As Integer, d As Integer
    SetProgBarMax PicWidthL
    Dim FTransfer() As Byte
    Dim tTransfer() As Integer
    Dim TempColor As Long
    ReDim FTransfer(0 To (PicWidthL + 1) * 3, 0 To PicHeightL) As Byte
    ReDim tTransfer(0 To PicWidthL, 0 To PicHeightL) As Integer
    Message "Preparing rank data..."
    GetImageData
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TempColor = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        tTransfer(x, y) = TempColor
    Next y
    Next x
    Message "Applying extreme rank function..."
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        MaxTotal = -1
        For c = x - 1 To x + 1
        For d = y - 1 To y + 1
            If c < 0 Or c > PicWidthL Or d < 0 Or d > PicHeightL Then GoTo 3032
            TempColor = Abs(tTransfer(x, y) - tTransfer(c, d))
            If TempColor > MaxTotal Then
                MaxTotal = TempColor
                MaxX = c
                MaxY = d
            End If
3032     Next d
        Next c
        FTransfer(QuickVal + 2, y) = ImageData(MaxX * 3 + 2, MaxY)
        FTransfer(QuickVal + 1, y) = ImageData(MaxX * 3 + 1, MaxY)
        FTransfer(QuickVal, y) = ImageData(MaxX * 3, MaxY)
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    For z = 0 To 2
        ImageData(QuickVal + z, y) = FTransfer(QuickVal + z, y)
    Next z
    Next y
    Next x
    SetProgBarVal cProgBar.Max
    Message "Extreme rank function applied successfully.  Generating final data..."
    SetImageData
End Sub

Private Sub Form_Load()
    'Possible methods of calculating rank filters:
    cboRank.AddItem "Maximum (Dilate)", 0
    cboRank.AddItem "Minimum (Erode)", 1
    cboRank.AddItem "Extreme (Furthest value)", 2
    'Make "Maximum" the default value
    cboRank.ListIndex = 0
    'Create the image previews
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    UpdatePreview
End Sub

'Same as above, but exclusively for previewing
Private Sub PreviewRank(ByVal Radius As Long, ByVal RankType As Byte)
    Dim tR As Long, tG As Long, tB As Long
    Dim MaxX As Long, MaxY As Long
    Dim MaxTotal As Long
    Dim c As Long, d As Long
    GetPreviewData PicPreview
    Dim FTransfer() As Byte
    Dim tTransfer() As Integer
    Dim TempColor As Long
    PicWidthL = PicPreview.ScaleWidth
    PicHeightL = PicPreview.ScaleHeight
    Dim tWidth As Long
    tWidth = (PicWidthL * 3) - 1
    tWidth = tWidth + (PicWidthL Mod 4)
    ReDim FTransfer(0 To tWidth, 0 To PicHeightL) As Byte
    ReDim tTransfer(0 To tWidth, 0 To PicHeightL) As Integer
    Dim QuickVal As Long
    Dim QuickVal2 As Long
    Dim fX As Long, fY As Long
    fX = PreviewX + PreviewWidth
    fY = PreviewY + PreviewHeight
    For x = PreviewX To fX
        QuickVal = x * 3
    For y = PreviewY To fY
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TempColor = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        tTransfer(x, y) = TempColor
    Next y
    Next x
    For x = PreviewX To fX
        QuickVal2 = x * 3
    For y = PreviewY To fY
        If RankType = 0 Then
            MaxTotal = -1
        ElseIf RankType = 1 Then
            MaxTotal = 256
        Else
            MaxTotal = -1
        End If
        For c = x - Radius To x + Radius
        For d = y - Radius To y + Radius
            If c < PreviewX Then GoTo 303303
            If c > fX Then GoTo 303303
            If d < PreviewY Then GoTo 303303
            If d > fY Then GoTo 303303
            If RankType = 0 Then
                If tTransfer(c, d) > MaxTotal Then
                    MaxTotal = tTransfer(c, d)
                    MaxX = c
                    MaxY = d
                End If
            ElseIf RankType = 1 Then
                If tTransfer(c, d) < MaxTotal Then
                    MaxTotal = tTransfer(c, d)
                    MaxX = c
                    MaxY = d
                End If
            Else
                TempColor = Abs(tTransfer(x, y) - tTransfer(c, d))
                If TempColor > MaxTotal Then
                    MaxTotal = TempColor
                    MaxX = c
                    MaxY = d
                End If
            End If
303303     Next d
        Next c
        QuickVal = MaxX * 3
        FTransfer(QuickVal2 + 2, y) = ImageData(QuickVal + 2, MaxY)
        FTransfer(QuickVal2 + 1, y) = ImageData(QuickVal + 1, MaxY)
        FTransfer(QuickVal2, y) = ImageData(QuickVal, MaxY)
    Next y
    Next x
    For x = PreviewX To fX
    QuickVal = x * 3
    For y = PreviewY To fY
        'I have removed the z-loop in an attempt to speed things up
        ImageData(QuickVal + 2, y) = FTransfer(QuickVal + 2, y)
        ImageData(QuickVal + 1, y) = FTransfer(QuickVal + 1, y)
        ImageData(QuickVal, y) = FTransfer(QuickVal, y)
    Next y
    Next x
    SetPreviewData PicEffect
End Sub

Private Sub hsRadius_Change()
    txtRadius.Text = hsRadius.Value
    UpdatePreview
End Sub

Private Sub hsRadius_Scroll()
    txtRadius.Text = hsRadius.Value
    UpdatePreview
End Sub

Private Sub txtRadius_Change()
    If EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, False, False) Then hsRadius.Value = val(txtRadius)
End Sub

Private Sub txtRadius_GotFocus()
    AutoSelectText txtRadius
End Sub

Private Sub UpdatePreview()
    GetPreviewData PicPreview
    Dim maxSide As Long
    If PicWidthL > PicHeightL Then
        maxSide = hsRadius.Value * (PicWidthL / pdImages(CurrentImage).PicWidth)
    Else
        maxSide = hsRadius.Value * (PicHeightL / pdImages(CurrentImage).PicHeight)
    End If
    PreviewRank maxSide, cboRank.ListIndex
End Sub
