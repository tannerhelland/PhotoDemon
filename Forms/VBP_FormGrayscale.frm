VERSION 5.00
Begin VB.Form FormGrayscale 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Grayscale Conversion"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5460
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
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.HScrollBar hsShades 
      Height          =   255
      Left            =   1080
      Max             =   254
      Min             =   3
      TabIndex        =   2
      Top             =   1200
      Value           =   3
      Width           =   4215
   End
   Begin VB.TextBox txtShades 
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
      Left            =   480
      TabIndex        =   1
      Text            =   "3"
      Top             =   1170
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboMethod 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3495
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
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   1125
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
      Left            =   4080
      TabIndex        =   4
      Top             =   1920
      Width           =   1125
   End
End
Attribute VB_Name = "FormGrayscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Grayscale Conversion Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 1/12/02
'Last updated: 18/August/09
'Last update: homebrew methods now use the simpler (R+G+B)\3 method
'
'NOTE: this code still needs to be optimized and cleaned up - look to the
' grayscale project on THDC for specifics.
'
'Updated version of the grayscale handler; utilizes five different methods
'(average, ISU, desaturate, X # of shades, X # of shades dithered).
'
'***************************************************************************

Option Explicit

'*********************************************************************
'The next three methods exist to activate/deactivate the textbox and scrollbar
Private Sub cboMethod_Click()
    If cboMethod.ListIndex = 3 Or cboMethod.ListIndex = 4 Then
        ShowControls True
    Else
        ShowControls False
    End If
End Sub

Private Sub cboMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If cboMethod.ListIndex = 3 Or cboMethod.ListIndex = 4 Then
        ShowControls True
    Else
        ShowControls False
    End If
End Sub

Private Sub ShowControls(ByVal toShow As Boolean)
    txtShades.Visible = toShow
    hsShades.Visible = toShow
End Sub
'*********************************************************************

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Error checking
    If EntryValid(txtShades, hsShades.Min, hsShades.Max) Then
        Me.Visible = False
        Select Case cboMethod.ListIndex
            Case 0
                Process GrayscaleAverage
            Case 1
                Process GrayScale
            Case 2
                Process Desaturate
            Case 3
                Process GrayscaleCustom, hsShades.Value
            Case 4
                Process GrayscaleDitherCustom, hsShades.Value
        End Select
        Unload Me
    Else
        AutoSelectText txtShades
    End If

End Sub

'Initialize the combo box
Private Sub Form_Load()
    
    cboMethod.AddItem "Average", 0
    cboMethod.AddItem "ITU Standard", 1
    cboMethod.AddItem "Desaturate", 2
    cboMethod.AddItem "X # of shades", 3
    cboMethod.AddItem "X # of shades (dithered)", 4
    cboMethod.ListIndex = 1
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'Reduce to X # gray shades
Public Sub fGrayscaleCustom(ByVal NumToConvertTo As Long)
    GetImageData
    Message "Converting to " & NumToConvertTo & " shades of gray..."
    Dim MagicNum As Single
    Dim CurrentColor As Long
    MagicNum = (255 / (NumToConvertTo - 1))
    Dim LookUp(0 To 255) As Long
    For x = 0 To 255
        LookUp(x) = Int((CDbl(x) / MagicNum) + 0.5) * MagicNum
        If LookUp(x) > 255 Then LookUp(x) = 255
    Next x
    SetProgBarMax PicHeightL
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim QuickVal As Long
    For y = 0 To PicHeightL
    For x = 0 To PicWidthL
        QuickVal = x * 3
        r1 = ImageData(QuickVal + 2, y)
        g1 = ImageData(QuickVal + 1, y)
        b1 = ImageData(QuickVal, y)
        CurrentColor = (r1 + g1 + b1) \ 3
        ImageData(QuickVal + 2, y) = LookUp(CurrentColor)
        ImageData(QuickVal + 1, y) = LookUp(CurrentColor)
        ImageData(QuickVal, y) = LookUp(CurrentColor)
    Next x
        If y Mod 20 = 0 Then SetProgBarVal y
    Next y
    SetImageData
    Message "Finished."
End Sub

'Reduce to X # gray shades (dithered)
Public Sub fGrayscaleCustomDither(ByVal NumToConvertTo As Long)
    GetImageData
    Message "Converting to " & NumToConvertTo & " shades of gray..."
    Dim MagicNum As Single
    Dim CurrentColor As Long
    MagicNum = (255 / (NumToConvertTo - 1))
    SetProgBarMax PicHeightL
    Dim EV As Long
    Dim CC As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim QuickVal As Long
    For y = 0 To PicHeightL
    For x = 0 To PicWidthL
        QuickVal = x * 3
        r1 = ImageData(QuickVal + 2, y)
        g1 = ImageData(QuickVal + 1, y)
        b1 = ImageData(QuickVal, y)
        CurrentColor = (r1 + g1 + b1) \ 3
        CurrentColor = CurrentColor + EV
        CC = Int((CDbl(CurrentColor) / MagicNum) + 0.5) * MagicNum
        EV = CurrentColor - CC
        If CC > 255 Then CC = 255
        If CC < 0 Then CC = 0
        ImageData(QuickVal + 2, y) = CC
        ImageData(QuickVal + 1, y) = CC
        ImageData(QuickVal, y) = CC
    Next x
        EV = 0
        If y Mod 20 = 0 Then SetProgBarVal y
    Next y
    SetImageData
    Message "Finished."
End Sub

'Reduce to gray via (r+g+b)/3
Public Sub MenuGrayscaleAverage()
    Dim CurrentColor As Byte
    Dim r As Integer, g As Integer, b As Integer
    Message "Converting to grayscale..."
    SetProgBarMax PicWidthL
    Dim LookUp(0 To 765) As Byte
    For x = 0 To 765
        LookUp(x) = x \ 3
    Next x
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        CurrentColor = LookUp(r + g + b)
        ImageData(QuickVal, y) = CurrentColor
        ImageData(QuickVal + 1, y) = CurrentColor
        ImageData(QuickVal + 2, y) = CurrentColor
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
    Message "Finished."
End Sub

'Reduce to gray in a more human-eye friendly manner
Public Sub MenuGrayscale()
    Dim CurrentColor As Integer
    Dim r As Long, g As Long, b As Long
    Message "Generating ITU standard grayscale image..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        CurrentColor = (222 * r + 707 * g + 71 * b) \ 1000
        ByteMe CurrentColor
        ImageData(QuickVal, y) = CurrentColor
        ImageData(QuickVal + 1, y) = CurrentColor
        ImageData(QuickVal + 2, y) = CurrentColor
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
    Message "Finished."
End Sub

'Reduce to gray via HSL -> convert S to 0
Public Sub MenuDesaturate()
    Dim r As Long, g As Long, b As Long
    Dim HH As Single, SS As Single, LL As Single
    Message "Desaturating image..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        'Get the temporary values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        'Get the hue and saturation
        tRGBToHSL r, g, b, HH, SS, LL
        'Convert back to RGB using our artificial saturation value
        tHSLToRGB HH, 0, LL, r, g, b
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
    Message "Finished."
End Sub

Private Sub hsShades_Change()
    txtShades.Text = hsShades.Value
End Sub

Private Sub hsShades_Scroll()
    txtShades.Text = hsShades.Value
End Sub

Private Sub txtShades_Change()
    If EntryValid(txtShades, hsShades.Min, hsShades.Max, False, False) Then
        hsShades.Value = val(txtShades)
    End If
End Sub

Private Sub txtShades_GotFocus()
    AutoSelectText txtShades
End Sub
