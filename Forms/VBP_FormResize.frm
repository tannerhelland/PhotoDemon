VERSION 5.00
Begin VB.Form FormResize 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Resize Image"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4005
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
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VSHeight 
      Height          =   420
      Left            =   2430
      Max             =   32766
      Min             =   1
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1140
      Value           =   15000
      Width           =   270
   End
   Begin VB.VScrollBar VSWidth 
      Height          =   420
      Left            =   2430
      Max             =   32766
      Min             =   1
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   540
      Value           =   15000
      Width           =   270
   End
   Begin VB.TextBox TxtHeight 
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
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "N/A"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox TxtWidth 
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
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "N/A"
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cboResample 
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
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3600
      Width           =   1125
   End
   Begin VB.CheckBox ChkRatio 
      Appearance      =   0  'Flat
      Caption         =   "Preserve size ratio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton CmdResize 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3600
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
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
      Left            =   2850
      TabIndex        =   10
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
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
      Left            =   2850
      TabIndex        =   9
      Top             =   645
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
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
      Left            =   840
      TabIndex        =   8
      Top             =   1230
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
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
      Left            =   840
      TabIndex        =   7
      Top             =   645
      Width           =   555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Resample method:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   2895
   End
End
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Size Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 5/July/11
'Last update: added support for additional FreeImage resampling filters
'
'Handles all image-size related functions.  Currently supports standard resizing and halftone resampling
' (via the API; not 100% accurate but faster than doing it in VB code) and bilinear resampling via pure VB)
'
'***************************************************************************

Option Explicit

'Resampling declarations
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Buffer As Long, Buffer_hBitmap As Long

'Used to prevent the scroll bars from getting stuck in update loops
Dim updateWidthBar As Boolean, updateHeightBar As Boolean

'Used for maintaining ratios when the check box is clicked
Dim WRatio As Double, HRatio As Double

'If the ratio button is checked, then update the height box to match
Private Sub ChkRatio_Click()
    If ChkRatio.Value = vbChecked Then UpdateHeightBox
End Sub

'Resize an image using bicubic, bilinear, or nearest neighbor resampling
Public Sub ResizeImage(ByVal iWidth As Long, ByVal iHeight As Long, ByVal iMethod As Byte)

    'Nearest neighbor...
    If iMethod = RESIZE_NORMAL Then
        Message "Resizing image..."
        SetProgBarMax PicHeightL
        FormMain.ActiveForm.BackBuffer2.Width = iWidth + 2
        FormMain.ActiveForm.BackBuffer2.Height = iHeight + 2
        FormMain.ActiveForm.BackBuffer2.PaintPicture FormMain.ActiveForm.BackBuffer.Picture, 0, 0, iWidth, iHeight, 0, 0, PicWidthL, PicHeightL, vbSrcCopy
        FormMain.ActiveForm.BackBuffer2.Picture = FormMain.ActiveForm.BackBuffer2.Image
        FormMain.ActiveForm.BackBuffer.Cls
        FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
        PicWidthL = iWidth
        PicHeightL = iHeight
        pdImages(CurrentImage).Width = PicWidthL
        pdImages(CurrentImage).Height = PicHeightL
        DisplaySize PicWidthL, PicHeightL
        FormMain.ActiveForm.BackBuffer.Width = iWidth + 2
        FormMain.ActiveForm.BackBuffer.Height = iHeight + 2
        FormMain.ActiveForm.BackBuffer.PaintPicture FormMain.ActiveForm.BackBuffer2.Picture, 0, 0, iWidth, iHeight, 0, 0, iWidth, iHeight, vbSrcCopy
        FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
        FormMain.ActiveForm.BackBuffer2.Cls
        FormMain.ActiveForm.BackBuffer2.Picture = LoadPicture("")
        PrepareViewport FormMain.ActiveForm, "Image resized"
    
    'Halftone resampling... I'm not sure what to actually call it, but since it seems to be based off the
    ' StretchBlt mode Microsoft calls "halftone," I'm running with it
    ElseIf iMethod = RESIZE_HALFTONE Then
        Message "Resampling image..."
        Dim tmpImagePath As String
        tmpImagePath = TempPath & "PDResample.tmp"
        SavePicture FormMain.ActiveForm.BackBuffer.Image, tmpImagePath
        Buffer = CreateCompatibleDC(0)
        Buffer_hBitmap = LoadImage(ByVal 0&, tmpImagePath, 0, iWidth, iHeight, &H10)
        SelectObject Buffer, Buffer_hBitmap
        FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
        FormMain.ActiveForm.BackBuffer.Cls
        FormMain.ActiveForm.BackBuffer.Width = iWidth + 2
        FormMain.ActiveForm.BackBuffer.Height = iHeight + 2
        StretchBlt FormMain.ActiveForm.BackBuffer.hDC, 0, 0, iWidth, iHeight, Buffer, 0, 0, iWidth, iHeight, vbSrcCopy
        DeleteDC Buffer
        DeleteObject Buffer_hBitmap
        FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
        If FileExist(tmpImagePath) Then Kill tmpImagePath
        PicWidthL = iWidth
        PicHeightL = iHeight
        pdImages(CurrentImage).Width = PicWidthL
        pdImages(CurrentImage).Height = PicHeightL
        DisplaySize PicWidthL, PicHeightL
        PrepareViewport FormMain.ActiveForm, "Image resized"
        
    'True bilinear sampling
    ElseIf iMethod = RESIZE_BILINEAR Then
        Message "Resampling image..."
        SetProgBarMax iHeight
        FormMain.ActiveForm.BackBuffer2.Width = iWidth + 2
        FormMain.ActiveForm.BackBuffer2.Height = iHeight + 2
        
        GetImageData True
        
        'Current width and height
        Dim cWidth As Long, cHeight As Long
        cWidth = PicWidthL
        cHeight = PicHeightL
        
        GetImageData2 True
        
        'The scaled ratio between the old x and y values and the new ones
        Dim xScale As Single, yScale As Single
        'The destination x and y
        Dim dstX As Single, dstY As Single
        'The interpolated X and Y values
        Dim IntrplX As Integer, IntrplY As Integer
        'Calculation variables
        Dim CalcX As Single, CalcY As Single
        'The red and green values the we use to interpolate the new pixel
        Dim r As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
        Dim g As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
        Dim b As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
        'The interpolated red, green, and blue
        Dim ir1 As Long, ig1 As Long, ib1 As Long
        Dim ir2 As Long, ig2 As Long, ib2 As Long
        'Get the ratio between the old image and the new one
        xScale = cWidth / iWidth
        yScale = cHeight / iHeight

        'Draw each pixel in the new image
        For y = 0 To iHeight - 1
            'Generate the y calculation variables
            dstY = y * yScale
            IntrplY = Int(dstY)
            CalcY = dstY - IntrplY
            For x = 0 To iWidth - 1
                'Generate the x calculation variables
                dstX = x * xScale
                IntrplX = Int(dstX)
                CalcX = dstX - IntrplX
                
                'Get the 4 pixels around the interpolated one
                r1 = ImageData(IntrplX * 3 + 2, IntrplY)
                g1 = ImageData(IntrplX * 3 + 1, IntrplY)
                b1 = ImageData(IntrplX * 3, IntrplY)
                
                r2 = ImageData((IntrplX + 1) * 3 + 2, IntrplY)
                g2 = ImageData((IntrplX + 1) * 3 + 1, IntrplY)
                b2 = ImageData((IntrplX + 1) * 3, IntrplY)
                
                r3 = ImageData(IntrplX * 3 + 2, IntrplY + 1)
                g3 = ImageData(IntrplX * 3 + 1, IntrplY + 1)
                b3 = ImageData(IntrplX * 3, IntrplY + 1)
    
                r4 = ImageData((IntrplX + 1) * 3 + 2, IntrplY + 1)
                g4 = ImageData((IntrplX + 1) * 3 + 1, IntrplY + 1)
                b4 = ImageData((IntrplX + 1) * 3, IntrplY + 1)
    
                'Interpolate the R,G,B values in the X direction
                ir1 = r1 * (1 - CalcY) + r3 * CalcY
                ig1 = g1 * (1 - CalcY) + g3 * CalcY
                ib1 = b1 * (1 - CalcY) + b3 * CalcY
                ir2 = r2 * (1 - CalcY) + r4 * CalcY
                ig2 = g2 * (1 - CalcY) + g4 * CalcY
                ib2 = b2 * (1 - CalcY) + b4 * CalcY
                'Intepolate the R,G,B values in the Y direction
                r = ir1 * (1 - CalcX) + ir2 * CalcX
                g = ig1 * (1 - CalcX) + ig2 * CalcX
                b = ib1 * (1 - CalcX) + ib2 * CalcX
                
                'Make sure that the values are in the acceptable range
                ByteMeL r
                ByteMeL g
                ByteMeL b
                'Set this pixel onto the new picture box
                ImageData2(x * 3 + 2, y) = r
                ImageData2(x * 3 + 1, y) = g
                ImageData2(x * 3, y) = b
            Next x
            If y Mod 5 = 0 Then SetProgBarVal y
        Next y
        
        SetImageData2 True
        FormMain.ActiveForm.BackBuffer.Width = iWidth + 2
        FormMain.ActiveForm.BackBuffer.Height = iHeight + 2
        FormMain.ActiveForm.BackBuffer.Cls
        FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
        FormMain.ActiveForm.BackBuffer.PaintPicture FormMain.ActiveForm.BackBuffer2.Picture, 0, 0, iWidth, iHeight, 0, 0, iWidth, iHeight, vbSrcCopy
        FormMain.ActiveForm.BackBuffer2.Cls
        FormMain.ActiveForm.BackBuffer2.Picture = LoadPicture("")
        PicWidthL = iWidth
        PicHeightL = iHeight
        pdImages(CurrentImage).Width = PicWidthL
        pdImages(CurrentImage).Height = PicHeightL
        DisplaySize PicWidthL, PicHeightL
        FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
        PrepareViewport FormMain.ActiveForm, "Image resized"
        
        SetProgBarVal 0
        
        'Attempt to save some memory (even though ImageData2 should already be deleted)
        Erase ImageData
        Erase ImageData2
    
    ElseIf iMethod = RESIZE_BSPLINE Then
        FreeImageResize iWidth, iHeight, FILTER_BSPLINE
        
    ElseIf iMethod = RESIZE_BICUBIC_MITCHELL Then
        FreeImageResize iWidth, iHeight, FILTER_BICUBIC
        
    ElseIf iMethod = RESIZE_BICUBIC_CATMULL Then
        FreeImageResize iWidth, iHeight, FILTER_CATMULLROM
    
    ElseIf iMethod = RESIZE_LANCZOS Then
        FreeImageResize iWidth, iHeight, FILTER_LANCZOS3
        
    End If
    
End Sub

'Perform a resize operation
Private Sub CmdResize_Click()
    
    'Before resizing anything, check to make sure the textboxes have valid input
    If Not EntryValid(TxtWidth, 1, 32767, True, True) Then
        AutoSelectText TxtWidth
        Exit Sub
    End If
    If Not EntryValid(TxtHeight, 1, 32767, True, True) Then
        AutoSelectText TxtHeight
        Exit Sub
    End If
    
    Me.Visible = False
    
    'Resample based on the combo box entry...
    Select Case cboResample.ListIndex
        Case 0
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_NORMAL
        Case 1
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_HALFTONE
        Case 2
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_BILINEAR
        Case 3
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_BSPLINE
        Case 4
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_BICUBIC_MITCHELL
        Case 5
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_BICUBIC_CATMULL
        Case 6
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_LANCZOS
    End Select
    
    Unload Me
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

'Upon form load, determine the ratio between the width and height of the image
Private Sub Form_Load()
    
    'Add one to the displayed width and height, since we store them -1 for loops
    TxtWidth.Text = PicWidthL + 1
    TxtHeight.Text = PicHeightL + 1
    
    'Make the scroll bars match the text boxes
    updateWidthBar = False
    updateHeightBar = False
    VSWidth.Value = Abs(32767 - CInt(TxtWidth))
    VSHeight.Value = Abs(32767 - CInt(TxtHeight))
    updateWidthBar = True
    updateHeightBar = True
    
    'Establish ratios
    WRatio = (PicWidthL + 1) / (PicHeightL + 1)
    HRatio = (PicHeightL + 1) / (PicWidthL + 1)

    'Load up the combo box
    cboResample.AddItem "Nearest Neighbor", 0
    cboResample.AddItem "Halftone", 1
    cboResample.AddItem "Bilinear", 2
    cboResample.ListIndex = 2
    'If the FreeImage library is available, add additional resize options to the combo box
    If FreeImageEnabled = True Then
        cboResample.AddItem "B-Spline", 3
        cboResample.AddItem "Bicubic (Mitchell and Netravali)", 4
        cboResample.AddItem "Bicubic (Catmull-Rom)", 5
        cboResample.AddItem "Sinc (Lanczos3)", 6
        cboResample.ListIndex = 5
    End If
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'*************************************************************************************
'If "Preserve Size Ratio" is selected, this set of routines handles the preservation

Private Sub txtHeight_GotFocus()
    AutoSelectText TxtHeight
End Sub

Private Sub txtHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeToHeight
End Sub

Private Sub TxtHeight_LostFocus()
    ChangeToHeight
End Sub

Private Sub txtWidth_GotFocus()
    AutoSelectText TxtWidth
End Sub

Private Sub TxtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeToWidth
End Sub

Private Sub TxtWidth_LostFocus()
    ChangeToWidth
End Sub

Private Sub UpdateHeightBox()
    updateHeightBar = False
    TxtHeight = Int((CDbl(val(TxtWidth)) * HRatio) + 0.5)
    VSHeight.Value = Abs(32767 - val(TxtHeight))
    updateHeightBar = True
End Sub

Private Sub UpdateWidthBox()
    updateWidthBar = False
    TxtWidth = Int((CDbl(val(TxtHeight)) * WRatio) + 0.5)
    VSWidth.Value = Abs(32767 - val(TxtWidth))
    updateWidthBar = True
End Sub
'*************************************************************************************

Private Sub FreeImageResize(ByVal iWidth As Long, iHeight As Long, ByVal interpolationMethod As Long)
        'Load the FreeImage dll into memory
        Dim hLib As Long
        hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
        Message "Preparing image for resampling..."
    
        'Dump the image to a temporary file (a temporary work-around so we can capture
        'the handle required by FreeImage)
        Dim temporaryImg As String
        temporaryImg = TempPath & "PDResample.tmp"
        SavePicture FormMain.ActiveForm.BackBuffer.Picture, temporaryImg
    
        'These two variables will hold pointers to the bitmaps created by FreeImage calls
        Dim resizedDib As Long

        Message "Resampling image..."

        'Load the temp image into FreeImage and tell it to resize accordingly
        resizedDib = FreeImage_LoadEx(temporaryImg, , iWidth, iHeight, , interpolationMethod)
    
        'Paint the resized image to the current picture box
        Message "Rendering..."
        FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
        'FormMain.ActiveForm.BackBuffer.Cls
        FormMain.ActiveForm.BackBuffer.Width = iWidth + 2
        FormMain.ActiveForm.BackBuffer.Height = iHeight + 2
        Dim PaintReturn As Long
        PaintReturn = FreeImage_PaintDC(FormMain.ActiveForm.BackBuffer.hDC, resizedDib)
        FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
        FormMain.ActiveForm.BackBuffer.Refresh
        PicWidthL = iWidth
        PicHeightL = iHeight
        pdImages(CurrentImage).Width = PicWidthL
        pdImages(CurrentImage).Height = PicHeightL
        DisplaySize PicWidthL, PicHeightL
        
        PrepareViewport FormMain.ActiveForm, "Image resized (via FreeImage)"
    
        'Clear out the images generated by FreeImage
        If resizedDib <> 0 Then FreeImage_UnloadEx resizedDib
    
        'Release the library
        FreeLibrary hLib
    
        'Delete the temp file
        If FileExist(temporaryImg) Then Kill temporaryImg
    
        SetProgBarVal 0
        Message "Image resized successfully "
End Sub

'Because the scrollbars work backward (up means up and down means down) we have to reverse its input
' relative to the associated text box
Private Sub VSHeight_Change()
    If updateHeightBar = True Then
        TxtHeight = Abs(32767 - CStr(VSHeight.Value))
        ChangeToHeight
    End If
End Sub

Private Sub VSWidth_Change()
    If updateWidthBar = True Then
        TxtWidth = Abs(32767 - CStr(VSWidth.Value))
        ChangeToWidth
    End If
End Sub

Private Sub ChangeToWidth()
    If EntryValid(TxtWidth, 1, 32767, False, True) Then
        updateWidthBar = False
        VSWidth.Value = Abs(32767 - CInt(TxtWidth))
        updateWidthBar = True
        If ChkRatio.Value = vbChecked Then
            UpdateHeightBox
        End If
    End If
End Sub

Private Sub ChangeToHeight()
    If EntryValid(TxtHeight, 1, 32767, False, True) Then
        updateHeightBar = False
        VSHeight.Value = Abs(32767 - CInt(TxtHeight))
        updateHeightBar = True
        If ChkRatio.Value = vbChecked Then
            UpdateWidthBox
        End If
    End If
End Sub
