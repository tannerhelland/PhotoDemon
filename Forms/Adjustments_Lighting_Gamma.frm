VERSION 5.00
Begin VB.Form FormGamma 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Gamma Correction"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12060
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
   ScaleWidth      =   804
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdCheckBox chkUnison 
      Height          =   330
      Left            =   6120
      TabIndex        =   4
      Top             =   5280
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      Caption         =   "keep all colors in sync"
   End
   Begin VB.PictureBox picChart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   8280
      ScaleHeight     =   159
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltGamma 
      Height          =   705
      Index           =   0
      Left            =   6000
      TabIndex        =   5
      Top             =   2640
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "red"
      Min             =   0.01
      Max             =   3
      SigDigits       =   2
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin PhotoDemon.pdSlider sltGamma 
      Height          =   705
      Index           =   1
      Left            =   6000
      TabIndex        =   6
      Top             =   3540
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "green"
      Min             =   0.01
      Max             =   3
      SigDigits       =   2
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin PhotoDemon.pdSlider sltGamma 
      Height          =   705
      Index           =   2
      Left            =   6000
      TabIndex        =   3
      Top             =   4440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "blue"
      Min             =   0.01
      Max             =   3
      SigDigits       =   2
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   1005
      Index           =   2
      Left            =   6000
      Top             =   1170
      Width           =   2040
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   1
      Caption         =   "new gamma curve:"
      FontSize        =   12
      ForeColor       =   4210752
      Layout          =   1
   End
End
Attribute VB_Name = "FormGamma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gamma Correction Handler
'Copyright 2000-2018 by Tanner Helland
'Created: 12/May/01
'Last updated: 19/July/17
'Last update: convert to XML params, minor optimizations
'
'Gamma correction isn't exactly rocket science, but it's an important part of any good editing tool.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private m_UserChange As Boolean

Private Sub chkUnison_Click()
    
    If chkUnison.Value Then
        Dim newGamma As Double
        newGamma = CDblCustom(sltGamma(0) + sltGamma(1) + sltGamma(2)) / 3
    
        m_UserChange = False
        sltGamma(0) = newGamma
        sltGamma(1) = newGamma
        sltGamma(2) = newGamma
        m_UserChange = True
    End If
    
    UpdatePreview
    
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Gamma", , GetLocalParamString(), UNDO_Layer
End Sub

'When randomizing, do not check the "unison" box
Private Sub cmdBar_RandomizeClick()
    chkUnison.Value = False
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltGamma(0).Value = 1#
    sltGamma(1).Value = 1#
    sltGamma(2).Value = 1#
End Sub

'Basic gamma correction.  It's a simple function - use an exponent to adjust R/G/B values.
' Inputs: new gamma level, which channels to adjust (r/g/b/all), and optional preview information
Public Sub GammaCorrect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
     
    If (Not toPreview) Then Message "Adjusting gamma values..."
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    Dim rGamma As Double, gGamma As Double, bGamma As Double
    With cParams
        rGamma = .GetDouble("redgamma", sltGamma(0).Value)
        gGamma = .GetDouble("greengamma", sltGamma(1).Value)
        bGamma = .GetDouble("bluegamma", sltGamma(2).Value)
    End With
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Make certain that the gamma adjustment values we were passed are not zero
    If (rGamma = 0#) Then rGamma = 0.01
    If (gGamma = 0#) Then gGamma = 0.01
    If (bGamma = 0#) Then bGamma = 0.01
    
    'Divisions are expensive, so invert gamma values in advance
    rGamma = 1# / rGamma
    gGamma = 1# / gGamma
    bGamma = 1# / bGamma
    
    'Gamma can be easily applied using look-up tables
    Dim rLookup() As Byte, gLookup() As Byte, bLookup() As Byte
    ReDim rLookup(0 To 255) As Byte: ReDim gLookup(0 To 255) As Byte: ReDim bLookup(0 To 255) As Byte
    
    Dim rTmp As Double, gTmp As Double, bTmp As Double
    For x = 0 To 255
        
        rTmp = x / 255#
        gTmp = rTmp
        bTmp = rTmp
        
        rTmp = rTmp ^ rGamma
        bTmp = bTmp ^ bGamma
        gTmp = gTmp ^ gGamma
        
        rTmp = rTmp * 255
        gTmp = gTmp * 255
        bTmp = bTmp * 255
        
        If (rTmp < 0#) Then rTmp = 0#
        If (rTmp > 255#) Then rTmp = 255#
        rLookup(x) = rTmp
        
        If (gTmp < 0#) Then gTmp = 0#
        If (gTmp > 255#) Then gTmp = 255#
        gLookup(x) = gTmp
        
        If (bTmp < 0#) Then bTmp = 0#
        If (bTmp > 255#) Then bTmp = 255#
        bLookup(x) = bTmp
        
    Next x
        
    'Loop through each pixel in the image, converting values as we go
    initX = initX * qvDepth
    finalX = finalX * qvDepth
    
    For y = initY To finalY
    For x = initX To finalX Step qvDepth
        
        b = imageData(x, y)
        g = imageData(x + 1, y)
        r = imageData(x + 2, y)
        
        imageData(x, y) = bLookup(b)
        imageData(x + 1, y) = gLookup(g)
        imageData(x + 2, y) = rLookup(r)
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
   
End Sub

Private Sub Form_Load()
    cmdBar.MarkPreviewStatus False
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    m_UserChange = True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the preview effect and the gamma chart
' TODO: rewrite this against pd2D
Private Sub UpdatePreview()

    If cmdBar.PreviewsAllowed Then
    
        Dim prevX As Double, prevY As Double
        Dim curX As Double, curY As Double
        Dim x As Long, y As Long
        
        Dim xWidth As Long, yHeight As Long
        xWidth = picChart.ScaleWidth
        yHeight = picChart.ScaleHeight
            
        'Clear out the old chart and draw a gray line across the diagonal for reference
        picChart.Picture = LoadPicture(vbNullString)
        picChart.ForeColor = RGB(127, 127, 127)
        GDIPlusDrawLineToDC picChart.hDC, 0, yHeight, xWidth, 0, RGB(127, 127, 127)
        
        Dim gamVal As Double, tmpVal As Double
        
        'Draw each of the current gamma curves for the user's reference
        For y = 0 To 2
            
            'If all channels are in sync, draw only blue; otherwise, color each channel individually
            gamVal = sltGamma(y)
            If (sltGamma(0) = sltGamma(1)) And (sltGamma(1) = sltGamma(2)) Then
                picChart.ForeColor = RGB(0, 0, 255)
            Else
            
                Select Case y
                    Case 0
                        picChart.ForeColor = RGB(255, 0, 0)
                    Case 1
                        picChart.ForeColor = RGB(0, 192, 0)
                    Case 2
                        picChart.ForeColor = RGB(0, 0, 255)
                End Select
                
            End If
            
            prevX = 0
            prevY = yHeight
            curX = 0
            curY = yHeight
        
            'Draw the next channel (with antialiasing!)
            For x = 0 To xWidth
                tmpVal = x / xWidth
                tmpVal = tmpVal ^ (1# / gamVal)
                tmpVal = yHeight - (tmpVal * yHeight)
                curY = tmpVal
                curX = x
                GDIPlusDrawLineToDC picChart.hDC, prevX, prevY, curX, curY, picChart.ForeColor
                prevX = curX
                prevY = curY
            Next x
            
        Next y
        
        picChart.Picture = picChart.Image
        picChart.Refresh
    
        'Once the chart is done, redraw the gamma preview as well
        GammaCorrect GetLocalParamString, True, pdFxPreview
        
    End If
    
End Sub

Private Sub sltGamma_Change(Index As Integer)

    If m_UserChange And cmdBar.PreviewsAllowed Then
        m_UserChange = False
        
        If chkUnison.Value Then
            Select Case Index
                Case 0
                    sltGamma(1).Value = sltGamma(0).Value
                    sltGamma(2).Value = sltGamma(0).Value
                Case 1
                    sltGamma(0).Value = sltGamma(1).Value
                    sltGamma(2).Value = sltGamma(1).Value
                Case 2
                    sltGamma(0).Value = sltGamma(2).Value
                    sltGamma(1).Value = sltGamma(2).Value
            End Select
        End If
        
        m_UserChange = True
        
        UpdatePreview
    End If

End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "redgamma", sltGamma(0).Value
        .AddParam "greengamma", sltGamma(1).Value
        .AddParam "bluegamma", sltGamma(2).Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
