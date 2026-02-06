VERSION 5.00
Begin VB.Form FormGamma 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Gamma"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12060
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   Begin PhotoDemon.pdCommandBar cmdBar 
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
   Begin PhotoDemon.pdPictureBox picChart 
      Height          =   2415
      Left            =   8280
      Top             =   120
      Width           =   3495
      _ExtentX        =   0
      _ExtentY        =   0
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
      TabIndex        =   2
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
'Copyright 2000-2026 by Tanner Helland
'Created: 12/May/01
'Last updated: 09/September/21
'Last update: subtle improvements to rounding when converting from float-to-int
'
'Gamma correction isn't exactly rocket science, but it's an important part of any photo editor.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_UserChange As Boolean

'Exposure curve image
Private m_Graph As pdDIB

Private Sub chkUnison_Click()
    
    If chkUnison.Value Then
        Dim newGamma As Double
        newGamma = CDblCustom(sltGamma(0).Value + sltGamma(1).Value + sltGamma(2).Value) / 3
    
        m_UserChange = False
        sltGamma(0).Value = newGamma
        sltGamma(1).Value = newGamma
        sltGamma(2).Value = newGamma
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
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim rGamma As Double, gGamma As Double, bGamma As Double
    With cParams
        rGamma = .GetDouble("redgamma", sltGamma(0).Value)
        gGamma = .GetDouble("greengamma", sltGamma(1).Value)
        bGamma = .GetDouble("bluegamma", sltGamma(2).Value)
    End With
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Make certain that the gamma adjustment values we were passed are not zero
    ' (we divide by these values)
    Const GAMMA_MIN As Double = 0.01
    If (rGamma < GAMMA_MIN) Then rGamma = GAMMA_MIN
    If (gGamma < GAMMA_MIN) Then gGamma = GAMMA_MIN
    If (bGamma < GAMMA_MIN) Then bGamma = GAMMA_MIN
    
    'Gamma can be easily applied using look-up tables
    Dim rLookup() As Byte, gLookup() As Byte, bLookup() As Byte
    ReDim rLookup(0 To 255) As Byte: ReDim gLookup(0 To 255) As Byte: ReDim bLookup(0 To 255) As Byte
    
    Dim rTmp As Double, gTmp As Double, bTmp As Double
    For x = 0 To 255
        
        rTmp = x / 255#
        
        rTmp = rTmp ^ (1# / rGamma)
        gTmp = rTmp ^ (1# / gGamma)
        bTmp = rTmp ^ (1# / bGamma)
        
        rTmp = rTmp * 255#
        gTmp = gTmp * 255#
        bTmp = bTmp * 255#
        
        If (rTmp < 0#) Then rTmp = 0#
        If (rTmp > 255#) Then rTmp = 255#
        rLookup(x) = Int(rTmp + 0.5)
        
        If (gTmp < 0#) Then gTmp = 0#
        If (gTmp > 255#) Then gTmp = 255#
        gLookup(x) = Int(gTmp + 0.5)
        
        If (bTmp < 0#) Then bTmp = 0#
        If (bTmp > 255#) Then bTmp = 255#
        bLookup(x) = Int(bTmp + 0.5)
        
    Next x
        
    'Loop through each pixel in the image, converting values as we go
    initX = initX * 4
    finalX = finalX * 4
    
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        imageData(x) = bLookup(imageData(x))
        imageData(x + 1) = gLookup(imageData(x + 1))
        imageData(x + 2) = rLookup(imageData(x + 2))
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
   
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    m_UserChange = True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the preview effect and the gamma chart
Private Sub UpdatePreview()

    If cmdBar.PreviewsAllowed And (Not g_Themer Is Nothing) Then
    
        Dim prevX As Double, prevY As Double
        Dim curX As Double, curY As Double
        Dim x As Long, y As Long
        
        If (m_Graph Is Nothing) Then Set m_Graph = New pdDIB
        
        Dim xWidth As Long, yHeight As Long
        xWidth = picChart.GetWidth
        yHeight = picChart.GetHeight
        If (m_Graph.GetDIBWidth <> xWidth) Or (m_Graph.GetDIBHeight <> yHeight) Then m_Graph.CreateBlank xWidth, yHeight, 32, vbWhite, 255
        
        'pd2D handles rendering duties
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundPDDIB m_Graph
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfacePixelOffset P2_PO_Normal
        
        'We first want to wipe the old chart, then draw a gray line across the diagonal for reference
        Dim cBrush As pd2DBrush
        Set cBrush = New pd2DBrush
        cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_Background)
        PD2D.FillRectangleI cSurface, cBrush, 0, 0, xWidth - 1, yHeight - 1
        
        Dim cPen As pd2DPen
        Set cPen = New pd2DPen
        cPen.SetPenWidth 1!
        cPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
        PD2D.DrawRectangleI cSurface, cPen, 0, 0, xWidth - 1, yHeight - 1
        
        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        cSurface.SetSurfacePixelOffset P2_PO_Half
        PD2D.DrawLineI cSurface, cPen, 0, yHeight, xWidth, 0
        
        cPen.SetPenColor g_Themer.GetGenericUIColor(UI_Accent)
        cPen.SetPenWidth 1.6!
        cPen.SetPenLineJoin P2_LJ_Round
        cPen.SetPenLineCap P2_LC_Round
        
        'Shrink the chart by two pixels (to account for borders); we will add 1 to
        ' all coordinates in the inner loop to ensure the chart is properly centered
        yHeight = yHeight - 2
        xWidth = xWidth - 2
        
        Dim gamVal As Double, tmpVal As Double
        
        'If all channels are in sync, their curves will overlap; don't waste time drawing
        ' each channel - only draw blue
        Dim idxStart As Long, idxEnd As Long
        idxEnd = 2
        If (sltGamma(0).Value = sltGamma(1).Value) And (sltGamma(1).Value = sltGamma(2).Value) Then
            idxStart = 2
        Else
            idxStart = 0
        End If
        
        Dim listOfPoints() As PointFloat, numOfPoints As Long
        ReDim listOfPoints(0 To xWidth) As PointFloat
        
        'Draw each of the current gamma curves for the user's reference
        For y = idxStart To idxEnd
            
            numOfPoints = 0
            
            'If all channels are in sync, draw only blue; otherwise, color each channel individually
            gamVal = sltGamma(y).Value
            
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
                If (x = 0) Then prevY = curY
                If (curY > yHeight - 1) Then curY = yHeight - 1
                
                listOfPoints(numOfPoints).x = curX + 1
                listOfPoints(numOfPoints).y = curY + 1
                numOfPoints = numOfPoints + 1
                
                prevX = curX
                prevY = curY
                
            Next x
            
            'Render in the current channel's color
            Select Case y
                Case 0
                    cPen.SetPenColor g_Themer.GetGenericUIColor(UI_ChannelRed)
                Case 1
                    cPen.SetPenColor g_Themer.GetGenericUIColor(UI_ChannelGreen)
                Case 2
                    cPen.SetPenColor g_Themer.GetGenericUIColor(UI_ChannelBlue)
            End Select
            
            'Draw the finished line
            PD2D.DrawLinesF_FromPtF cSurface, cPen, numOfPoints, VarPtr(listOfPoints(0)), False
            
        Next y
        
        'Flip the finished buffer to screen
        Set cSurface = Nothing
        picChart.RequestRedraw True
        
        'Once the chart is done, redraw the gamma preview as well
        GammaCorrect GetLocalParamString, True, pdFxPreview
        
    End If
    
End Sub

Private Sub picChart_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_Graph Is Nothing) Then GDI.BitBltWrapper targetDC, 0, 0, ctlWidth, ctlHeight, m_Graph.GetDIBDC, 0, 0, vbSrcCopy
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
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "redgamma", sltGamma(0).Value
        .AddParam "greengamma", sltGamma(1).Value
        .AddParam "bluegamma", sltGamma(2).Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
