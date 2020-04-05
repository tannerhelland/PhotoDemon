VERSION 5.00
Begin VB.Form FormRechannel 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Rechannel"
   ClientHeight    =   6570
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   11895
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
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsColorSpace 
      Height          =   1095
      Left            =   5880
      TabIndex        =   2
      Top             =   1440
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1931
      Caption         =   "color space"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5820
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1323
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
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   255
      Index           =   1
      Left            =   5880
      Top             =   2760
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      Caption         =   "channel"
      FontSize        =   12
   End
   Begin PhotoDemon.pdButtonStrip btsChannel 
      Height          =   615
      Index           =   0
      Left            =   6000
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonStrip btsChannel 
      Height          =   615
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonStrip btsChannel 
      Height          =   615
      Index           =   2
      Left            =   6000
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1085
   End
End
Attribute VB_Name = "FormRechannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Rechannel Interface
'Copyright 2000-2020 by Tanner Helland
'Created: original rechannel algorithm - sometimes 2001, this form 28/September/12
'Last updated: 04/December/15
'Last update: overhaul interface, switch to new XML parameter class
'
'Rechannel (or "channel isolation") tool.  This allows the user to isolate a single color channel from
' the RGB and CMY/CMYK color spaces.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsChannel_Click(Index As Integer, ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsColorSpace_Click(ByVal buttonIndex As Long)
    
    Dim i As Long
    For i = btsChannel.lBound To btsChannel.UBound
        btsChannel(i).Visible = (i = buttonIndex)
    Next i
    
    UpdatePreview
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Rechannel", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Align all channel button strips
    Dim i As Long
    For i = 0 To btsChannel.Count - 1
        If i = 0 Then
            btsChannel(i).Visible = True
        Else
            btsChannel(i).Top = btsChannel(0).Top
            btsChannel(i).Visible = False
        End If
    Next i
    
    'Populate all button strip captions
    btsColorSpace.AddItem "RGB", 0
    btsColorSpace.AddItem "CMY", 1
    btsColorSpace.AddItem "CMYK", 2
    btsColorSpace.ListIndex = 0
    
    btsChannel(0).AddItem "red", 0
    btsChannel(0).AddItem "green", 1
    btsChannel(0).AddItem "blue", 2
    
    btsChannel(1).AddItem "cyan", 0
    btsChannel(1).AddItem "magenta", 1
    btsChannel(1).AddItem "yellow", 2
    
    btsChannel(2).AddItem "cyan", 0
    btsChannel(2).AddItem "magenta", 1
    btsChannel(2).AddItem "yellow", 2
    btsChannel(2).AddItem "key (black)", 3
    
    'Apply translations and visual themes, and supply an initial effect preview
    ApplyThemeAndTranslations Me
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Rechannel an image
' INPUTS:
' - color space (currently supports RGB, CMY, CMYK)
' - channel (varies by color space)
Public Sub RechannelImage(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim dstColorSpace As Long, dstChannel As Long
    dstColorSpace = cParams.GetLong("colorspace", 0)
    dstChannel = cParams.GetLong("channel", 0)
    
    'Based on the color space and channel the user has selected, display a user-friendly description of this filter
    Dim cName As String
    cName = GetNameFromColorSpaceAndChannel(dstColorSpace, dstChannel)
        
    If (Not toPreview) Then Message "Isolating the %1 channel...", cName
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    Dim cK As Double, mK As Double, yK As Double, bK As Double, invBK As Double
    
    'After all that work, the Rechannel code itself is relatively small and unexciting!
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
    
        Select Case dstColorSpace
        
            'RGB
            Case 0
            
                Select Case dstChannel
                
                    'Rechannel red
                    Case 0
                        imageData(xStride, y) = 0
                        imageData(xStride + 1, y) = 0
                    'Rechannel green
                    Case 1
                        imageData(xStride, y) = 0
                        imageData(xStride + 2, y) = 0
                    'Rechannel blue
                    Case 2
                        imageData(xStride + 1, y) = 0
                        imageData(xStride + 2, y) = 0
                        
                End Select
                
            'CMY
            Case 1
            
                Select Case dstChannel
                
                    'Rechannel cyan
                    Case 0
                        imageData(xStride, y) = 255
                        imageData(xStride + 1, y) = 255
                    'Rechannel magenta
                    Case 1
                        imageData(xStride, y) = 255
                        imageData(xStride + 2, y) = 255
                    'Rechannel yellow
                    Case 2
                        imageData(xStride + 1, y) = 255
                        imageData(xStride + 2, y) = 255
                        
                End Select
            
            'Rechannel CMYK
            Case Else
            
                cK = 255 - imageData(xStride + 2, y)
                mK = 255 - imageData(xStride + 1, y)
                yK = 255 - imageData(xStride, y)
                
                cK = cK / 255
                mK = mK / 255
                yK = yK / 255
                
                bK = Min3Float(cK, mK, yK)
    
                invBK = 1# - bK
                If (invBK = 0#) Then invBK = 0.000001
                
                'cyan
                If (dstChannel = 0) Then
                    cK = ((cK - bK) / invBK) * 255#
                    imageData(xStride + 2, y) = 255 - cK
                    imageData(xStride + 1, y) = 255
                    imageData(xStride, y) = 255
                
                'magenta
                ElseIf (dstChannel = 1) Then
                    mK = ((mK - bK) / invBK) * 255#
                    imageData(xStride + 2, y) = 255
                    imageData(xStride + 1, y) = 255 - mK
                    imageData(xStride, y) = 255
                
                'yellow
                ElseIf (dstChannel = 2) Then
                    yK = ((yK - bK) / invBK) * 255#
                    imageData(xStride + 2, y) = 255
                    imageData(xStride + 1, y) = 255
                    imageData(xStride, y) = 255 - yK
                
                'key
                Else
                    imageData(xStride + 2, y) = invBK * 255
                    imageData(xStride + 1, y) = invBK * 255
                    imageData(xStride, y) = invBK * 255
                End If
                
        End Select
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
        
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then RechannelImage GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'This function displays a user-friendly message with the name of the destination color channel.  Use this function to generate such
' a name from the input color space and channel constants.
Private Function GetNameFromColorSpaceAndChannel(ByVal srcColorSpace As Long, ByVal srcChannel As Long) As String

    Select Case srcColorSpace
    
        'RGB
        Case 0
            Select Case srcChannel
                Case 0
                    GetNameFromColorSpaceAndChannel = g_Language.TranslateMessage("red")
                Case 1
                    GetNameFromColorSpaceAndChannel = g_Language.TranslateMessage("green")
                Case 2
                    GetNameFromColorSpaceAndChannel = g_Language.TranslateMessage("blue")
            End Select
        
        'CMY and CMYK
        Case 1, 2
            Select Case srcChannel
                Case 0
                    GetNameFromColorSpaceAndChannel = g_Language.TranslateMessage("cyan")
                Case 1
                    GetNameFromColorSpaceAndChannel = g_Language.TranslateMessage("magenta")
                Case 2
                    GetNameFromColorSpaceAndChannel = g_Language.TranslateMessage("yellow")
                Case 3
                    GetNameFromColorSpaceAndChannel = g_Language.TranslateMessage("key (black)")
            End Select
        
    End Select

End Function

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("colorspace", btsColorSpace.ListIndex, "channel", btsChannel(btsColorSpace.ListIndex).ListIndex)
End Function
