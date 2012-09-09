VERSION 5.00
Begin VB.Form FormDespeckle 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Advanced Despeckle"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5820
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
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsDespeckle 
      Height          =   255
      Left            =   2160
      Max             =   5
      Min             =   2
      TabIndex        =   1
      Top             =   480
      Value           =   5
      Width           =   3255
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Strong"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   840
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weak"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Despeckle Strength:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1725
   End
End
Attribute VB_Name = "FormDespeckle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Advanced Despeckle Form
'Copyright ©2000-2012 by Tanner Helland
'Created: 12/September/11
'Last updated: 09/September/12
'Last update: rewrote all functions against the new layer class
'
'This advanced despeckle form allows the user to attempt a more vigorous
' despeckling than that allowed by the default routine.  The default routine
' finds pixels surrounded by eight pixels of a single color, and removes them.
' This routine is more nuanced; it compares a pixel to its surrounding pixels,
' then allows the user to specify how many pixels have to differ in color before
' "despeckling" the current pixel (minimum of 4 matching pixels - at highest
' strength).  At its weakest setting, this routine should perform identically to
' the stock despeckle routine.
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
    Process CustomDespeckle, CLng(10 - hsDespeckle.Value)
    Unload Me
End Sub

'Subroutine for advanced removal of pixels that don't match their surroundings
Public Sub Despeckle(ByVal dThreshold As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Despeckling image..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent bleeding from the top-left corner as we perform the despeckling.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, x2 As Long, y2 As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left + 1
    initY = curLayerValues.Top + 1
    finalX = curLayerValues.Right - 1
    finalY = curLayerValues.Bottom - 1
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
        
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Reference RGB values will be used for comparison during the despeckling checks
    Dim refR As Byte, refG As Byte, refB As Byte
    Dim curR As Byte, curG As Byte, curB As Byte
    
    'Loop variable for the despeckle check
    Dim DX As Long
    
    'Whether or not we found this color in our despeckling array
    Dim dFoundColor As Boolean
    
    'dArray is our array of currently discovered colors
    Dim dArrayR(0 To 9) As Byte
    Dim dArrayG(0 To 9) As Byte
    Dim dArrayB(0 To 9) As Byte
    Dim dArrayCount(0 To 9) As Byte
    
    'dArrayMax is the location of the current available spot in the despeckling array
    Dim dArrayMax As Long
    dArrayMax = 8
    
    'dMost is the count of the highest despeckle option, while dMostLoc is the array location for the max
    Dim dMost As Long, dMostLoc As Long
    
    'Despeckle the image
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'These variables store the color of the current pixel
        refR = srcImageData(QuickVal + 2, y)
        refG = srcImageData(QuickVal + 1, y)
        refB = srcImageData(QuickVal, y)
        
        'Erase despeckle data from the last pixel
        For DX = 0 To dArrayMax
            dArrayR(DX) = 0
            dArrayG(DX) = 0
            dArrayB(DX) = 0
            dArrayCount(DX) = 0
        Next DX
        
        dArrayMax = 0
        dMost = 0
        dMostLoc = 0
        
        For x2 = x - 1 To x + 1
            QuickValInner = x * qvDepth
        For y2 = y - 1 To y + 1
            
            'Ignore the center pixel in the ring (obviously)
            If (x2 <> x) Or (y2 <> y) Then
            
                curR = srcImageData(QuickValInner + 2, y2)
                curG = srcImageData(QuickValInner + 1, y2)
                curB = srcImageData(QuickValInner, y2)
            
                'If this pixel matches the center pixel, ignore it
                If refR <> curR Or refG <> curG Or refB <> curB Then
            
                    'If we are here, we can assume that the current pixel does not match the center pixel
                    
                    'First, see if this is our first pixel
                    If dArrayMax = 0 Then
                        dArrayR(0) = curR
                        dArrayG(0) = curG
                        dArrayB(0) = curB
                        dArrayCount(0) = 1
                        dMost = 1
                        dMostLoc = 0
                        dArrayMax = 1
                    Else
                    'If not, scan the despeckle array to see if this color matches any of the others that we've found
                                        
                        dFoundColor = False
                                        
                        For DX = 0 To dArrayMax - 1
                    
                            'If this color matches an existing color, increase the count and exit the loop
                            If curR = dArrayR(DX) And curG = dArrayG(DX) And curB = dArrayB(DX) Then
                                dArrayCount(DX) = dArrayCount(DX) + 1
                                If dArrayCount(DX) > dMost Then
                                    dMost = dArrayCount(DX)
                                    dMostLoc = DX
                                    dFoundColor = True
                                End If
                            End If
                    
                        Next DX
                        
                        'Check to see if this color was found in the array
                        If dFoundColor = False Then
                            
                            'If it wasn't, add it now
                            dArrayR(dArrayMax) = curR
                            dArrayG(dArrayMax) = curG
                            dArrayB(dArrayMax) = curB
                            dArrayCount(dArrayMax) = 1
                            dArrayMax = dArrayMax + 1
                        
                        End If
                        
                    End If
            
                End If
            
            End If

        Next y2
        Next x2
        
        If dMost >= dThreshold Then
            dstImageData(QuickVal + 2, y) = dArrayR(dMostLoc)
            dstImageData(QuickVal + 1, y) = dArrayG(dMostLoc)
            dstImageData(QuickVal, y) = dArrayB(dMostLoc)
        End If
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Now that despeckling is complete, we can erase our temporary layer
    srcLayer.eraseLayer
    Set srcLayer = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Subroutine for removing orphan pixels (otherwise known as "despeckling")
Public Sub QuickDespeckle(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Removing orphaned pixels..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left + 1
    initY = curLayerValues.Top + 1
    finalX = curLayerValues.Right - 1
    finalY = curLayerValues.Bottom - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValTopLeft As Long, QuickValInner As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    'Additional despeckling variables
    Dim x2 As Long, y2 As Long
    Dim refR As Byte, refB As Byte, refG As Byte
    Dim dChecker As Long
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
        QuickValTopLeft = (x - 1) * qvDepth
    For y = initY To finalY
    
        dChecker = 0
        
        refR = ImageData(QuickValTopLeft + 2, y - 1)
        refG = ImageData(QuickValTopLeft + 1, y - 1)
        refB = ImageData(QuickValTopLeft, y - 1)
        
        'Perform a quick check to see if the current pixel matches the one to the above-left; if it does, skip this one
        ' (because orphaned pixels must differ in color from ALL their surrounding pixels)
        If ImageData(QuickVal + 2, y) <> refR Or ImageData(QuickVal + 1, y) <> refG Or ImageData(QuickVal, y) <> refB Then
        
            For x2 = x - 1 To x + 1
                QuickValInner = x2 * 3
            For y2 = y - 1 To y + 1
                If (x2 <> x - 1) Or (y2 <> y - 1) Then
                    If (x2 <> x) Or (y2 <> y) Then
                        If refR = ImageData(QuickValInner + 2, y2) And refG = ImageData(QuickValInner + 1, y2) And refB = ImageData(QuickValInner, y2) Then dChecker = dChecker + 1
                    End If
                End If
            Next y2
            Next x2
            
            If dChecker >= 7 Then
                ImageData(QuickVal + 2, y) = refR
                ImageData(QuickVal + 1, y) = refG
                ImageData(QuickVal, y) = refB
            End If
            
        End If
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub
