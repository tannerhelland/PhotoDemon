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
      MouseIcon       =   "VBP_FormDespeckle.frx":0000
      MousePointer    =   99  'Custom
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
      MouseIcon       =   "VBP_FormDespeckle.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1440
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      MouseIcon       =   "VBP_FormDespeckle.frx":02A4
      MousePointer    =   99  'Custom
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
'©2000-2012 Tanner Helland
'Created: 12/September/11
'Last updated: 12/September/11
'Last update: first build of the form and the custom routine
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
Public Sub Despeckle(ByVal dThreshold As Long)

    Message "Despeckling image..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    
    Dim X2 As Long, Y2 As Long
    
    'To prevent bleeding from the top-left, we need a second array to store our despeckled data
    Dim tArray() As Byte
    ReDim tArray(0 To PicWidthL * 3 + 2, 0 To PicHeightL) As Byte
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tArray(QuickVal + 2, y) = ImageData(QuickVal + 2, y)
        tArray(QuickVal + 1, y) = ImageData(QuickVal + 1, y)
        tArray(QuickVal, y) = ImageData(QuickVal, y)
    Next y
    Next x
        
    Dim refR As Byte, refG As Byte, refB As Byte
    Dim curR As Byte, curG As Byte, curB As Byte
    
    'Loop variable for the despeckle check
    Dim dx As Long
    
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
    
    For x = 1 To PicWidthL - 1
        QuickVal = x * 3
        
    For y = 1 To PicHeightL - 1
        
        'These variables store the color of the current pixel
        refR = ImageData(QuickVal + 2, y)
        refG = ImageData(QuickVal + 1, y)
        refB = ImageData(QuickVal, y)
        
        'Erase despeckle data from the last pixel
        For dx = 0 To dArrayMax
            dArrayR(dx) = 0
            dArrayG(dx) = 0
            dArrayB(dx) = 0
            dArrayCount(dx) = 0
        Next dx
        
        dArrayMax = 0
        dMost = 0
        dMostLoc = 0
        
        For X2 = x - 1 To x + 1
        For Y2 = y - 1 To y + 1
            
            'Ignore the center pixel in the ring (obviously)
            If (X2 <> x) Or (Y2 <> y) Then
            
                curR = ImageData(X2 * 3 + 2, Y2)
                curG = ImageData(X2 * 3 + 1, Y2)
                curB = ImageData(X2 * 3, Y2)
            
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
                                        
                        For dx = 0 To dArrayMax - 1
                    
                            'If this color matches an existing color, increase the count and exit the loop
                            If curR = dArrayR(dx) And curG = dArrayG(dx) And curB = dArrayB(dx) Then
                                dArrayCount(dx) = dArrayCount(dx) + 1
                                If dArrayCount(dx) > dMost Then
                                    dMost = dArrayCount(dx)
                                    dMostLoc = dx
                                    dFoundColor = True
                                End If
                            End If
                    
                        Next dx
                        
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

        Next Y2
        Next X2
        
        If dMost >= dThreshold Then
            tArray(QuickVal + 2, y) = dArrayR(dMostLoc)
            tArray(QuickVal + 1, y) = dArrayG(dMostLoc)
            tArray(QuickVal, y) = dArrayB(dMostLoc)
        End If
        
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    'Transfer the temporary array back into the main array
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        ImageData(QuickVal + 2, y) = tArray(QuickVal + 2, y)
        ImageData(QuickVal + 1, y) = tArray(QuickVal + 1, y)
        ImageData(QuickVal, y) = tArray(QuickVal, y)
    Next y
    Next x
    
    SetImageData
    
    Message "Finished."

End Sub

'Subroutine for removing orphan pixels (otherwise known as "despeckling")
Public Sub QuickDespeckle()

    Message "Despeckling image..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    
    Dim X2 As Long, Y2 As Long
    
    Dim refR As Byte, refB As Byte, refG As Byte
    
    Dim dChecker As Long
    
    For x = 1 To PicWidthL - 1
        QuickVal = x * 3
    For y = 1 To PicHeightL - 1
        
        dChecker = 0
        
        refR = ImageData((x - 1) * 3 + 2, y - 1)
        refG = ImageData((x - 1) * 3 + 1, y - 1)
        refB = ImageData((x - 1) * 3, y - 1)
        
        'Perform a quick check to see if the current pixel matches the one to the above-right; if it does, skip this one.
        If ImageData(QuickVal + 2, y) <> refR Or ImageData(QuickVal + 1, y) <> refG Or ImageData(QuickVal, y) <> refB Then
        
            For X2 = x - 1 To x + 1
            For Y2 = y - 1 To y + 1
                If (X2 <> x - 1) Or (Y2 <> y - 1) Then
                    If (X2 <> x) Or (Y2 <> y) Then
                        If refR = ImageData(X2 * 3 + 2, Y2) And refG = ImageData(X2 * 3 + 1, Y2) And refB = ImageData(X2 * 3, Y2) Then dChecker = dChecker + 1
                    End If
                End If
            Next Y2
            Next X2
            
            If dChecker >= 7 Then
                ImageData(QuickVal + 2, y) = refR
                ImageData(QuickVal + 1, y) = refG
                ImageData(QuickVal, y) = refB
            End If
            
        End If
            
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
    
    Message "Finished."

End Sub
