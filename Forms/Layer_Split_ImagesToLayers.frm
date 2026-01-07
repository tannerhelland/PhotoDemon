VERSION 5.00
Begin VB.Form FormLayerSplit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Split images into layers"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12510
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
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   834
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStripVertical btsImages 
      Height          =   3495
      Left            =   6480
      TabIndex        =   12
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6165
      Caption         =   "after importing source images"
   End
   Begin PhotoDemon.pdButtonStripVertical btsCanvas 
      Height          =   1695
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2990
      Caption         =   "canvas size"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   6870
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   4680
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   4680
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   4680
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   5280
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   5280
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   5
      Left            =   2160
      TabIndex        =   5
      Top             =   5280
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   5880
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   5880
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   8
      Left            =   2160
      TabIndex        =   8
      Top             =   5880
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdLabel lblAnchor 
      Height          =   285
      Index           =   0
      Left            =   240
      Top             =   4320
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   503
      Caption         =   "align imported layers"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdButtonStripVertical btsLayerNames 
      Height          =   1695
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2990
      Caption         =   "if imported layer names match existing layer names"
   End
   Begin PhotoDemon.pdLabel lblAnchor 
      Height          =   285
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   503
      Caption         =   "destination image"
      FontBold        =   -1  'True
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblAnchor 
      Height          =   285
      Index           =   2
      Left            =   6360
      Top             =   120
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   503
      Caption         =   "source image(s)"
      FontBold        =   -1  'True
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormLayerSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Merge images into single multi-layered image dialog
'Copyright 2019-2026 by Tanner Helland
'Created: 31/August/19
'Last updated: 03/September/19
'Last update: wrap up initial build
'
'To make editing animated images easier, PhotoDemon 8.0 gained the ability to split multi-layer
' images into a multi-image session, then merge those images back into a single multi-layer image.
'
'Merging is fairly complicated, with many possible options for "how to" merge (especially in an animation
' context), and this dialog gives the user some control over the process.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Current anchor position; used to render the anchor selection command buttons, among other things
Private m_CurrentAnchor As Long

'We must also track which arrows are drawn where on the command button array
Private m_ArrowLocations() As String

Private Sub cmdAnchor_Click(Index As Integer)
    m_CurrentAnchor = Index
    UpdateAnchorButtons
End Sub

Private Sub cmdBar_AddCustomPresetData()
    cmdBar.AddPresetData "currentAnchor", Trim$(Str$(m_CurrentAnchor))
End Sub

Private Sub cmdBar_OKClick()
    Process "Split images into layers", False, GetLocalParamString(), UNDO_Image
End Sub

Private Sub cmdBar_RandomizeClick()
    
    Dim cRandomize As pdRandomize
    Set cRandomize = New pdRandomize
    cRandomize.SetSeed_AutomaticAndRandom
    cRandomize.SetRndIntegerBounds 0, 8
    m_CurrentAnchor = cRandomize.GetRandomInt_WH()
    UpdateAnchorButtons
    
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    m_CurrentAnchor = CLng(cmdBar.RetrievePresetData("currentAnchor"))
    UpdateAnchorButtons
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdateAnchorButtons
End Sub

Private Sub cmdBar_ResetClick()

    'Set the middle position as the anchor
    m_CurrentAnchor = 4
    UpdateAnchorButtons
    
End Sub

Private Sub Form_Load()
    
    'Populate any list-style controls
    btsCanvas.AddItem "keep current size", 0
    btsCanvas.AddItem "resize to fit imported layers", 1
    btsCanvas.ListIndex = 0
    
    btsLayerNames.AddItem "replace existing layers with imported ones", 0
    btsLayerNames.AddItem "add duplicates as new layers", 1
    btsLayerNames.ListIndex = 0
    
    btsImages.AddItem "leave images open", 0
    btsImages.AddItem "close images normally (prompt for unsaved changes)", 1
    btsImages.AddItem "close images without prompting", 2
    btsImages.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me

    'Update the anchor button layout
    UpdateAnchorButtons
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub FillArrowLocations(ByRef aLocations() As String)

    'Start with the current position.  It's the easiest one to fill
    aLocations(m_CurrentAnchor) = "generic_image"
    
    'Next, fill in upward arrows as necessary
    If (m_CurrentAnchor > 2) Then
        aLocations(m_CurrentAnchor - 3) = "arrow_up"
        If ((m_CurrentAnchor Mod 3) <> 0) Then aLocations(m_CurrentAnchor - 4) = "arrow_upl"
        If (((m_CurrentAnchor + 1) Mod 3) <> 0) Then aLocations(m_CurrentAnchor - 2) = "arrow_upr"
    End If
    
    'Next, fill in left/right arrows as necessary
    If (((m_CurrentAnchor + 1) Mod 3) <> 0) Then aLocations(m_CurrentAnchor + 1) = "arrow_right"
    If ((m_CurrentAnchor Mod 3) <> 0) Then aLocations(m_CurrentAnchor - 1) = "arrow_left"
    
    'Finally, fill in downward arrows as necessary
    If (m_CurrentAnchor < 6) Then
        aLocations(m_CurrentAnchor + 3) = "arrow_down"
        If ((m_CurrentAnchor Mod 3) <> 0) Then aLocations(m_CurrentAnchor + 2) = "arrow_downl"
        If (((m_CurrentAnchor + 1) Mod 3) <> 0) Then aLocations(m_CurrentAnchor + 4) = "arrow_downr"
    End If
    
End Sub

'The user can use an array of command buttons to specify the image's anchor position on the new canvas.  I adopted this
' model from comparable tools in Photoshop and Paint.NET, among others.  Images are loaded from the resource section
' of the EXE and applied to the command buttons as necessary.
Private Sub UpdateAnchorButtons()
    
    Dim i As Long
    
    'Build an array that contains the arrow to appear in each location.
    ReDim m_ArrowLocations(0 To 8) As String
    FillArrowLocations m_ArrowLocations
    
    Dim dibSize As Long
    dibSize = Interface.FixDPI(24)
                
    'Next, extract relevant icons from the resource file, and render them onto the buttons at run-time.
    For i = 0 To 8
    
        If (LenB(m_ArrowLocations(i)) <> 0) Then
            If (StrComp(m_ArrowLocations(i), "generic_image", vbBinaryCompare) = 0) Then
                cmdAnchor(i).AssignImage m_ArrowLocations(i), , dibSize, dibSize
            Else
                
                Dim tmpDIB As pdDIB
                
                If (StrComp(m_ArrowLocations(i), "arrow_up", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowUp, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_upr", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowUpR, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_right", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowRight, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_downr", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowDownR, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_down", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowDown, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_downl", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowDownL, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_left", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowLeft, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_upl", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowUpL, dibSize)
                End If
                
                cmdAnchor(i).AssignImage vbNullString, tmpDIB, dibSize, dibSize
                    
            End If
            
        Else
            cmdAnchor(i).AssignImage vbNullString, Nothing
        End If
        
    Next i
    
End Sub

Private Function GetLocalParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        cParams.AddParam "overwrite-layers", CBool(btsLayerNames.ListIndex = 0)
        cParams.AddParam "resize-canvas-fit", CBool(btsCanvas.ListIndex = 1)
        cParams.AddParam "layer-anchor", m_CurrentAnchor
        cParams.AddParam "close-source-images", btsImages.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()

End Function

