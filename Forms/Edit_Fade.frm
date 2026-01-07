VERSION 5.00
Begin VB.Form FormFadeLast 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Fade"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   753
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   735
      Left            =   6000
      TabIndex        =   1
      Top             =   2880
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltOpacity 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1244
      Caption         =   "opacity"
      Min             =   1
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
End
Attribute VB_Name = "FormFadeLast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fade Previous Action Dialog
'Copyright 2000-2026 by Tanner Helland
'Created: 13/October/00
'Last updated: 08/August/17
'Last update: migrate to XML params
'
'This new and improved Fade dialog gives the user a great deal of control over how they blend the results of the latest
' destructive edit with the original layer contents.  Opacity and blend mode can be custom-set, allowing for great
' flexibility when trying to get an edit "just right".
'
'Note that this function relies heavily on the pdUndo class for retrieving data on previous image states.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To save ourselves some processing time, we're going to load a copy of the relevant Undo data as soon as the dialog loads.
' Any changes to the on-screen settings can use that copy directly, instead of requesting new ones from file.
' Note that we also cache a "current layer DIB" - this is a bit of misnomer, because it is *not necessarily the current
' active layer*.  It is the current state of the layer relevant to the Fade action, which may or may not be the currently
' selected layer.
Private m_curLayerDIB As pdDIB, m_prevLayerDIB As pdDIB

'To improve preview performance, we also make local preview-sized copies of each image
Private m_curLayerDIBCopy As pdDIB, m_prevLayerDIBCopy As pdDIB

'These variables will store the layer ID of the relevant layer, and the name of the action being faded (pre-translation,
' so always in English).
Private m_relevantLayerID As Long, m_actionName As String

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Fade", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Suspend previews until the dialog has been fully initialized
    cmdBar.SetPreviewStatus False
    
    'Populate the blend mode drop-down
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    
    'Retrieve a copy of the relevant previous image state
    Set m_prevLayerDIB = New pdDIB
    
    If (Not PDImages.GetActiveImage.UndoManager.FillDIBWithLastUndoCopy(m_prevLayerDIB, m_relevantLayerID, m_actionName, False)) Then
        
        'Many checks are performed prior to initiating this form, to make sure a valid previous Undo state exists - so this failsafe
        ' code should never trigger.  FYI!
        PDDebug.LogAction "WARNING! Fade data could not be retrieved; something went horribly wrong!  Crash imminent!"
        
    End If
    
    'Also retrieve a copy of the layer being operated on, as it appears right now; this is faster than re-retrieving a copy
    ' every time we need to redraw the preview box.
    Set m_curLayerDIB = New pdDIB
    m_curLayerDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetLayerByID(m_relevantLayerID).GetLayerDIB
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Fade the current image against its most recent previous state, using the opacity and blend mode supplied by the user.
Public Sub fxFadeLastAction(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim fadeOpacity As Double, dstBlendMode As PD_BlendMode
    
    With cParams
        fadeOpacity = .GetDouble("opacity", sltOpacity.Value)
        dstBlendMode = .GetLong("blendmode", cboBlendMode.ListIndex)
    End With
    
    'Status bar and message updates are only provided for non-previews.  Also, because PD's central compositor does all the legwork
    ' for this function, and it does not provide detailed progress reports, we use a cheap progress bar estimation method.
    ' (It really shouldn't matter as this function is extremely fast.)
    If (Not toPreview) Then
        ProgressBars.SetProgBarMax 2
        ProgressBars.SetProgBarVal 0
        Message "Fading previous action (%1)...", g_Language.TranslateMessage(m_actionName)
    End If
    
    'During a preview, we only retrieve the portion of each layer that's visible in the current preview box
    If toPreview Then
        Dim tmpSafeArray As SafeArray2D
        
        'Retrieve the preview box portion of the previous layer image
        EffectPrep.ResetPreviewIDs
        PreviewNonStandardImage tmpSafeArray, m_prevLayerDIB, dstPic, True
        If (m_prevLayerDIBCopy Is Nothing) Then Set m_prevLayerDIBCopy = New pdDIB
        m_prevLayerDIBCopy.CreateFromExistingDIB workingDIB
        
        'Retrieve the preview box portion of the current layer image
        EffectPrep.ResetPreviewIDs
        PreviewNonStandardImage tmpSafeArray, m_curLayerDIB, dstPic, True
        If (m_curLayerDIBCopy Is Nothing) Then Set m_curLayerDIBCopy = New pdDIB
        m_curLayerDIBCopy.CreateFromExistingDIB workingDIB
        
    Else
        Set m_prevLayerDIBCopy = m_prevLayerDIB
        Set m_curLayerDIBCopy = m_curLayerDIB
    End If
    
    'All of the hard blending work will be handled by PD's central compositor, which makes our lives much easier!
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    
    If (Not toPreview) Then ProgressBars.SetProgBarVal 1
    cComposite.QuickMergeTwoDibsOfEqualSize m_prevLayerDIBCopy, m_curLayerDIBCopy, dstBlendMode, fadeOpacity
    
    'If this is a preview, draw the composited image to the picture box and exit.
    If toPreview Then
        workingDIB.CreateFromExistingDIB m_prevLayerDIBCopy
        FinalizeNonstandardPreview dstPic, True
        
    'If this is not a preview, overwrite the relevant layer's contents, then refresh the interface to match.
    Else
        
        PDImages.GetActiveImage.GetLayerByID(m_relevantLayerID).GetLayerDIB.CreateFromExistingDIB m_prevLayerDIBCopy
        
        'Notify the parent of the change
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetLayerIndexFromID(m_relevantLayerID)
        
        SyncInterfaceToCurrentImage
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
        ProgressBars.SetProgBarVal 0
        ReleaseProgressBar
        
        Message "Fade complete."
        
    End If
    
End Sub

'Use this sub to update the on-screen preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxFadeLastAction GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltOpacity_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "opacity", sltOpacity.Value
        .AddParam "blendmode", cboBlendMode.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
