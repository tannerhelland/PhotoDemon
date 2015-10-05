VERSION 5.00
Begin VB.Form FormFadeLast 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Fade"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12315
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   821
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdComboBox cboBlendMode 
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   635
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltOpacity 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1270
      Caption         =   "opacity"
      Min             =   1
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "blend mode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   2880
      Width           =   1260
   End
End
Attribute VB_Name = "FormFadeLast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fade Previous Action Dialog
'Copyright 2000-2015 by Tanner Helland
'Created: 13/October/00
'Last updated: 14/April/14
'Last update: give function a full dialog, with variable opacity and blend modes of the user's choosing
'
'This new and improved Fade dialog gives the user a great deal of control over how they blend the results of the latest
' destructive edit with the original layer contents.  Opacity and blend mode can be custom-set, allowing for great
' flexibility when trying to get an edit "just right".
'
'Note that this function relies heavily on the pdUndo class for retrieving data on previous image states.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'To save ourselves some processing time, we're going to load a copy of the relevant Undo data as soon as the dialog loads.
' Any changes to the on-screen settings can use that copy directly, instead of requesting new ones from file.
' Note that we also cache a "current layer DIB" - this is a bit of misnomer, because it is *not necessarily the current
' active layer*.  It is the current state of the layer relevant to the Fade action, which may or may not be the currently
' selected layer.
Dim m_curLayerDIB As pdDIB, m_prevLayerDIB As pdDIB

'These variables will store the layer ID of the relevant layer, and the name of the action being faded (pre-translation,
' so always in English).
Dim m_relevantLayerID As Long
Dim m_actionName As String
    
Private Sub cboBlendMode_Click()
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Fade", , buildParams(sltOpacity, cboBlendMode.ListIndex), UNDO_LAYER
End Sub

Private Sub cmdBar_ResetClick()
    sltOpacity.Value = 50
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Render a preview
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()

    'Suspend previews until the dialog has been fully initialized
    cmdBar.markPreviewStatus False
    
    'Populate the blend mode drop-down
    Interface.PopulateBlendModeComboBox cboBlendMode, BL_NORMAL
    
    'Retrieve a copy of the relevant previous image state
    Set m_prevLayerDIB = New pdDIB
    
    If Not pdImages(g_CurrentImage).undoManager.fillDIBWithLastUndoCopy(m_prevLayerDIB, m_relevantLayerID, m_actionName, False) Then
        
        'Many checks are performed prior to initiating this form, to make sure a valid previous Undo state exists - so this failsafe
        ' code should never trigger.  FYI!
        Debug.Print "WARNING! Fade data could not be retrieved; something went horribly wrong!"
        Unload Me
        
    End If
    
    'Also retrieve a copy of the layer being operated on, as it appears right now; this is faster than re-retrieving a copy
    ' every time we need to redraw the preview box.
    Set m_curLayerDIB = New pdDIB
    m_curLayerDIB.createFromExistingDIB pdImages(g_CurrentImage).getLayerByID(m_relevantLayerID).layerDIB
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Fade the current image against its most recent previous state, using the opacity and blend mode supplied by the user.
Public Sub fxFadeLastAction(ByVal fadeOpacity As Double, ByVal dstBlendMode As LAYER_BLENDMODE, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    'Status bar and message updates are only provided for non-previews.  Also, because PD's central compositor does all the legwork
    ' for this function, and it does not provide detailed progress reports, we use a cheap progress bar estimation method.
    ' (It really shouldn't matter as this function is extremely fast.)
    If Not toPreview Then
        SetProgBarMax 2
        SetProgBarVal 0
        Message "Fading previous action (%1)...", g_Language.TranslateMessage(m_actionName)
    End If
    
    'During a preview, this function will operate on small, preview-sized copies of both the old and current layer states.
    ' This approach allows us to render a much faster preview (vs entire full-size layers).
    Dim curLayerDIBCopy As pdDIB, prevLayerDIBCopy As pdDIB
    
    Dim tmpSafeArray As SAFEARRAY2D
    
    'Retrieve previous layer; note that the method used to retrieve this layer varies according to preview state.
    If toPreview Then
        previewNonStandardImage tmpSafeArray, m_prevLayerDIB, fxPreview, True
        Set prevLayerDIBCopy = New pdDIB
        prevLayerDIBCopy.createFromExistingDIB workingDIB
    Else
        Set prevLayerDIBCopy = m_prevLayerDIB
    End If
    
    'Retrieve current layer (same steps as above)
    If toPreview Then
        previewNonStandardImage tmpSafeArray, m_curLayerDIB, fxPreview, True
        Set curLayerDIBCopy = New pdDIB
        curLayerDIBCopy.createFromExistingDIB workingDIB
    Else
        Set curLayerDIBCopy = m_curLayerDIB
    End If
    
    'All of the hard blending work will be handled by PD's central compositor, which makes our lives much easier!
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    
    'Composite the current image state against the previous image state, using the supplied opacity and blend mode.
    Dim tmpLayerTop As pdLayer, tmpLayerBottom As pdLayer
    Set tmpLayerTop = New pdLayer
    Set tmpLayerBottom = New pdLayer
    
    tmpLayerTop.InitializeNewLayer PDL_IMAGE, , curLayerDIBCopy
    tmpLayerBottom.InitializeNewLayer PDL_IMAGE, , prevLayerDIBCopy
    
    tmpLayerTop.setLayerBlendMode dstBlendMode
    tmpLayerTop.setLayerOpacity fadeOpacity
    
    SetProgBarVal 1
    cComposite.mergeLayers tmpLayerTop, tmpLayerBottom, False
    
    'If this is a preview, draw the composited image to the picture box and exit.
    If toPreview Then
    
        workingDIB.createFromExistingDIB tmpLayerBottom.layerDIB
        finalizeNonstandardPreview fxPreview, True
        
    'If this is not a preview, overwrite the relevant layer's contents, then refresh the interface to match.
    Else
        
        pdImages(g_CurrentImage).getLayerByID(m_relevantLayerID).layerDIB.createFromExistingDIB tmpLayerBottom.layerDIB
        
        'Notify the parent of the change
        pdImages(g_CurrentImage).notifyImageChanged UNDO_LAYER, pdImages(g_CurrentImage).getLayerIndexFromID(m_relevantLayerID)
        
        SyncInterfaceToCurrentImage
        Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        SetProgBarVal 0
        releaseProgressBar
        
        Message "Fade complete."
        
    End If
    
End Sub

'Use this sub to update the on-screen preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then fxFadeLastAction sltOpacity, cboBlendMode.ListIndex, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltOpacity_Change()
    updatePreview
End Sub
