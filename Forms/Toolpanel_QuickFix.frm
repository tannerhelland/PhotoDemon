VERSION 5.00
Begin VB.Form toolpanel_NDFX 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdButtonToolbox cmdQuickFix 
      Height          =   570
      Index           =   0
      Left            =   13290
      TabIndex        =   1
      Top             =   120
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1005
      AutoToggle      =   -1  'True
   End
   Begin PhotoDemon.sliderTextCombo sltQuickFix 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   0
      Left            =   1530
      TabIndex        =   2
      Top             =   165
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   873
      Min             =   -2
      Max             =   2
      SigDigits       =   2
      SliderTrackStyle=   2
   End
   Begin PhotoDemon.sliderTextCombo sltQuickFix 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   1530
      TabIndex        =   3
      Top             =   780
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
   End
   Begin PhotoDemon.sliderTextCombo sltQuickFix 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   5520
      TabIndex        =   4
      Top             =   165
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
   End
   Begin PhotoDemon.sliderTextCombo sltQuickFix 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   3
      Left            =   5520
      TabIndex        =   5
      Top             =   780
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
   End
   Begin PhotoDemon.sliderTextCombo sltQuickFix 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   4
      Left            =   9660
      TabIndex        =   6
      Top             =   165
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   3
      GradientColorLeft=   16752699
      GradientColorRight=   2990335
      GradientColorMiddle=   16777215
   End
   Begin PhotoDemon.sliderTextCombo sltQuickFix 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   5
      Left            =   9660
      TabIndex        =   7
      Top             =   780
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   3
      GradientColorLeft=   15102446
      GradientColorRight=   8253041
      GradientColorMiddle=   16777215
   End
   Begin PhotoDemon.pdLabel lblOptions 
      Height          =   240
      Index           =   7
      Left            =   8190
      Top             =   270
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "temperature:"
   End
   Begin PhotoDemon.pdLabel lblOptions 
      Height          =   240
      Index           =   6
      Left            =   8190
      Top             =   885
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "tint:"
   End
   Begin PhotoDemon.pdLabel lblOptions 
      Height          =   240
      Index           =   5
      Left            =   4050
      Top             =   885
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "vibrance:"
   End
   Begin PhotoDemon.pdLabel lblOptions 
      Height          =   240
      Index           =   4
      Left            =   4050
      Top             =   270
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "clarity:"
   End
   Begin PhotoDemon.pdLabel lblOptions 
      Height          =   240
      Index           =   3
      Left            =   120
      Top             =   885
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "contrast:"
   End
   Begin PhotoDemon.pdLabel lblOptions 
      Height          =   240
      Index           =   2
      Left            =   120
      Top             =   270
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "exposure:"
   End
   Begin PhotoDemon.pdButtonToolbox cmdQuickFix 
      Height          =   570
      Index           =   1
      Left            =   13290
      TabIndex        =   0
      Top             =   720
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1005
      AutoToggle      =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblOptions 
      Height          =   240
      Index           =   13
      Left            =   12360
      Top             =   270
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "all:"
   End
End
Attribute VB_Name = "toolpanel_NDFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Non-Destructive Effect (NDFX) Tool Panel
'Copyright 2013-2015 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 13/May/15
'Last update: finish migrating all relevant controls to this dedicated form
'
'This form includes all user-editable settings for the "Quick Fix" canvas tools.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Whether or not non-destructive FX can be applied to the image
Private m_NonDestructiveFXAllowed As Boolean

'If external functions want to disable automatic non-destructive FX syncing, then can do so via this function
Public Sub setNDFXControlState(ByVal newNDFXState As Boolean)
    m_NonDestructiveFXAllowed = newNDFXState
End Sub

Private Sub cmdQuickFix_Click(Index As Integer)

    'Do nothing unless an image has been loaded
    If pdImages(g_CurrentImage) Is Nothing Then Exit Sub
    If Not pdImages(g_CurrentImage).loadedSuccessfully Then Exit Sub

    Dim i As Long

    'Regardless of the action we're applying, we start by disabling all auto-refreshes
    setNDFXControlState False
    
    Select Case Index
    
        'Reset quick-fix settings
        Case 0
            
            'Resetting does not affect the Undo/Redo chain, so simply reset all sliders, then redraw the screen
            For i = 0 To sltQuickFix.Count - 1
                
                sltQuickFix(i).Value = 0
                pdImages(g_CurrentImage).getActiveLayer.setLayerNonDestructiveFXState i, 0
                
            Next i
            
        'Make quick-fix settings permanent
        Case 1
            
            'First, make sure at least one or more quick-fixes are active
            If pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState Then
                
                'Now we do something odd; we reset all sliders, then forcibly set an image checkpoint.  This prevents PD's internal
                ' processor from auto-detecting the slider resets and applying *another* entry to the Undo/Redo chain.
                For i = 0 To sltQuickFix.Count - 1
                    sltQuickFix(i).Value = 0
                Next i
                
                'Ask the central processor to permanently apply the quick-fix changes
                Process "Make quick-fixes permanent", , , UNDO_LAYER
                                
            End If
    
    End Select
    
    'After one of these buttons has been used, all quick-fix values will be reset - so we can disable the buttons accordingly.
    For i = 0 To cmdQuickFix.Count - 1
        If cmdQuickFix(i).Enabled Then cmdQuickFix(i).Enabled = False
    Next i
    
    'Re-enable auto-refreshes
    setNDFXControlState True
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub Form_Load()

    'Initialize quick-fix tools
    cmdQuickFix(0).AssignImage "CMDBAR_RESET", , 50
    cmdQuickFix(1).AssignImage "TO_APPLY", , 50
    
    cmdQuickFix(0).AssignTooltip "Reset all quick-fix adjustment values"
    cmdQuickFix(1).AssignTooltip "Make quick-fix adjustments permanent.  This action is never required, but if viewport rendering is sluggish and many quick-fix adjustments are active, it may improve performance."
    
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.setParentForm Me
    lastUsedSettings.loadAllControlValues
    
    'Update everything against the current theme.  This will also set tooltips for various controls.
    UpdateAgainstCurrentTheme
    
    'Allow non-destructive effects
    m_NonDestructiveFXAllowed = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues
    lastUsedSettings.setParentForm Nothing

End Sub

'Non-destructive effect changes will force an immediate redraw of the viewport
Private Sub sltQuickFix_Change(Index As Integer)

    If (Not pdImages(g_CurrentImage) Is Nothing) And m_NonDestructiveFXAllowed Then
        
        'Check the state of the layer's non-destructive FX tracker before making any changes
        Dim initFXState As Boolean
        initFXState = pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState
        
        'The index of sltQuickFix controls aligns exactly with PD's constants for non-destructive effects.  This is by design.
        pdImages(g_CurrentImage).getActiveLayer.setLayerNonDestructiveFXState Index, sltQuickFix(Index).Value
        
        'Redraw the viewport
        Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        'If the layer now has non-destructive effects active, enable the quick fix buttons (if they aren't already)
        Dim i As Long
        
        If pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState Then
        
            For i = 0 To cmdQuickFix.Count - 1
                If Not cmdQuickFix(i).Enabled Then cmdQuickFix(i).Enabled = True
            Next i
        
        Else
            
            For i = 0 To cmdQuickFix.Count - 1
                If cmdQuickFix(i).Enabled Then cmdQuickFix(i).Enabled = False
            Next i
        
        End If
        
        'Even though this action is not destructive, we want to allow the user to save after making non-destructive changes.
        If pdImages(g_CurrentImage).getSaveState(pdSE_AnySave) And (pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState <> initFXState) Then
            pdImages(g_CurrentImage).setSaveState False, pdSE_AnySave
            SyncInterfaceToCurrentImage
        End If
        
    End If

End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) MakeFormPretty is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()

    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    MakeFormPretty Me

End Sub

Private Sub sltQuickFix_GotFocusAPI(Index As Integer)
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_NDFX Index, sltQuickFix(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub sltQuickFix_LostFocusAPI(Index As Integer)
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_NDFX Index, sltQuickFix(Index).Value
End Sub
