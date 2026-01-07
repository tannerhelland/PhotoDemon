VERSION 5.00
Begin VB.UserControl pdCommandBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   637
   ToolboxBitmap   =   "pdCommandBar.ctx":0000
   Begin PhotoDemon.pdButtonToolbox cmdAction 
      Height          =   285
      Index           =   3
      Left            =   1365
      TabIndex        =   6
      Top             =   90
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      AutoToggle      =   -1  'True
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdButton cmdOK 
      Height          =   510
      Left            =   6600
      TabIndex        =   0
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   900
      Caption         =   "OK"
      UseCustomBackgroundColor=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdAction 
      Height          =   570
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1005
      AutoToggle      =   -1  'True
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdDropDown cboPreset 
      Height          =   345
      Left            =   1845
      TabIndex        =   4
      Top             =   195
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   609
      UseCustomBackgroundColor=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdAction 
      Height          =   570
      Index           =   1
      Left            =   690
      TabIndex        =   3
      Top             =   90
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1005
      AutoToggle      =   -1  'True
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdAction 
      Height          =   570
      Index           =   2
      Left            =   5040
      TabIndex        =   5
      Top             =   90
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1005
      AutoToggle      =   -1  'True
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   510
      Left            =   8160
      TabIndex        =   1
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   900
      Caption         =   "Cancel"
      UseCustomBackgroundColor=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdAction 
      Height          =   285
      Index           =   4
      Left            =   1365
      TabIndex        =   7
      Top             =   375
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      AutoToggle      =   -1  'True
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdAction 
      Height          =   570
      Index           =   5
      Left            =   5670
      TabIndex        =   8
      Top             =   90
      Visible         =   0   'False
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1005
      AutoToggle      =   -1  'True
      UseCustomBackColor=   -1  'True
   End
End
Attribute VB_Name = "pdCommandBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Tool Dialog Command Bar custom control
'Copyright 2013-2026 by Tanner Helland
'Created: 14/August/13
'Last updated: 04/November/24
'Last update: allow owner to cancel reset actions (for example, if the user declines a "want to proceed?" prompt)
'
'For the first decade of its life, PhotoDemon relied on a simple OK and CANCEL button at the bottom of each tool dialog.
' These two buttons were dutifully copy+pasted on each new tool, but beyond that they received little attention.
'
'As the program has grown more complex, I have wanted to add a variety of new features to each tool - things like dedicated
' "Help" and "Reset" buttons.  Tool presets.  Maybe even a Randomize button.  Adding so many features to each individual
' tool would be RIDICULOUSLY time-consuming, so rather than do that, I've wrapped all universal tool features into a
' single command bar, which can be dropped onto any new tool form.
'
'This command bar control encapsulates a huge variety of functionality: some obvious, some not.  Tasks this control
' performs for its parent dialog includes:
' - Validating the contents of all UI elements when OK is pressed
' - Hiding and unloading the parent form when OK is pressed and all controls succesfully validate
' - Unloading the parent when Cancel is pressed
' - Saving/loading last-used settings for all standard controls on the parent
' - Automatically resetting control values if no last-used settings exist
' - When Reset is pressed, all standard controls are reset using an elegant system (described in cmdReset comments)
' - Saving, loading, and otherwise managing user-created presets
' - Randomizing all standard controls when Randomize is pressed
' - Suspending effect previewing while operations are being performed, and requesting new previews when relevant
'
'This functionality spares me from writing a great deal of repetitive code in each tool dialog, but it can be
' confusing for developers who can't figure out why PD is capable of certain actions - so be forewarned: if PD
' is "magically" handling things on a tool dialog, it's probably offloading the task to this control.
'
'As of March 2015, the actual business of loading and storing presets is handled by a separate pdToolPreset object.
' Look there for details on how preset files are managed.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Clicking the OK and CANCEL buttons raise their respective events
Public Event OKClick()
Public Event CancelClick()

'Clicking the RESET button raises the corresponding event.  The rules PD uses for resetting controls are explained
' in the cmdReset_Click() sub below.  Additionally, if no last-used settings are found in the Data/Presets folder,
' this event will be automatically triggered when the parent dialog is loaded.
Public Event BeforeResetClick(ByRef cancelReset As Boolean)
Public Event ResetClick()

'Clicking the RANDOMIZE button raises the corresponding event.  Most dialogs won't need to use this event,
' as this control is capable of randomizing all stock PD controls.  But for tool dialogs like Curves,
' where a completely custom interface exists, the parent can use this event to perform their own randomizing
' on custom controls.
Public Event BeforeRandomizeClick()
Public Event RandomizeClick()

'All custom PD controls are auto-validated when OK is pressed.  If other custom items need validation, the OK
' button will trigger this event, which the parent can use to perform additional validation as necessary.
Public Event ExtraValidations()

'After this control has modified other controls on the page (e.g. when Randomize is pressed), it needs to request
' an updated preview from the parent.  This event is used for that; the parent form simply needs to add an
' "updatePreview" call inside.  (I could automate this, but some dialogs - like Resize - do not offer previews,
' so I thought it better to leave the implementation of this event up to the client.)
Public Event RequestPreviewUpdate()

'Certain dialogs (like Curves) use custom user controls whose settings cannot automatically be read/written as
' part of preset data.  For that reason, two events exist that allow the user to add their own information to a
' given preset.  These events are raised whenever a preset needs to be saved or loaded from file (either the
' last-used settings, or some user-saved preset).
Public Event AddCustomPresetData()
Public Event ReadCustomPresetData()

'Like other PD controls, this control raises its own specialized focus events.  If you need to track focus,
' use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Sometimes, for brevity and clarity's sake I use a single dialog for multiple tools (e.g. median/dilate/erode).
' Such forms create a problem when reading/writing presets, because the command bar has no idea which tool is
' currently active, or even that multiple tools exist on the same form.  In the _Load statement of the parent,
' the setToolName function can be called to append a unique tool name to the default one (which is generated from
' the Form title by default).  This separates the presets for each tool on that form.  For example, on the Median
' dialog, I append the name of the current tool to the default name (Median_<name>, e.g. Median_Dilate).
Private m_userSuppliedToolName As String

'Results of extra user validations will be stored here
Private m_userValidationFailed As Boolean

'If the user wants us to postpone a Cancel-initiated unload, for example if they displayed a custom confirmation
' window, this will let us know to suspend the unload for now.
Private m_dontShutdownYet As Boolean

'Each instance of this control lives on a unique tool dialog.  That dialog's name is stored here (automatically
' generated at initialization time).
Private m_parentToolName As String, m_parentToolPath As String

'While the control is loading, this will be set to FALSE.  When the control is ready for interactions, this will be
' set to TRUE.
Private m_controlFullyLoaded As Boolean

'When the user control is in the midst of setting control values, this will be set to FALSE.
Private m_allowPreviews As Boolean

'If the user wants to enable/disable previews, this value will be set.  We will check this in addition to our own
' internal preview checks when requesting previews.
Private m_userAllowsPreviews As Boolean

'When a tool dialog needs to read or write custom preset data (e.g. the Curves dialog, with its unique Curves
' user control), we use these variables to store all custom data supplied to us.
Private m_numCustomPresetEntries As Long
Private m_customPresetNames() As String, m_customPresetData() As String

'If a parent dialog wants to suspend auto-loading last-used settings (e.g. the Resize dialog,
' because last-used settings contain some other image's dimensions), this will be set to TRUE
Private m_suspendLastUsedAutoLoad As Boolean

'If the parent does not want the command bar to auto-unload it when OK or CANCEL is pressed, this will be set to TRUE
Private m_dontAutoUnloadParent As Boolean

'If the caller doesn't want us to reset control values (e.g. when the preset file is damaged or missing),
' they can set this to TRUE.  When set, the caller bears responsibility for restoring control state at first-run.
Private m_dontResetAutomatically As Boolean

'Some dialogs don't need the randomize button.  It can be suspended via property, and the UI will
' automatically reflow as needed.
Private m_hideRandomizeButton As Boolean

'To avoid "Client Site not available (Error 398)", we wait to access certain parent properties until
' Init/ReadProperty events have fired.  (See MSDN: https://msdn.microsoft.com/en-us/library/aa243344(v=vs.60).aspx)
Private m_ParentAvailable As Boolean

'As of March 2015, presets are now handled by a separate class.  This greatly simplifies the complexity of this user control.
Private m_Presets As pdToolPreset

'Individual preset values are parsed using PD's high-performance XML parser
Private m_Params As pdSerialize

'The caller may not want us to auto-load certain values.  That's fine.  They'll notify us
' of any controls they do *not* want loaded via a dedicated function, and we'll store all
' flagged control IDs here.
Private m_NoLoadList As pdStringStack

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Changes to the underlying control are tracked and cached every [n] seconds; this provides a
' poor-man's "undo/redo" feature without the need for complicated state tracking.
Private WithEvents m_Timer As pdTimer
Attribute m_Timer.VB_VarHelpID = -1
Private m_AllowedToCheckUndo As Boolean

'Stack of param strings that holds the list of settings modifications to the current dialog;
' the user can step through these willy-nilly.
Private m_StateStack As pdStringStack
Private m_StackIndex As Long

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDCB_COLOR_LIST
    [_First] = 0
    PDCB_Background = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_CommandBar
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

'The command bar is set to auto-unload its parent object when OK or CANCEL is pressed.  In some instances (e.g. forms prefaced with
' "dialog_", which return a VBMsgBoxResult), this behavior is not desirable.  It can be overridden by setting this property to TRUE.
Public Property Get DontAutoUnloadParent() As Boolean
    DontAutoUnloadParent = m_dontAutoUnloadParent
End Property

Public Property Let DontAutoUnloadParent(ByVal newValue As Boolean)
    m_dontAutoUnloadParent = newValue
    PropertyChanged "DontAutoUnloadParent"
End Property

'Some dialogs (e.g. Resize) may not want us to automatically load their last-used settings, because they need to
' populate the dialog with values unique to the current image.  If this property is set, last-used settings will
' still be saved and made available as a preset, but they WILL NOT be auto-loaded when the parent dialog loads.
Public Property Get DontAutoLoadLastPreset() As Boolean
    DontAutoLoadLastPreset = m_suspendLastUsedAutoLoad
End Property

Public Property Let DontAutoLoadLastPreset(ByVal newValue As Boolean)
    m_suspendLastUsedAutoLoad = newValue
    PropertyChanged "DontAutoLoadLastPreset"
End Property

'Some dialogs (e.g. brush selection) may not want us to automatically reset the dialog when the form is first loaded.
' If this property is set, the caller is 100% responsible for initializing controls.  (Note that, by design, this setting
' does *not* prevent the command bar from firing Reset events in response to UI events.)
Public Property Get DontResetAutomatically() As Boolean
    DontResetAutomatically = m_dontResetAutomatically
End Property

Public Property Let DontResetAutomatically(ByVal newValue As Boolean)
    m_dontResetAutomatically = newValue
    PropertyChanged "DontResetAutomatically"
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

Public Property Get HideRandomizeButton() As Boolean
    HideRandomizeButton = m_hideRandomizeButton
End Property

Public Property Let HideRandomizeButton(ByVal newValue As Boolean)
    m_hideRandomizeButton = newValue
    PropertyChanged "HideRandomizeButton"
End Property

'If multiple tools exist on the same form, the parent can use this in its _Load statement to identify which tool
' is currently active.  The command bar will then limit its preset actions to that tool name alone.
Public Sub SetToolName(ByVal customName As String)
    m_userSuppliedToolName = customName
End Sub

'Because this user control will change the values of dialog controls at run-time, it is necessary to suspend previewing
' while changing values (so that each value change doesn't prompt a preview redraw, and thus slow down the process.)
' This property will be automatically set by this control as necessary, and the parent form can also set it - BUT IF IT
' DOES, IT NEEDS TO RESET IT WHEN DONE, as obviously this control won't know when the parent is finished with its work.
Public Function PreviewsAllowed() As Boolean
    PreviewsAllowed = (m_allowPreviews And m_controlFullyLoaded And m_userAllowsPreviews)
End Function

Public Sub SetPreviewStatus(ByVal newPreviewSetting As Boolean)
    m_userAllowsPreviews = newPreviewSetting
End Sub

'If the user wants to postpone a Cancel-initiated unload for some reason, they can call this function during their
' Cancel event.
Public Sub DoNotUnloadForm()
    m_dontShutdownYet = True
End Sub

'If any user-applied extra validations failed, they can call this sub to notify us, and we will adjust our behavior accordingly
Public Sub ValidationFailed()
    m_userValidationFailed = True
End Sub

'An hWnd is needed for external tooltip handling
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'To support high-DPI settings properly, we expose specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, , True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition , newTop, True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize , newHeight, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

'If you want to request that this control does *NOT* load presets for a given control,
' you must call this function during Form_Load (or somewhere else prior to preset loading)
' and pass any objects that you do *not* want loaded.
Public Sub RequestPresetNoLoad(ByRef srcObject As Object)
    
    'Generate a name for this control using the same schema we use for preset storage
    Dim controlName As String, controlIndex As Long
    controlName = srcObject.Name
    If VBHacks.InControlArray(srcObject) Then controlIndex = srcObject.Index Else controlIndex = -1
    If (controlIndex >= 0) Then controlName = controlName & ":" & controlIndex
    
    'Store the name
    If (m_NoLoadList Is Nothing) Then Set m_NoLoadList = New pdStringStack
    m_NoLoadList.AddString controlName
    
End Sub

'When a preset is selected from the drop-down, load it.  Note that we change the combo box .ListIndex when adding a new preset;
' to prevent this from causing a redraw, we ignore click events if m_allowPreviews is FALSE.
Private Sub cboPreset_Click()
    If (cboPreset.ListIndex > 0) And m_allowPreviews Then LoadPreset cboPreset.List(cboPreset.ListIndex)
End Sub

'Randomize all control values on the page.  This control will automatically handle all standard controls, and a separate
' event is exposed for dialogs that need to do their own randomization (Curves, etc).
Private Sub RandomizeSettings()

    'Disable previews
    m_allowPreviews = False
    
    'Allow the caller to perform any pre-randomization tasks
    RaiseEvent BeforeRandomizeClick
    
    'TODO: rework randomization against pdRandomize, for better results
    Randomize Timer
    
    'By default, controls are randomized according to the following pattern:
    ' 1) If a control is numeric, it will be set to a random value between its Min and Max properties.
    ' 2) Color pickers will be assigned a random color.
    ' 3) Check boxes will be randomly set to checked or unchecked.
    ' 4) Each option button has a 1 in (num of option buttons) chance of being set to TRUE.
    ' 5) Listboxes and dropdowns will be given a random ListIndex value.
    ' 6) Text boxes will be set to a value between -10 and 10.
    ' If other settings are expected or required, they must be set by the client in the RandomizeClick event.
    
    Dim numOfOptionButtons As Long
    numOfOptionButtons = 0
    
    'Count the number of option buttons on the parent form; this will help us randomize properly
    Dim eControl As Object
    For Each eControl In Parent.Controls
        If (TypeOf eControl Is pdRadioButton) Then numOfOptionButtons = numOfOptionButtons + 1
    Next eControl
    
    'Now, pick a random option button to be set as TRUE
    Dim selectedOptionButton As Long
    If (numOfOptionButtons > 0) Then selectedOptionButton = Int(Rnd * numOfOptionButtons)
    numOfOptionButtons = 0
    
    'Iterate through each control on the form.  Check its type, then write out its relevant "value" property.
    Dim controlType As String
    
    For Each eControl In Parent.Controls
        
        controlType = TypeName(eControl)
            
        'How we randomize a control is dependent on its type (obviously).
        Select Case controlType
        
            'Custom PD numeric controls have exposed .Min, .Max, and .Value properties; use them to randomize properly
            Case "pdSlider", "pdSpinner"
                
                Select Case eControl.SigDigits
                
                    Case 0
                        eControl.Value = eControl.Min + Int(Rnd * (eControl.Max - eControl.Min + 1))
                        
                    Case 1, 2
                        eControl.Value = eControl.Min + (Rnd * (eControl.Max - eControl.Min))
                                            
                End Select
            
            Case "pdColorSelector", "pdColorWheel", "pdColorVariants"
                eControl.Color = Rnd * 16777216
            
            'Check boxes have a 50/50 chance of being set to checked
            Case "pdCheckBox"
                eControl.Value = (Int(Rnd * 2) <> 0)
            
            'Option buttons have a 1 in (num of option buttons) chance of being set to TRUE; see code above
            Case "pdRadioButton"
                If numOfOptionButtons = selectedOptionButton Then eControl.Value = True
                numOfOptionButtons = numOfOptionButtons + 1
                
            'Scroll bars use the same rule as other numeric controls
            Case "HScrollBar", "VScrollBar"
                eControl.Value = eControl.Min + Int(Rnd * (eControl.Max - eControl.Min + 1))
            
            'Button strips work like list and drop-down boxes
            Case "pdButtonStrip", "pdButtonStripVertical"
                eControl.ListIndex = Int(Rnd * eControl.ListCount)
                
            'List boxes and combo boxes are assigned a random ListIndex
            Case "pdListBox", "pdListBoxView", "pdListBoxOD", "pdListBoxViewOD", "pdDropDown", "pdDropDownFont"
                
                'Make sure the combo box is not the preset box on this control!
                If (eControl.hWnd <> cboPreset.hWnd) Then
                    eControl.ListIndex = Int(Rnd * eControl.ListCount)
                End If
            
            'Text boxes are set to a random value between -10 and 10
            Case "pdTextBox"
                eControl.Text = Trim$(Str$(-10 + Int(Rnd * 21)))
            
            'pdRandomize controls have built-in randomization capabilities (as you'd expect!)
            Case "pdRandomizeUI"
                eControl.Randomize
        
        End Select
        
    Next eControl
    
    'Finally, raise the RandomizeClick event in case the user needs to do their own randomizing of custom controls
    RaiseEvent RandomizeClick
    
    'For good measure, erase any preset name in the combo box
    cboPreset.ListIndex = 0
    
    'Enable preview
    m_allowPreviews = True
    
    'Request a preview update
    If m_controlFullyLoaded Then RaiseEvent RequestPreviewUpdate
    
End Sub

'Save the current dialog settings as a new preset
Private Function SavePreset() As Boolean

    Message "Saving preset..."
    
    'Prompt the user for a name
    Dim newNameReturn As VbMsgBoxResult
    newNameReturn = Dialogs.PromptNewPreset(m_Presets, Me, UserControl.Parent)
    
    If (newNameReturn = vbOK) Then
    
        'The user may have made one or more changes to the preset object, including adding or deleting presets.
        
        'Start by disabling previews
        m_allowPreviews = False
        
        'Reset the preset names combo box to match any changes the user has made
        LoadAllPresets
        
        'If the user just added a preset, set the combo box index to match the preset they added
        Dim newlyAddedPresetName As String
        newlyAddedPresetName = m_Presets.GetActivePresetName()
        
        If (LenB(newlyAddedPresetName) <> 0) Then
            Dim i As Long
            For i = 0 To cboPreset.ListCount - 1
                If Strings.StringsEqual(newlyAddedPresetName, Trim$(cboPreset.List(i)), True) Then
                    cboPreset.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        
        m_Presets.ClearActivePresetName
        
        'Re-enable previews
        m_allowPreviews = True
        
        Message "Preset saved."
        SavePreset = True
        
    Else
        Message "Preset save canceled."
        SavePreset = False
        Exit Function
    End If
    
End Function

Private Sub cmdAction_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    Select Case Index
    
        'Reset settings
        Case 0
            ResetSettings
        
        'Randomize settings
        Case 1
            RandomizeSettings
        
        'Save new preset
        Case 2
            SavePreset
            
        'Undo
        Case 3
            If (m_StackIndex > 0) Then
                m_StackIndex = m_StackIndex - 1
                LoadPresetFromString m_StateStack.GetString(m_StackIndex)
                UpdateStateStack False, True
            End If
            
        'Redo
        Case 4
            If (m_StackIndex < m_StateStack.GetNumOfStrings() - 1) Then
                m_StackIndex = m_StackIndex + 1
                LoadPresetFromString m_StateStack.GetString(m_StackIndex)
                UpdateStateStack False, True
            End If
            
        'Submit bug report
        Case 5
            Actions.LaunchAction_ByName "help_reportbug"
    
    End Select
    
End Sub

'CANCEL button
Private Sub cmdCancel_Click()
    HandleCancelButton
End Sub

Private Sub HandleCancelButton()

    'The user may have Cancel actions they want to apply - let them do that
    RaiseEvent CancelClick
    
    'If the user asked us to not shutdown yet, obey - otherwise, unload the parent form
    If m_dontShutdownYet Then
        m_dontShutdownYet = False
        Exit Sub
    End If
    
    'Notify the central Interface handler that CANCEL was clicked; this lets other functions bypass a subsequent UI sync
    Interface.NotifyShowDialogResult vbCancel
    
    'Automatically unload our parent, unless the override property is set (as it is in dialogs that return some value)
    If (Not m_dontAutoUnloadParent) Then Unload UserControl.Parent
    
End Sub

'OK button
Private Sub CmdOK_Click()
    HandleOKButton
End Sub

Private Sub HandleOKButton()

    'Automatically validate all relevant controls on the parent object.  This is a huge perk, because it saves us
    ' from having to write validation code individually.
    Dim validateCheck As Boolean
    validateCheck = True
    
    'To validate everything, start by enumerating through each control on our parent form.
    Dim eControl As Object
    For Each eControl In Parent.Controls
        
        'Obviously, we can only validate our own custom objects that have built-in auto-validate functions.
        ' (This is currently limited to controls that support text input; anything else is automatically protected
        ' against bad inputs.)
        If (TypeOf eControl Is pdSlider) Or (TypeOf eControl Is pdSpinner) Or (TypeOf eControl Is pdResize) Then
                
            If (Not eControl.IsValid) Then
                validateCheck = False
                Exit For
            End If
            
        End If
        
    Next eControl
        
    'Raise an extra validation process, which the parent form can use if necessary to check additional controls.
    ' (We do this now because the parent may have a customized way to respond to invalid data - see the Grayscale dialog, for example.)
    RaiseEvent ExtraValidations
    
    'If any validations failed (ours or the client's), terminate further processing
    If m_userValidationFailed Or (Not validateCheck) Then
        m_userValidationFailed = False
        Exit Sub
    End If
    
    'At this point, we are now free to proceed like any normal OK click.
    
    'Immediately deactivate the Undo/Redo timer
    If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
    Set m_Timer = Nothing
    
    'Write the current control values to the XML engine.  These will be loaded the next time the user uses this tool.
    StorePreset
    
    'Save the current window location
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SaveWindowLocation UserControl.Parent, False
    
    'Notify the central Interface handler that OK was clicked; this lets other functions know that a UI sync is required
    Interface.NotifyShowDialogResult vbOK
    
    'Hide the parent form from view
    UserControl.Parent.Visible = False
    
    'Finally, let the user proceed with whatever comes next!
    RaiseEvent OKClick
    
    'When everything is done, unload our parent form (unless the override property is set, as it is by default)
    If (Not m_dontAutoUnloadParent) Then Unload UserControl.Parent
    
End Sub

'RESET button
Private Sub ResetSettings()
    
    'Disable previews
    m_allowPreviews = False
    
    'Allow the caller to perform any pre-randomization tasks
    Dim cancelReset As Boolean: cancelReset = False
    RaiseEvent BeforeResetClick(cancelReset)
    If cancelReset Then Exit Sub
    
    'By default, controls are reset according to the following pattern:
    ' 1) If a numeric control can be set to 0, it will be.
    ' 2) If a numeric control cannot be set to 0, it will be set to its MINIMUM value.
    ' 3) Color pickers will be turned WHITE.
    ' 4) Check boxes will be CHECKED.
    ' 5) The FIRST encountered option button on the dialog will be selected.
    ' 6) The FIRST entry in a listbox or dropdown will be selected.
    ' 7) Text boxes will be set to 0.
    ' If other settings are expected or required, they must be set by the client in the ResetClick event.
    
    Dim controlType As String
    Dim optButtonHasBeenSet As Boolean
    optButtonHasBeenSet = False
    
    Dim controlName As String, controlIndex As Long
    
    'Iterate through each control on the form.  Check its type, then write out its relevant "value" property.
    Dim eControl As Object
    For Each eControl In Parent.Controls
        
        'If the caller has requested that we don't load certain presets, that same behavior
        ' extends to auto-resets.
        Dim okToLoad As Boolean: okToLoad = True
        If (Not m_NoLoadList Is Nothing) Then
            controlName = eControl.Name
            If VBHacks.InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
            If (controlIndex >= 0) Then controlName = controlName & ":" & controlIndex
            If (Not m_NoLoadList Is Nothing) Then okToLoad = (m_NoLoadList.ContainsString(controlName, True) < 0)
        End If
        
        If okToLoad Then
            
            controlType = TypeName(eControl)
                
            'How we reset a control is dependent on its type (obviously).
            Select Case controlType
            
                'Custom PD numeric controls support a built-in RESET property
                Case "pdSlider", "pdSpinner"
                    eControl.Reset
                    
                'Color pickers are turned white
                Case "pdColorSelector", "pdColorWheel", "pdColorVariants"
                    eControl.Color = RGB(255, 255, 255)
                
                'Check boxes are always checked
                Case "pdCheckBox"
                    eControl.Value = True
                
                'The first option button on the page is selected
                Case "pdRadioButton"
                    If (Not optButtonHasBeenSet) Then
                        eControl.Value = True
                        optButtonHasBeenSet = True
                    End If
                    
                'Button strips are set to their first entry
                Case "pdButtonStrip"
                    If (Not eControl.DontAutoReset) Then eControl.ListIndex = 0
                
                Case "pdButtonStripVertical"
                    eControl.ListIndex = 0
                
                'Scroll bars obey the same rules as other numeric controls
                Case "HScrollBar", "VScrollBar"
                    If (eControl.Min <= 0) Then eControl.Value = 0 Else eControl.Value = eControl.Min
                    
                'List boxes and combo boxes are set to their first entry
                Case "pdListBox", "pdListBoxView", "pdListBoxOD", "pdListBoxViewOD", "pdDropDown"
                
                    'Make sure the combo box is not the preset box on this command button!
                    If (eControl.hWnd <> cboPreset.hWnd) Then eControl.ListIndex = 0
                    
                'PD's font combo box is reset to the current system font
                Case "pdDropDownFont"
                    eControl.ListIndex = eControl.ListIndexByString(Fonts.GetUIFontName())
                
                'Text boxes are set to 0
                Case "TextBox", "pdTextBox"
                    eControl.Text = "0"
                    
                'More modern PD controls have built-in reset functionality
                Case "pdMetadataExport", "pdRandomizeUI"
                    eControl.Reset
                    
            End Select
            
        End If
            
    Next eControl
    
    RaiseEvent ResetClick
    
    'For good measure, erase any preset name in the combo box
    cboPreset.ListIndex = 0
    
    'Enable previews
    m_allowPreviews = True
    
    'If the control has finished loading, request a preview update
    If m_controlFullyLoaded Then RaiseEvent RequestPreviewUpdate
    
End Sub

Private Sub m_Timer_Timer()
    If m_AllowedToCheckUndo Then UpdateStateStack
End Sub

Private Sub UpdateStateStack(Optional ByVal isFirstCall As Boolean = False, Optional ByVal forceUIUpdate As Boolean = False)
    
    Dim tmpString As String
    tmpString = GetPresetParamString()
    
    Dim updateNeeded As Boolean
    updateNeeded = False
    
    'If this is the first call to this object, forcibly save the state string to the stack
    If isFirstCall Then
        m_StateStack.ResetStack
        m_StateStack.AddString tmpString
        updateNeeded = True
    
    'If this is *not* the first call to this object, compare the dialog's current state to the
    ' stored state at the present index.  If they differ, add a new string to the stack at
    ' position [current-index + 1].
    Else
        If Strings.StringsNotEqual(tmpString, m_StateStack.GetString(m_StackIndex), False) Then
            m_StateStack.SetNumOfStrings m_StackIndex + 1
            m_StateStack.AddString tmpString
            m_StackIndex = m_StackIndex + 1
            updateNeeded = True
        End If
    End If
    
    'Update Undo/Redo button state based on current stack position
    If updateNeeded Or forceUIUpdate Then
        cmdAction(3).Enabled = (m_StackIndex > 0)
        cmdAction(4).Enabled = (m_StackIndex < m_StateStack.GetNumOfStrings - 1)
    End If
    
End Sub

'This control subclasses some internal PD messages, which is how we support "OK" and "Cancel" shortcuts via
' "Enter" and "Esc" keypresses
Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)

    If (wMsg = WM_PD_DIALOG_NAVKEY) Then
    
        'This is a relevant navigation key!
        
        'Interpret Enter or Space as OK...
        If (wParam = pdnk_Enter) Or (wParam = pdnk_Space) Then
            HandleOKButton
            bHandled = True
            
        '...and Esc as CANCEL.
        ElseIf (wParam = pdnk_Escape) Then
            HandleCancelButton
            bHandled = True
        End If
        
    End If

End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

'If the command bar itself has focus, manually handle Enter/Esc as OK/Cancel events
Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)

    'Interpret Enter or Space as OK...
    If (whichSysKey = pdnk_Enter) Or (whichSysKey = pdnk_Space) Then
        markEventHandled = True
        HandleOKButton
        
    '...and Esc as CANCEL.
    ElseIf (whichSysKey = pdnk_Escape) Then
        markEventHandled = True
        HandleCancelButton
        
    End If

End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub ucSupport_VisibilityChange(ByVal newVisibility As Boolean)
    m_AllowedToCheckUndo = newVisibility
End Sub

Private Sub UserControl_Initialize()

    'Disable certain actions until the control is fully prepped and ready
    m_controlFullyLoaded = False
    m_allowPreviews = False
    m_userAllowsPreviews = True
    
    'Initialize a preset handler
    Set m_Presets = New pdToolPreset
        
    'Validations succeed by default
    m_userValidationFailed = False
    
    'Parent forms will be unloaded by default when pressing Cancel
    m_dontShutdownYet = False
    
    'By default, the user hasn't appended a special name for this instance
    m_userSuppliedToolName = vbNullString
    
    'We don't enable previews yet - that happens after the Show event fires
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    
    'This control can automatically handle "Enter" and "Esc" keypresses coming from its child form.
    ucSupport.SubclassCustomMessage WM_PD_DIALOG_NAVKEY, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCB_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDCommandBar", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
End Sub

Private Sub UserControl_InitProperties()
    DontAutoLoadLastPreset = False
    DontAutoUnloadParent = False
    DontResetAutomatically = False
    HideRandomizeButton = False
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        DontAutoLoadLastPreset = .ReadProperty("AutoloadLastPreset", False)
        DontAutoUnloadParent = .ReadProperty("DontAutoUnloadParent", False)
        DontResetAutomatically = .ReadProperty("DontResetAutomatically", False)
        HideRandomizeButton = .ReadProperty("HideRandomizeButton", False)
    End With
    m_ParentAvailable = True
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_Show()

    'Disable previews
    m_allowPreviews = False
    
    'When the control is first made visible, rebuild individual tooltips using a custom solution
    ' (which allows for linebreaks and theming).
    If PDMain.IsProgramRunning() Then
        
        'Prep a preset file location.  In most cases, this is just the name of the parent form...
        m_parentToolName = Replace$(UserControl.Parent.Name, "Form", vbNullString, , , vbTextCompare)
        
        '...but the caller can also specify a custom name.  This is used when a single PD form handled multiple effects,
        ' like PD's Median/Dilate/Erode implementation.
        If (LenB(m_userSuppliedToolName) <> 0) Then m_parentToolName = m_parentToolName & "_" & m_userSuppliedToolName
        
        'PD stores all preset files in a set preset folder.  This folder is not user-editable.
        m_parentToolPath = UserPrefs.GetPresetPath & m_parentToolName & ".xml"
        
        'If our parent tool has an XML settings file, load it now.  (If one doesn't exist, the preset engine will create
        ' a default one for us.)
        m_Presets.SetPresetFilePath m_parentToolPath, m_parentToolName, Trim$(UserControl.Parent.Caption)
        
        'Populate the preset combo box with any presets found in the file.
        LoadAllPresets
        
        'The XML object is now primed and ready for use.  Look for last-used control settings, and load them if available.
        ' (Update 25 Aug 2014 - check to see if the parent dialog has disabled this behavior.)
        If (Not m_suspendLastUsedAutoLoad) Then
        
            'Attempt to load last-used settings.  If none were found, fire the Reset event, which will supply proper
            ' default values.
            If (Not LoadPreset()) Then
            
                ResetSettings
                
                'Note that the ResetClick event will re-enable previews, so we must forcibly disable them until the
                ' end of this function.
                m_allowPreviews = False
        
            End If
        
        'If the parent dialog doesn't want us to auto-load last-used settings, we still want to request a RESET event to
        ' populate all dialog controls with default values.
        Else
            If (Not m_dontResetAutomatically) Then ResetSettings
            
            'If there are any custom values stored in this dialog, load them using a specially formatted LoadPreset call
            LoadPreset , False
            
            'Reset the "allow previews" tracker
            m_allowPreviews = False
            
        End If
        
        'Start the state-tracking timer and initialize all state-tracking values
        Set m_StateStack = New pdStringStack
        m_StackIndex = 0
        UpdateStateStack True
        
        Set m_Timer = New pdTimer
        
        Const STATE_TRACKING_INTERVAL_MS As Long = 1000
        m_Timer.Interval = STATE_TRACKING_INTERVAL_MS
        m_Timer.StartTimer
        
    End If
        
    'At run-time, give the OK button focus by default.  (Note that using the .Default property to do this will
    ' BREAK THINGS.  .Default overrides catching the Enter key anywhere else in the form, so we cannot do things
    ' like save a preset via Enter keypress, because the .Default control will always eat the Enter keypress.)
    
    'Additional note: some forms may chose to explicitly set focus away from the OK button.  If that happens, the line below
    ' will throw a critical error.  To avoid that, simply ignore any errors that arise from resetting focus.
    On Error GoTo SomethingStoleFocus
    If PDMain.IsProgramRunning() And (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI cmdOK.hWnd

SomethingStoleFocus:
    
    'Enable previews, and request a refresh
    m_controlFullyLoaded = True
    m_allowPreviews = True
    RaiseEvent RequestPreviewUpdate
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "AutoloadLastPreset", m_suspendLastUsedAutoLoad, False
        .WriteProperty "DontAutoUnloadParent", m_dontAutoUnloadParent, False
        .WriteProperty "DontResetAutomatically", m_dontResetAutomatically, False
        .WriteProperty "HideRandomizeButton", m_hideRandomizeButton, False
    End With
End Sub

'This sub will fill the class's pdXML class (xmlEngine) with the values of all controls on this form, and it will store
' those values in the section titled "presetName".
Public Sub StorePreset(Optional ByVal srcPresetName As String = "last-used settings")
    
    'Make sure PD's built-in "last-used settings" text is properly translated
    If (Not g_Language Is Nothing) And Strings.StringsEqual(srcPresetName, "last-used settings", True) Then srcPresetName = g_Language.TranslateMessage("last-used settings")
    srcPresetName = Trim$(srcPresetName)
    
    'An external function handles the actual XML assembly.
    m_Presets.AddPreset srcPresetName, GetPresetParamString(srcPresetName)
    
    'Because the user may still cancel the dialog, we want to request an XML file dump immediately,
    ' so the recently added preset won't be lost.
    m_Presets.WritePresetFile
    
End Sub

'Record the current value of all UI objects on our parent dialog, and return their combined value as an XML string.
' An optional preset name can be passed; note that this gets embedded in the XML, as well.
Private Function GetPresetParamString(Optional ByVal srcPresetName As String = "last-used settings") As String
    
    'Failsafe only; errors are not expected in this function
    On Error GoTo SkipPreset
    
    'Initialize a param handler and initialize it with the passed preset name
    If (m_Params Is Nothing) Then Set m_Params = New pdSerialize
    m_Params.Reset
    If (LenB(srcPresetName) <> 0) Then m_Params.AddParam "fullPresetName", srcPresetName, True
    
    Dim controlName As String, controlType As String, controlValue As String
    Dim controlIndex As Long
    
    'Next, we're going to iterate through each control on the form.  For each control, we're going to assemble two things:
    ' a name (basically, the control name plus its index, if any), and its value.  These are forwarded to the preset manager,
    ' which handles the actual XML storage for each entry.
    Dim eControl As Object
    For Each eControl In Parent.Controls
        
        'Retrieve the control name and index, if any
        controlName = eControl.Name
        If VBHacks.InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
        
        'Reset our control value checker
        controlValue = vbNullString
            
        'Value retrieval must be handled uniquely for each possible control type (including custom PD-specific controls).
        controlType = TypeName(eControl)
        Select Case controlType
        
            'PD-specific sliders, checkboxes, option buttons, and text up/downs return a .Value property
            Case "pdSlider", "pdCheckBox", "pdRadioButton", "pdSpinner", "pdTitle", "pdScrollBar", "pdButtonToolbox"
                controlValue = Str$(eControl.Value)
            
            'List-type objects have a .ListIndex property
            Case "pdButtonStrip", "pdButtonStripVertical"
                controlValue = Str$(eControl.ListIndex)
            
            'Note that we don't store presets for the preset combo box itself!
            Case "pdListBox", "pdListBoxView", "pdListBoxOD", "pdListBoxViewOD", "pdDropDown", "pdDropDownFont"
                If (eControl.hWnd <> cboPreset.hWnd) Then controlValue = Str$(eControl.ListIndex)
                
            'Various PD controls have their own custom "value"-type properties.
            Case "pdColorSelector", "pdColorWheel", "pdColorVariants"
                controlValue = Str$(eControl.Color)
            
            Case "pdBrushSelector"
                controlValue = eControl.Brush
                
            Case "pdPenSelector"
                controlValue = eControl.Pen
                
            Case "pdGradientSelector"
                controlValue = eControl.Gradient
                
            'Text boxes will store a copy of their current text
            Case "pdTextBox"
                controlValue = eControl.Text
                
            'PhotoDemon's resize UC is a special case.  Because it uses multiple properties (despite being
            ' a single control), we must combine its various values into a single string.
            Case "pdResize"
                controlValue = eControl.GetCurrentSettingsAsXML()
                
            'History managers also provide their own XML string
            Case "pdHistory"
                controlValue = eControl.GetHistoryAsString()
                
            'Metadata management controls provide their own XML string
            Case "pdMetadataExport"
                controlValue = eControl.GetMetadataSettings()
            
            Case "pdColorDepth"
                controlValue = eControl.GetAllSettings()
            
            Case "pdPaletteUI"
                controlValue = eControl.SerializeToXML()
                
            Case "pdRandomizeUI"
                controlValue = eControl.Value
                
        End Select
        
        'Remove VB's default padding from the generated string.  (Str() prepends positive numbers with a space)
        If (LenB(controlValue) <> 0) Then controlValue = Trim$(controlValue)
        
        'If the control value still has a non-zero length, add it now
        If (LenB(controlValue) <> 0) Then
            If (controlIndex >= 0) Then
                m_Params.AddParam controlName & ":" & controlIndex, controlValue
            Else
                m_Params.AddParam controlName, controlValue
            End If
        End If
        
    'Continue with the next control on the parent dialog
    Next eControl
    
    'After all controls are handled, we give the caller a chance to write their own custom preset entries.  Most dialogs
    ' don't need this functionality, but those with custom interfaces (such as the Curves dialog, which has its own
    ' special UI requirements) use this to write any additional values to this preset.
    m_numCustomPresetEntries = 0
    RaiseEvent AddCustomPresetData
    
    'If the user added one or more custom preset entries, the custom preset count will be non-zero.
    If (m_numCustomPresetEntries > 0) Then
    
        'Loop through all custom data, and add it one-at-a-time to the preset object
        Dim i As Long
        For i = 0 To m_numCustomPresetEntries - 1
            m_Params.AddParam "custom:" & m_customPresetNames(i), m_customPresetData(i)
        Next i
    
    End If
    
    GetPresetParamString = m_Params.GetParamString()

SkipPreset:

End Function

'This function is called when the user wants to add new preset data to the current preset
Public Sub AddPresetData(ByVal presetName As String, ByVal presetData As String)
    
    'Increase the array size
    ReDim Preserve m_customPresetNames(0 To m_numCustomPresetEntries) As String
    ReDim Preserve m_customPresetData(0 To m_numCustomPresetEntries) As String

    'Add the entries
    m_customPresetNames(m_numCustomPresetEntries) = presetName
    m_customPresetData(m_numCustomPresetEntries) = presetData

    'Increment the custom data count
    m_numCustomPresetEntries = m_numCustomPresetEntries + 1
    
End Sub

'Inside the ReadCustomPresetData event, the caller can call this function to retrieve any custom preset data from
' the active preset.
Public Function RetrievePresetData(ByRef m_customPresetName As String) As String
    
    'For this function, we ignore the boolean return of .retrievePresetValue, and simply let the caller deal with
    ' blank strings if they occur.
    If (Not m_Params Is Nothing) Then RetrievePresetData = m_Params.GetString("custom:" & m_customPresetName)
    
End Function

'This sub will set the values of all controls on this form, using the values stored in the tool's XML file under the
' "presetName" section.  By default, it will look for the last-used settings, as this is its most common request.
Private Function LoadPreset(Optional ByVal srcPresetName As String = "last-used settings", Optional ByVal loadEverything As Boolean = True) As Boolean
    
    'Start by asking the preset engine if the requested preset even exists in the file
    Dim presetExists As Boolean
    presetExists = m_Presets.DoesPresetExist(srcPresetName)
    
    'If the preset exists, continue with the load process
    If presetExists Then
        LoadPreset = LoadPresetFromString(m_Presets.GetPresetXML(srcPresetName), loadEverything)
                
    'If the preset does *not* exist, exit without further processing
    Else
        LoadPreset = False
        Exit Function
    End If
    
End Function

Private Function LoadPresetFromString(ByRef srcString As String, Optional ByVal loadEverything As Boolean = True) As Boolean

    'Copy this preset's XML into a local param evaluator
    If (m_Params Is Nothing) Then Set m_Params = New pdSerialize
    m_Params.SetParamString srcString
    
    'Loading preset values involves (potentially) changing the value of every single object on this form.  To prevent each
    ' of these changes from triggering a full preview redraw, we forcibly suspend previews now.
    m_allowPreviews = False
    
    Dim controlName As String, controlType As String, controlValue As String
    Dim controlIndex As Long
    
    'If parameters allow, iterate through each control on the form and attempt to retrieve its last-used value
    Dim eControl As Object
    
    If loadEverything Then
    
        For Each eControl In Parent.Controls
            
            'Control values are saved by control name, and if it exists, control index.  We start by generating a matching preset
            ' name for this control.
            controlName = eControl.Name
            If VBHacks.InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
            If (controlIndex >= 0) Then controlName = controlName & ":" & controlIndex
            
            Dim okToLoad As Boolean: okToLoad = True
            If (Not m_NoLoadList Is Nothing) Then okToLoad = (m_NoLoadList.ContainsString(controlName, True) < 0)
            
            'See if a preset exists for this control and this particular preset
            If (okToLoad And m_Params.GetStringEx(controlName, controlValue)) Then
                
                'A value for this control exists, and it has been retrieved into controlValue.  We sort handling of this value
                ' by control type, as different controls require different input values (bool, int, etc).
                controlType = TypeName(eControl)
            
                Select Case controlType
                
                    'Sliders and text up/downs allow for floating-point values, so we always cast these returns as doubles
                    Case "pdSlider", "pdSpinner"
                        eControl.Value = CDblCustom(controlValue)
                    
                    'Check boxes use a long (technically a boolean, as PD's custom check box doesn't support a gray state, but for
                    ' backward compatibility with VB check box constants, we cast to a Long)
                    Case "pdCheckBox"
                        eControl.Value = CBool(controlValue)
                    
                    'Option buttons use booleans
                    Case "pdRadioButton"
                        If CBool(controlValue) Then eControl.Value = CBool(controlValue)
                        
                    'Button strips are similar to list boxes, so they use a .ListIndex property
                    Case "pdButtonStrip", "pdButtonStripVertical"
                    
                        'To protect against future changes that modify the number of available entries in a button strip, we always
                        ' validate the list index against the current list count prior to setting it.
                        If (CLng(controlValue) < eControl.ListCount) Then
                            eControl.ListIndex = CLng(controlValue)
                        Else
                            If (eControl.ListCount > 0) Then eControl.ListIndex = eControl.ListCount - 1
                        End If
                    
                    'Various PD controls have their own custom "value"-type properties.
                    Case "pdColorSelector", "pdColorWheel", "pdColorVariants"
                        eControl.Color = CLng(controlValue)
                               
                    Case "pdBrushSelector"
                        eControl.Brush = controlValue
                    
                    Case "pdPenSelector"
                        eControl.Pen = controlValue
                    
                    Case "pdGradientSelector"
                        eControl.Gradient = controlValue
                    
                    'Traditional scroll bar values are cast as Longs, despite them only having Int ranges
                    ' (hopefully the original caller planned for this!)
                    Case "HScrollBar", "VScrollBar"
                        eControl.Value = CLng(controlValue)
                    
                    'List boxes and dropdowns all use a Long-type .ListIndex property
                    Case "pdListBox", "pdListBoxView", "pdListBoxOD", "pdListBoxViewOD", "pdDropDown", "pdDropDownFont"
                    
                        'Validate range before setting
                        If (CLng(controlValue) < eControl.ListCount) Then
                            eControl.ListIndex = CLng(controlValue)
                        Else
                            If (eControl.ListCount > 0) Then eControl.ListIndex = eControl.ListCount - 1
                        End If
                    
                    'Text boxes just take the stored string as-is
                    Case "TextBox", "pdTextBox"
                        eControl.Text = controlValue
                    
                    'pdTitle is just a boolean
                    Case "pdTitle"
                        eControl.Value = CBool(controlValue)
                    
                    Case "pdColorDepth"
                        eControl.SetAllSettings controlValue
                    
                    Case "pdResize"
                        eControl.SetAllSettingsFromXML controlValue
                        
                    'Metadata management controls handle their own XML parsing
                    Case "pdMetadataExport"
                        eControl.SetMetadataSettings controlValue, True
                        
                    'History managers handle their own XML parsing
                    Case "pdHistory"
                        eControl.SetHistoryFromString controlValue
                        
                    Case "pdPaletteUI"
                        eControl.CreateFromXML controlValue
                        
                    Case "pdRandomizeUI"
                        eControl.Value = controlValue
                        
                End Select
    
            End If
        
        'Iterate through the next control
        Next eControl
        
    End If
    
    'Raise the ReadCustomPresetData event.  This allows the caller to retrieve any custom preset data from the file (e.g. data that
    ' does not directly correspond to a traditional control, like the Curves dialog which supports custom curve point data)
    RaiseEvent ReadCustomPresetData
    
    'Re-enable previews
    m_allowPreviews = True
    
    'If the parent dialog is active (e.g. this function is not occurring during the parent dialog's Load process),
    ' request a preview update as the preview has likely changed due to the new control values.
    If m_controlFullyLoaded Then RaiseEvent RequestPreviewUpdate
    
    'This function's return isn't meaningful at present
    LoadPresetFromString = True
        
End Function

'Search the preset file for all valid presets.  This sub doesn't actually load any of the presets - it just adds their
' names to the preset combo box.
Private Sub LoadAllPresets(Optional ByVal newListIndex As Long = 0)

    cboPreset.Clear
    cboPreset.SetAutomaticRedraws False
    
    'We always add one blank entry to the preset combo box, which is selected by default
    cboPreset.AddItem " ", 0

    'Query the preset manager for any available presets.  If found, it will return the number of available presets
    Dim listOfPresets As pdStringStack
    If (m_Presets.GetListOfPresets(listOfPresets) > 0) Then
        
        'Add all discovered presets to the combo box.  Note that we do not use a traditional stack pop here,
        ' as that would cause the preset order to be reversed!
        Dim i As Long
        For i = 0 To listOfPresets.GetNumOfStrings - 1
            cboPreset.AddItem listOfPresets.GetString(i), i + 1, (i = 0) And (listOfPresets.GetNumOfStrings > 1)
        Next i
        
    End If
    
    'When finished, set the requested list index
    cboPreset.SetAutomaticRedraws True
    cboPreset.ListIndex = newListIndex

End Sub

'The command bar's layout is all handled programmatically.  This lets it look good, regardless of the parent form's size or
' the current monitor's DPI setting.
Private Sub UpdateControlLayout()
    
    'Note that error handling is relevant for this control, as the parent hWnd may not be available under all circumstances
    On Error GoTo SkipUpdateLayout
    
    If (Not m_ParentAvailable) Then Exit Sub
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetControlWidth
    bHeight = ucSupport.GetControlHeight
    
    'Force a standard user control size and bottom-alignment
    Dim parentWindowWidth As Long, parentWindowHeight As Long
    parentWindowWidth = g_WindowManager.GetClientWidth(UserControl.Parent.hWnd)
    parentWindowHeight = g_WindowManager.GetClientHeight(UserControl.Parent.hWnd)
    
    Dim moveRequired As Boolean
    If (bHeight <> Interface.FixDPI(50)) Then moveRequired = True
    If (ucSupport.GetControlTop <> parentWindowHeight - Interface.FixDPI(50)) Then moveRequired = True
    
    If moveRequired Then
        ucSupport.RequestNewSize , Interface.FixDPI(50)
        ucSupport.RequestNewPosition 0, parentWindowHeight - ucSupport.GetControlHeight
    End If
    
    'Make the control the same width as its parent
    If PDMain.IsProgramRunning() Then
        
        If (bWidth <> parentWindowWidth) Then ucSupport.RequestNewSize parentWindowWidth
        
        'Right-align the Cancel and OK buttons
        cmdCancel.SetLeft parentWindowWidth - cmdCancel.GetWidth - Interface.FixDPI(8)
        cmdOK.SetLeft cmdCancel.GetLeft - cmdOK.GetWidth - Interface.FixDPI(8)
        
        'Next, we need to align other controls on this dialog under certain circumstances.
        
        'The left-most button (reset, cmdAction(0)) is *always* available; leave it as-is
        
        'The next button (randomize, cmdAction(1)) can be disabled at design-time
        If m_hideRandomizeButton Then
            
            cmdAction(1).Visible = False
            
            'Other controls need to be shifted accordingly.
            
            'Undo/redo for the dialog
            cmdAction(3).SetLeft cmdAction(0).GetLeft + cmdAction(0).GetWidth + Interface.FixDPI(3)
            cmdAction(4).SetLeft cmdAction(3).GetLeft
            
            'Preset dropdown
            cboPreset.SetLeft cmdAction(3).GetLeft + cmdAction(3).GetWidth + Interface.FixDPI(8)
            
            'Save as new preset
            cmdAction(2).SetLeft cboPreset.GetLeft + cboPreset.GetWidth + Interface.FixDPI(4)
            
            'Report a bug (only available in nightly builds)
            cmdAction(5).SetLeft cmdAction(2).GetLeft + cmdAction(2).GetWidth
        
        End If
        
    End If
    
'NOTE: this error catch is important, as VB will attempt to update the user control's size even after the parent has
'       been unloaded, raising error 398 "Client site not available". If we don't catch the error, the compiled .exe
'       will fail every time a command bar is unloaded (e.g. on almost every tool).
SkipUpdateLayout:

End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDCB_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDCB_Background, "Background", IDE_GRAY
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
    
        'When running, we can assign images and tooltips to the image-only command buttons
        If PDMain.IsProgramRunning() Then
            
            Dim cmdButtonImageSize As Long
            cmdButtonImageSize = Interface.FixDPI(24)
            cmdAction(0).AssignImage "generic_reset", , cmdButtonImageSize, cmdButtonImageSize
            cmdAction(1).AssignImage "generic_random", , cmdButtonImageSize, cmdButtonImageSize
            cmdAction(2).AssignImage "generic_savepreset", , cmdButtonImageSize, cmdButtonImageSize
            
            'In nightly builds, a special "submit bug report" button is made available
            If (PD_BUILD_QUALITY = PD_PRODUCTION) Then
                cmdAction(5).Visible = False
            Else
                cmdAction(5).Visible = True
                cmdAction(5).AssignImage "help_reportbug", , cmdButtonImageSize, cmdButtonImageSize
            End If
            
            cmdButtonImageSize = Interface.FixDPI(14)
            cmdAction(3).AssignImage "generic_undo", , cmdButtonImageSize, cmdButtonImageSize
            cmdAction(4).AssignImage "generic_redo", , cmdButtonImageSize, cmdButtonImageSize
            
        End If
        
        'The other buttons on the command-bar do require some explanation.
        cmdAction(0).AssignTooltip UserControls.GetCommonTranslation(pduct_CommandBarReset)
        cmdAction(1).AssignTooltip UserControls.GetCommonTranslation(pduct_CommandBarRandom)
        cmdAction(2).AssignTooltip UserControls.GetCommonTranslation(pduct_CommandBarSavePreset)
        cmdAction(3).AssignTooltip UserControls.GetCommonTranslation(pduct_CommandBarUndo)
        cmdAction(4).AssignTooltip UserControls.GetCommonTranslation(pduct_CommandBarRedo)
        cboPreset.AssignTooltip UserControls.GetCommonTranslation(pduct_CommandBarPresetList)
        
        'The "submit bug" button only needs localization if it's visible
        If (PD_BUILD_QUALITY <> PD_PRODUCTION) Then cmdAction(5).AssignTooltip "This button will open a new issue page in your web browser.  You will need a GitHub account to proceed.", "Submit bug report or feedback"
        
        'Because all controls on the command bar are synchronized against a non-standard backcolor,
        ' we need to make sure any new colors are loaded FIRST
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        
        Dim cbBackgroundColor As Long
        cbBackgroundColor = m_Colors.RetrieveColor(PDCB_Background, Me.Enabled)
        
        'Synchronize the background color of individual controls against the command bar's backcolor
        cmdOK.BackgroundColor = cbBackgroundColor
        cmdCancel.BackgroundColor = cbBackgroundColor
        cmdOK.UpdateAgainstCurrentTheme
        cmdCancel.UpdateAgainstCurrentTheme
        
        cboPreset.BackgroundColor = cbBackgroundColor
        cboPreset.UpdateAgainstCurrentTheme
        
        Dim i As Long
        For i = cmdAction.lBound To cmdAction.UBound
            cmdAction(i).BackColor = cbBackgroundColor
            cmdAction(i).UpdateAgainstCurrentTheme
        Next i
        
    End If
    
End Sub
