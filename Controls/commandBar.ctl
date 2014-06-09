VERSION 5.00
Begin VB.UserControl commandBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   637
   ToolboxBitmap   =   "commandBar.ctx":0000
   Begin VB.CommandButton cmdRandomize 
      Caption         =   "Randomize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdSavePreset 
      Caption         =   "Save preset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5010
      TabIndex        =   5
      Top             =   120
      Width           =   720
   End
   Begin VB.ComboBox cmbPreset 
      Height          =   360
      Left            =   1800
      TabIndex        =   4
      Top             =   195
      Width           =   3105
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   510
      Left            =   8070
      TabIndex        =   1
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   510
      Left            =   6600
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "commandBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Tool Dialog Command Bar custom control
'Copyright ©2013-2014 by Tanner Helland
'Created: 14/August/13
'Last updated: 08/June/14
'Last update: trim prepended spaces from control values generated via Str()
'
'For the first decade of its life, PhotoDemon relied on a simple OK and CANCEL button at the bottom of each tool dialog.
' These two buttons were dutifully copy+pasted on each new tool, but beyond that they received little attention.
'
'As the program has grown more complex, I have wanted to add a variety of new features to each tool - things like dedicated
' "Help" and "Reset" buttons.  Tool presets.  Maybe even a Randomize button.  Adding each of these features to each tool
' individually would be a RIDICULOUSLY time-consuming task, so rather than do that, I have wrapped all universal tool
' features into a single command bar, which can be dropped onto any new tool form at will.
'
'This command bar control encapsulates a huge variety of functionality: some obvious, some not.  Things this control handles
' for a tool dialog includes:
' - Unloading the parent when Cancel is pressed
' - Validating the contents of all numeric controls when OK is pressed
' - Hiding and unloading the parent form when OK is pressed and all controls succesfully validate
' - Saving/loading last-used settings for all standard controls on the parent
' - Automatically resetting control values if no last-used settings are found
' - When Reset is pressed, all standard controls will be reset using an elegant system (described in cmdReset comments)
' - Saving and loading user-created presets
' - Randomizing all standard controls when Randomize is pressed
' - Suspending effect previewing while operations are being performed, and requesting new previews when relevant
'
'This impressive functionality spares me from writing a great deal of repetitive code in each tool dialog, but it
' can be confusing for developers who can't figure out why PD is capable of certain actions - so be forewarned: if
' PD seems to be "magically" handling things on a tool dialog, it's actually off-loading the heavy lifting to this
' control!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Clicking the OK and CANCEL buttons raise their respective events
Public Event OKClick()
Public Event CancelClick()

'Clicking the RESET button raises the corresponding event.  The rules PD uses for resetting controls are explained
' in the cmdReset_Click() sub below.  Additionally, if no last-used settings are found in the Data/Presets folder,
' this event will be automatically triggered when the parent dialog is loaded.
Public Event ResetClick()

'Clicking the RANDOMIZE button raises the corresponding event.  Most dialogs won't need to use this event, as this
' control is capable of randomizing all stock PD controls.  But for tool dialogs like Curves, where a completely
' custom interface exists, this event can be used by the parent to perform their own randomizing on non-stock
' controls.
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

'Sometimes, for brevity and clarity's sake I use a single dialog for multiple tools (e.g. median/dilate/erode).
' Such forms create a problem when reading/writing presets, because the command bar has no idea which tool is
' currently active, or even that multiple tools exist on the same form.  In the _Load statement of the parent,
' the setToolName function can be called to append a unique tool name to the default one (which is generated from
' the Form title by default).  This separates the presets for each tool on that form.  For example, on the Median
' dialog, I append the name of the current tool to the default name (Median_<name>, e.g. Median_Dilate).
Private userSuppliedToolName As String

'Used to render images onto the command buttons at run-time (doesn't work in the IDE, as a manifest is required)
Private cImgCtl As clsControlImage

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'XML handling (used to save/load presets) is handled through a specialized class
Dim xmlEngine As pdXML

'Font handling for user controls requires some extra work; see below for details
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

'Results of extra user validations will be stored here
Private userValidationFailed As Boolean

'If the user wants us to postpone a Cancel-initiated unload, for example if they displayed a custom confirmation
' window, this will let us know to suspend the unload for now.
Private dontShutdownYet As Boolean

'Each instance of this control lives on a unique tool dialog.  That dialog's name is stored here (automatically
' generated at initialization time).
Private parentToolName As String, parentToolPath As String

'While the control is loading, this will be set to FALSE.  When the control is ready for interactions, this will be
' set to TRUE.
Private controlFullyLoaded As Boolean

'When the user control is in the midst of setting control values, this will be set to FALSE.
Private allowPreviews As Boolean

'If the user wants to enable/disable previews, this value will be set.  We will check this in addition to our own
' internal preview checks when requesting previews.
Private userAllowsPreviews As Boolean

'When a tool dialog needs to read or write custom preset data (e.g. the Curves dialog, with its unique Curves
' user control), we use these variables to store all custom data supplied to us.
Private numUserPresetEntries As Long
Private userPresetNames() As String
Private userPresetData() As String
Private curPresetEntry As String

'If a parent dialog wants to suspend auto-load of last-used settings (e.g. the Resize dialog, because last-used
' settings will be some other image's dimensions), this bool will be set to TRUE
Private suspendLastUsedAutoLoad As Boolean

'When the user presses "Enter" while inside the preset combo box,

'Some dialogs (e.g. Resize) may not want us to automatically load their last-used settings, because they need to
' populate the dialog with values unique to the current image.  If this property is set, last-used settings will
' still be saved and made available as a preset, but they WILL NOT be auto-loaded when the parent dialog loads.
Public Property Get dontAutoLoadLastPreset() As Boolean
    dontAutoLoadLastPreset = suspendLastUsedAutoLoad
End Property

Public Property Let dontAutoLoadLastPreset(ByVal newValue As Boolean)
    suspendLastUsedAutoLoad = newValue
    PropertyChanged "dontAutoLoadLastPreset"
End Property

'If multiple tools exist on the same form, the parent can use this in its _Load statement to identify which tool
' is currently active.  The command bar will then limit its preset actions to that tool name alone.
Public Sub setToolName(ByVal customName As String)
    userSuppliedToolName = customName
End Sub

'Because this user control will change the values of dialog controls at run-time, it is necessary to suspend previewing
' while changing values (so that each value change doesn't prompt a preview redraw, and thus slow down the process.)
' This property will be automatically set by this control as necessary, and the parent form can also set it - BUT IF IT
' DOES, IT NEEDS TO RESET IT WHEN DONE, as obviously this control won't know when the parent is finished with its work.
Public Function previewsAllowed() As Boolean
    previewsAllowed = (allowPreviews And controlFullyLoaded And userAllowsPreviews)
End Function

Public Sub markPreviewStatus(ByVal newPreviewSetting As Boolean)
    
    userAllowsPreviews = newPreviewSetting
    
    'If the client is setting this value to true, it means their work is done - which in turn means we should
    ' request a new preview.
    'DISABLED because it causes endless re-preview loops.  Need to check existing implementations to make sure disabling is okay.
    'If userAllowsPreviews Then RaiseEvent RequestPreviewUpdate
    
End Sub

'If the user wants to postpone a Cancel-initiated unload for some reason, they can call this function during their
' Cancel event.
Public Sub doNotUnloadForm()
    dontShutdownYet = True
End Sub

'If any user-applied extra validations failed, they can call this sub to notify us, and we will adjust our behavior accordingly
Public Sub validationFailed()
    userValidationFailed = True
End Sub

'An hWnd is needed for external tooltip handling
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Font handling is a bit specialized for user controls; see http://msdn.microsoft.com/en-us/library/aa261313%28v=vs.60%29.aspx
Public Property Get Font() As StdFont
    Set Font = mFont
End Property

Public Property Set Font(mNewFont As StdFont)
    
    With mFont
        .Bold = mNewFont.Bold
        .Italic = mNewFont.Italic
        .Name = mNewFont.Name
        .Size = mNewFont.Size
    End With
    PropertyChanged "Font"
    
End Property

'When a preset is selected from the drop-down, load it
Private Sub cmbPreset_Click()
    readXMLSettings cmbPreset.List(cmbPreset.ListIndex)
End Sub

Private Sub cmbPreset_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        savePreset
    End If
End Sub

'Randomize all control values on the page.  This control will automatically handle all standard controls, and a separate
' event is exposed for dialogs that need to do their own randomization (Curves, etc).
Private Sub cmdRandomize_Click()

    'Disable previews
    allowPreviews = False
    
    Randomize Timer
    
    'By default, controls are randomized according to the following pattern:
    ' 1) If a control is numeric, it will be set to a random value between its Min and Max properties.
    ' 2) Color pickers will be assigned a random color.
    ' 3) Check boxes will be randomly set to checked or unchecked.
    ' 4) Each option button has a 1 in (num of option buttons) chance of being set to TRUE.
    ' 5) Listboxes and comboboxes will be given a random ListIndex value.
    ' 6) Text boxes will be set to a value between -10 and 10.
    ' If other settings are expected or required, they must be set by the client in the RandomizeClick event.
    
    Dim numOfOptionButtons As Long
    numOfOptionButtons = 0
    
    'Count the number of option buttons on the parent form; this will help us randomize properly
    Dim eControl As Object
    For Each eControl In Parent.Controls
        If (TypeOf eControl Is smartOptionButton) Then numOfOptionButtons = numOfOptionButtons + 1
    Next eControl
    
    'Now, pick a random option button to be set as TRUE
    Dim selectedOptionButton As Long
    If numOfOptionButtons > 0 Then selectedOptionButton = Int(Rnd * numOfOptionButtons)
    numOfOptionButtons = 0
    
    'Iterate through each control on the form.  Check its type, then write out its relevant "value" property.
    Dim controlType As String
    
    For Each eControl In Parent.Controls
        
        controlType = TypeName(eControl)
            
        'How we randomize a control is dependent on its type (obviously).
        Select Case controlType
        
            'Custom PD numeric controls have exposed .Min, .Max, and .Value properties; use them to randomize properly
            Case "sliderTextCombo", "textUpDown"
                
                Select Case eControl.SigDigits
                
                    Case 0
                        eControl.Value = eControl.Min + Int(Rnd * (eControl.Max - eControl.Min + 1))
                        
                    Case 1, 2
                        eControl.Value = eControl.Min + (Rnd * (eControl.Max - eControl.Min))
                                            
                End Select
            
            Case "colorSelector"
                eControl.Color = Rnd * 16777216
            
            'Check boxes have a 50/50 chance of being set to checked
            Case "smartCheckBox"
                If Int(Rnd * 2) = 0 Then
                    eControl.Value = vbUnchecked
                Else
                    eControl.Value = vbChecked
                End If
            
            'Option buttons have a 1 in (num of option buttons) chance of being set to TRUE; see code above
            Case "smartOptionButton"
                If numOfOptionButtons = selectedOptionButton Then eControl.Value = True
                numOfOptionButtons = numOfOptionButtons + 1
                
            'Scroll bars use the same rule as other numeric controls
            Case "HScrollBar", "VScrollBar"
                eControl.Value = eControl.Min + Int(Rnd * (eControl.Max - eControl.Min + 1))
            
            'List boxes and combo boxes are assigned a random ListIndex
            Case "ListBox", "ComboBox"
            
                'Make sure the combo box is not the preset box on this control!
                If (eControl.hWnd <> cmbPreset.hWnd) Then
                    eControl.ListIndex = Int(Rnd * eControl.ListCount)
                End If
            
            'Text boxes are set to a random value between -10 and 10
            Case "TextBox"
                eControl.Text = Str(-10 + Int(Rnd * 21))
        
        End Select
        
    Next eControl
    
    'Finally, raise the RandomizeClick event in case the user needs to do their own randomizing of custom controls
    RaiseEvent RandomizeClick
    
    'For good measure, erase any preset name in the combo box
    cmbPreset.Text = ""
    
    'Enable preview
    allowPreviews = True
    
    'Request a preview update
    If controlFullyLoaded Then RaiseEvent RequestPreviewUpdate
    
End Sub

'Save the current settings as a new preset
Private Sub cmdSavePreset_Click()
    savePreset
End Sub

'When the user
Private Function savePreset() As Boolean

    Message "Saving preset..."

    'If no name has been entered, prompt the user to do it now.
    If Len(cmbPreset.Text) = 0 Then
        pdMsgBox "Before saving, please enter a name for this preset (in the box next to the save button).", vbInformation + vbOKOnly + vbApplicationModal, "Preset name required"
        cmbPreset.Text = g_Language.TranslateMessage("(enter name here)")
        cmbPreset.SetFocus
        cmbPreset.SelStart = 0
        cmbPreset.SelLength = Len(cmbPreset.Text)
        Message "Preset save canceled."
        savePreset = False
        Exit Function
    End If
    
    'If a name has been entered but it is the same as an existing preset, prompt the user to overwrite.
    Dim overwritingExistingPreset As Boolean
    overwritingExistingPreset = False
    
    Dim i As Long
    For i = 0 To cmbPreset.ListCount - 1
        If (StrComp(cmbPreset.List(i), cmbPreset.Text, vbTextCompare) = 0) Or ((StrComp(xmlEngine.getXMLSafeTagName(cmbPreset.List(i)), xmlEngine.getXMLSafeTagName(cmbPreset.Text), vbTextCompare) = 0)) Then
            
            Dim msgReturn As VbMsgBoxResult
            msgReturn = pdMsgBox("A preset with this name already exists.  Do you want to overwrite it?", vbYesNoCancel + vbApplicationModal + vbInformation, "Overwrite existing preset")
            
            'Based on the user's answer to the confirmation message box, continue or exit
            Select Case msgReturn
            
                'If the user selects YES, continue on like normal
                Case vbYes
                    overwritingExistingPreset = True
                
                'If the user selects NO, exit and let them enter a new name
                Case vbNo
                    cmbPreset.Text = g_Language.TranslateMessage("(enter name here)")
                    cmbPreset.SetFocus
                    cmbPreset.SelStart = 0
                    cmbPreset.SelLength = Len(cmbPreset.Text)
                    Message "Preset save canceled."
                    savePreset = False
                    Exit Function
                
                'If the user selects CANCEL, just exit
                Case vbCancel
                    Message "Preset save canceled."
                    savePreset = False
                    Exit Function
            
            End Select
            
        End If
    Next i
    
    'If we've made it all the way here, the combo box contains the user's desired name for this preset.
    
    'Write the preset out to file.
    fillXMLSettings cmbPreset.Text
    
    'Because the user may still cancel the dialog, we want to request an XML file dump immediately, so
    ' this preset is not lost.
    xmlEngine.writeXMLToFile parentToolPath
    
    'Also, add this preset to the combo box
    If Not overwritingExistingPreset Then
        Dim newPresetName As String
        newPresetName = " " & Trim$(cmbPreset.Text)
        cmbPreset.AddItem newPresetName
    End If
    
    Message "Preset saved."
    
    savePreset = True
    
End Function

'When the font is changed, all controls must manually have their fonts set to match
Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = mFont
    Set cmdOK.Font = mFont
    Set cmdCancel.Font = mFont
    Set cmdReset.Font = mFont
    Set cmdSavePreset.Font = mFont
    Set cmdRandomize.Font = mFont
    Set cmbPreset.Font = mFont
End Sub

'Backcolor is used to control the color of the base user control; nothing else is affected by it
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal newColor As OLE_COLOR)
    UserControl.BackColor = newColor
    PropertyChanged "BackColor"
End Property

'CANCEL button
Private Sub CmdCancel_Click()

    'The user may have Cancel actions they want to apply - let them do that
    RaiseEvent CancelClick
    
    'If the user asked us to not shutdown yet, obey - otherwise, unload the parent form
    If dontShutdownYet Then
        dontShutdownYet = False
        Exit Sub
    End If
    
    'If the current form's progress bar is visible, hide it
    If g_OpenImageCount > 0 Then
        'If pdImages(g_CurrentImage).containingForm.picProgressBar.Visible Then pdImages(g_CurrentImage).containingForm.picProgressBar.Visible = False
    End If
    
    'Automatically unload our parent
    Unload UserControl.Parent
    
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Automatically validate all relevant controls on the parent object.  This is a huge perk, because it saves us
    ' from having to write validation code individually.
    Dim validateCheck As Boolean
    validateCheck = True
    
    'To validate everything, start by enumerating through each control on our parent form.
    Dim eControl As Object
    For Each eControl In Parent.Controls
        
        'Obviously, we can only validate our own custom objects that have built-in auto-validate functions.
        If (TypeOf eControl Is sliderTextCombo) Or (TypeOf eControl Is textUpDown) Or (TypeOf eControl Is smartResize) Then
            
            'Just to be safe, verify matching container hWnd properties
            If eControl.Container.hWnd = UserControl.containerHwnd Then
                
                'Finally, ask the control to validate itself
                If Not eControl.IsValid Then
                    validateCheck = False
                    Exit For
                End If
                
            End If
            
        End If
    Next eControl
        
    'Raise an extra validation process, which the parent form can use if necessary to check additional controls.
    ' (We do this now because the parent may have a customized way to respond to invalid data - see the Grayscale dialog, for example.)
    RaiseEvent ExtraValidations
    
    'If any validations failed (ours or the client's), terminate further processing
    If userValidationFailed Or (Not validateCheck) Then
        userValidationFailed = False
        Exit Sub
    End If
    
    'At this point, we are now free to proceed like any normal OK click.
    
    'Write the current control values to the XML engine.  These will be loaded the next time the user uses this tool.
    fillXMLSettings
    xmlEngine.writeXMLToFile parentToolPath
    
    'Hide the parent form from view
    UserControl.Parent.Visible = False
    
    'Finally, let the user proceed with whatever comes next!
    RaiseEvent OKClick
    
    'When everything is done, unload our parent form
    Unload UserControl.Parent
    
End Sub

'RESET button
Private Sub cmdReset_Click()

    'Disable previews
    allowPreviews = False
    
    'By default, controls are reset according to the following pattern:
    ' 1) If a numeric control can be set to 0, it will be.
    ' 2) If a numeric control cannot be set to 0, it will be set to its MINIMUM value.
    ' 3) Color pickers will be turned WHITE.
    ' 4) Check boxes will be CHECKED.
    ' 5) The FIRST encountered option button on the dialog will be selected.
    ' 6) The FIRST entry in a listbox or combobox will be selected.
    ' 7) Text boxes will be set to 0.
    ' If other settings are expected or required, they must be set by the client in the ResetClick event.
    
    Dim controlType As String
    Dim optButtonHasBeenSet As Boolean
    optButtonHasBeenSet = False
    
    'Iterate through each control on the form.  Check its type, then write out its relevant "value" property.
    Dim eControl As Object
    For Each eControl In Parent.Controls
        
        controlType = TypeName(eControl)
            
        'How we reset a control is dependent on its type (obviously).
        Select Case controlType
        
            'Custom PD numeric controls have exposed .Min, .Max, and .Value properties
            Case "sliderTextCombo", "textUpDown"
                If eControl.Min <= 0 Then
                    eControl.Value = 0
                Else
                    eControl.Value = eControl.Min
                End If
                
            'Color pickers are turned white
            Case "colorSelector"
                eControl.Color = RGB(255, 255, 255)
            
            'Check boxes are always checked
            Case "smartCheckBox"
                eControl.Value = vbChecked
            
            'The first option button on the page is selected
            Case "smartOptionButton"
                If Not optButtonHasBeenSet Then
                    eControl.Value = True
                    optButtonHasBeenSet = True
                End If
            
            'Scroll bars obey the same rules as other numeric controls
            Case "HScrollBar", "VScrollBar"
                If eControl.Min <= 0 Then eControl.Value = 0 Else eControl.Value = eControl.Min
                
            'List boxes and combo boxes are set to their first entry
            Case "ListBox", "ComboBox"
            
                'Make sure the combo box is not the preset box on this control!
                If (eControl.hWnd <> cmbPreset.hWnd) Then
                    eControl.ListIndex = 0
                End If
            
            'Text boxes are set to 0
            Case "TextBox"
                eControl.Text = "0"
        
        End Select
        
    Next eControl
    
    RaiseEvent ResetClick
    
    'For good measure, erase any preset name in the combo box
    cmbPreset.Text = ""
    
    'Enable previews
    allowPreviews = True
    
    'If the control has finished loading, request a preview update
    If controlFullyLoaded Then RaiseEvent RequestPreviewUpdate
    
End Sub

Private Sub UserControl_Initialize()

    'Disable certain actions until the control is fully prepped and ready
    controlFullyLoaded = False
    allowPreviews = False
    userAllowsPreviews = True

    'Apply the hand cursor to all command buttons
    setHandCursorToHwnd cmdOK.hWnd
    setHandCursorToHwnd cmdCancel.hWnd
    setHandCursorToHwnd cmdReset.hWnd
    setHandCursorToHwnd cmdRandomize.hWnd
    setHandCursorToHwnd cmdSavePreset.hWnd

    'Certain actions are only applied in the compiled EXE
    If g_IsProgramCompiled Then
    
        'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
        ' (NOTE: because the icons require manifest theming, they will not appear in the IDE.)
        Set cImgCtl = New clsControlImage
        With cImgCtl
            
            cmdReset.Caption = ""
            .LoadImageFromStream cmdReset.hWnd, LoadResData("RESETBUTTON", "CUSTOM"), fixDPI(24), fixDPI(24)
            .Align(cmdReset.hWnd) = Icon_Center
            
            cmdSavePreset.Caption = ""
            .LoadImageFromStream cmdSavePreset.hWnd, LoadResData("PRESETSAVE", "CUSTOM"), fixDPI(24), fixDPI(24)
            .Align(cmdSavePreset.hWnd) = Icon_Center
            
            cmdRandomize.Caption = ""
            .LoadImageFromStream cmdRandomize.hWnd, LoadResData("RANDOMIZE24", "CUSTOM"), fixDPI(24), fixDPI(24)
            .Align(cmdRandomize.hWnd) = Icon_Center
            
        End With
        
    End If
    
    UserControl.BackColor = BackColor
    
    'Prepare a font object for use
    Set mFont = New StdFont
    Set UserControl.Font = mFont
    
    'Validations succeed by default
    userValidationFailed = False
    
    'Parent forms will be unloaded by default when pressing Cancel
    dontShutdownYet = False
    
    'By default, the user hasn't appended a special name for this instance
    userSuppliedToolName = ""
    
    'We don't enable previews yet - that happens after the Show event fires
    
End Sub

Private Sub UserControl_InitProperties()
    
    Set mFont = UserControl.Font
    mFont_FontChanged ("")
    BackColor = &HEEEEEE
    dontAutoLoadLastPreset = False
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
        BackColor = .ReadProperty("BackColor", &HEEEEEE)
        dontAutoLoadLastPreset = .ReadProperty("AutoloadLastPreset", False)
    End With
    
End Sub

Private Sub UserControl_Resize()
    updateControlLayout
End Sub

'The command bar's layout is all handled programmatically.  This lets it look good, regardless of the parent form's size or
' the current monitor's DPI setting.
Private Sub updateControlLayout()

    On Error GoTo skipUpdateLayout

    'Force a standard user control size
    UserControl.Height = fixDPI(50) * TwipsPerPixelYFix
    
    'Make the control the same width as its parent
    If g_UserModeFix Then
    
        UserControl.Width = UserControl.Parent.ScaleWidth * TwipsPerPixelXFix
        
        'Right-align the Cancel and OK buttons
        cmdCancel.Left = UserControl.Parent.ScaleWidth - cmdCancel.Width - fixDPI(8)
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - fixDPI(8)
        
    End If
    
'NOTE: this error catch is important, as VB will attempt to update the user control's size even after the parent has
'       been unloaded, raising error 398 "Client site not available". If we don't catch the error, the compiled .exe
'       will fail every time a command bar is unloaded (e.g. on almost every tool).
skipUpdateLayout:

End Sub

Private Sub UserControl_Show()

    'Disable previews
    allowPreviews = False
    
    'When the control is first made visible, rebuild individual tooltips using a custom solution
    ' (which allows for linebreaks and theming).
    If g_UserModeFix Then
        
        Set m_ToolTip = New clsToolTip
        With m_ToolTip
        
            .Create Me
            .MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
            .AddTool cmdOK, g_Language.TranslateMessage("Apply this action to the current image.")
            .AddTool cmdCancel, g_Language.TranslateMessage("Exit this tool.  No changes will be made to the image.")
            .AddTool cmdReset, g_Language.TranslateMessage("Reset all settings to their default values.")
            .AddTool cmdRandomize, g_Language.TranslateMessage("Randomly select new settings for this tool.  This is helpful for exploring how different settings affect the image.")
            .AddTool cmdSavePreset, g_Language.TranslateMessage("Save the current settings as a preset.  Please enter a descriptive preset name before saving.")
            .AddTool cmbPreset, g_Language.TranslateMessage("Previously saved presets can be selected here.  You can save the current settings as a new preset by clicking the Save Preset button on the right.")
            
        End With
        
        'Translate all control captions
        cmdOK.Caption = g_Language.TranslateMessage(cmdOK.Caption)
        cmdCancel.Caption = g_Language.TranslateMessage(cmdCancel.Caption)
        
        'In the IDE, we also need to translate the left-hand buttons
        If Not g_IsProgramCompiled Then
            cmdReset.Caption = g_Language.TranslateMessage(cmdReset.Caption)
            cmdRandomize.Caption = g_Language.TranslateMessage(cmdRandomize.Caption)
            cmdSavePreset.Caption = g_Language.TranslateMessage(cmdSavePreset.Caption)
        End If
        
        'If our parent tool has an XML settings file, load it now, and if it doesn't have one, create a blank one
        Set xmlEngine = New pdXML
        parentToolName = Replace$(UserControl.Parent.Name, "Form", "", , , vbTextCompare)
        
        'If the user has supplied a custom name for this tool, append it to the default name
        If Len(userSuppliedToolName) > 0 Then parentToolName = parentToolName & "_" & userSuppliedToolName
        
        parentToolPath = g_UserPreferences.getPresetPath & parentToolName & ".xml"
    
        If FileExist(parentToolPath) Then
            
            'Attempt to load and validate the relevant preset file; if we can't, create a new, blank XML object
            If (Not xmlEngine.loadXMLFile(parentToolPath)) Or Not (xmlEngine.validateLoadedXMLData("toolName")) Then
                Message "This tool's preset file may be corrupted.  A new preset file has been created."
                resetXMLData
            End If
            
        Else
            resetXMLData
        End If
        
        'Populate the preset combo box with any presets found in the file.
        findAllXMLPresets
        
        'The XML object is now primed and ready for use.  Look for last-used control settings, and load them if available.
        ' (Update 25 Aug 2014 - check to see if the parent dialog has disabled this behavior.)
        If Not suspendLastUsedAutoLoad Then
        
            'Attempt to load last-used settings.  If none were found, fire the Reset event, which will supply proper
            ' default values.
            If Not readXMLSettings() Then
            
                cmdReset_Click
                
                'Note that the ResetClick event will re-enable previews, so we must forcibly disable them until the
                ' end of this function.
                allowPreviews = False
        
            End If
        
        'If the parent dialog doesn't want us to auto-load last-used settings, we still want to request a RESET event to
        ' populate all dialog controls with usable values.
        Else
            cmdReset_Click
            allowPreviews = False
        End If
        
    End If
    
    'For now, I'm going to set a standard font size of 10.  May revisit later.
    mFont.Size = 10
    mFont_FontChanged ""
    
    'At run-time, give the OK button focus by default.  (Note that using the .Default property to do this will
    ' BREAK THINGS.  .Default overrides catching the Enter key anywhere else in the form, so we cannot do things
    ' like save a preset via Enter keypress, because the .Default control will always eat the Enter keypress.)
    
    'Additional note: some forms may chose to explicitly set focus away from the OK button.  If that happens, the line below
    ' will throw a critical error.  To avoid that, simply ignore any errors that arise from resetting focus.
    On Error GoTo somethingStoleFocus
    If g_UserModeFix Then cmdOK.SetFocus

somethingStoleFocus:
    
    'Enable previews, and request a refresh
    controlFullyLoaded = True
    allowPreviews = True
    RaiseEvent RequestPreviewUpdate
    
End Sub

'Reset the XML engine for this tool.  Note that the XML object SHOULD ALREADY BE INSTANTIATED before calling this function.
Private Function resetXMLData()

    xmlEngine.prepareNewXML "Tool preset"
    xmlEngine.writeBlankLine
    xmlEngine.writeTag "toolName", parentToolName
    xmlEngine.writeTag "toolDescription", Trim$(UserControl.Parent.Caption)
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "Everything past this point is tool preset data.  Presets are sorted in the order they were created."
    xmlEngine.writeBlankLine

End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    'Store all associated properties
    With PropBag
        .WriteProperty "Font", mFont, "Tahoma"
        .WriteProperty "BackColor", BackColor, &HEEEEEE
        .WriteProperty "AutoloadLastPreset", suspendLastUsedAutoLoad, False
    End With
    
End Sub

'This sub will fill the class's pdXML class (xmlEngine) with the values of all controls on this form, and it will store
' those values in the section titled "presetName".
Private Sub fillXMLSettings(Optional ByVal presetName As String = "last-used settings")
    
    presetName = Trim$(presetName)
    
    'Create an XML-valid preset name here (e.g. remove spaces, etc).  The proper name will still be stored in the file,
    ' but we need a valid tag name for this section, and we need it before doing subsequent processing.
    Dim xmlSafePresetName As String
    xmlSafePresetName = xmlEngine.getXMLSafeTagName(presetName)
    
    'Start by looking for this preset name in the file.  If it does not exist, create a new section for it.
    If Not xmlEngine.doesTagExist("presetEntry", "id", xmlSafePresetName) Then
    
        xmlEngine.writeTagWithAttribute "presetEntry", "id", xmlSafePresetName, "", True
        xmlEngine.writeTag "fullPresetName", presetName
        xmlEngine.closeTag "presetEntry"
        xmlEngine.writeBlankLine
        
    End If
    
    'Iterate through each control on the form.  Check its type, then write out its relevant "value" property.
    Dim controlName As String, controlType As String, controlValue As String
    Dim controlIndex As Long
    
    Dim eControl As Object
    For Each eControl In Parent.Controls
        
        controlName = eControl.Name
        If InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
        controlType = TypeName(eControl)
        controlValue = ""
            
        'We only want to write out the value property of relevant controls.  Check that list now.
        Select Case controlType
        
            'Our custom controls all have a .Value property
            Case "sliderTextCombo", "smartCheckBox", "smartOptionButton", "textUpDown"
                controlValue = Str(eControl.Value)
            
            'Color pickers have a .Color property
            Case "colorSelector"
                controlValue = Str(eControl.Color)
            
            'Intrinsic VB controls may have different names for their value properties
            Case "HScrollBar", "VScrollBar"
                controlValue = Str(eControl.Value)
                
            Case "ListBox", "ComboBox"
            
                'Make sure the combo box is not the preset box on this control!
                If (eControl.hWnd <> cmbPreset.hWnd) Then controlValue = Str(eControl.ListIndex)
                
            Case "TextBox"
                controlValue = eControl.Text
                
            'PhotoDemon's new resize control is a special case.  Because it uses multiple properties (despite being
            ' a single control), we must combine its various values into a single string.
            Case "smartResize"
                controlValue = Str(eControl.imgWidth) & "|" & Str(eControl.imgHeight) & "|" & Str(eControl.lockAspectRatio) _
                                & "|" & Str(eControl.unitOfMeasurement) & "|" & Str(eControl.imgDPI) & "|" & Str(eControl.unitOfResolution)
                
        
        End Select
        
        'Remove VB's default padding from the generated string.  (Str() prepends positive numbers with a space)
        If Len(controlValue) > 0 Then controlValue = Trim$(controlValue)
        
        'If this control has a valid value property, add it to the XML file
        If Len(controlValue) > 0 Then
        
            'If this control is part of a control array, we need to remember its index as well
            If controlIndex >= 0 Then
                xmlEngine.updateTag controlName & ":" & controlIndex, controlValue, "presetEntry", "id", xmlSafePresetName
            Else
                xmlEngine.updateTag controlName, controlValue, "presetEntry", "id", xmlSafePresetName
            End If
        End If
        
    Next eControl
    
    'We assume the user does not have any additional entries
    numUserPresetEntries = 0
    
    'Allow the user to add any custom attributes here
    RaiseEvent AddCustomPresetData
    
    'If the user added any custom preset data, the numUserPresetEntries value will have incremented
    If numUserPresetEntries > 0 Then
    
        'Loop through the user data, and add each entry to the XML file
        Dim i As Long
        For i = 0 To numUserPresetEntries - 1
            xmlEngine.updateTag "custom:" & userPresetNames(i), userPresetData(i), "presetEntry", "id", xmlSafePresetName
        Next i
    
    End If
    
    'We have now added all relevant values to the XML file.
    
End Sub

'This function is called when the user wants to add new preset data to the current preset
Public Function addPresetData(ByVal presetName As String, ByVal presetData As String)
    
    'Increase the array size
    ReDim Preserve userPresetNames(0 To numUserPresetEntries) As String
    ReDim Preserve userPresetData(0 To numUserPresetEntries) As String

    'Add the entries
    userPresetNames(numUserPresetEntries) = presetName
    userPresetData(numUserPresetEntries) = presetData

    'Increment the custom data count
    numUserPresetEntries = numUserPresetEntries + 1
    
End Function

'This function is called when the user wants to read custom preset data from file
Public Function retrievePresetData(ByVal presetName As String) As String
    retrievePresetData = xmlEngine.getUniqueTag_String("custom:" & presetName, "", , "presetEntry", "id", curPresetEntry)
End Function

'This sub will set the values of all controls on this form, using the values stored in the tool's XML file under the
' "presetName" section.  By default, it will look for the last-used settings, as this is its most common request.
Private Function readXMLSettings(Optional ByVal presetName As String = "last-used settings") As Boolean
    
    presetName = Trim$(presetName)
    
    'Disable previews
    allowPreviews = False
    
    'Create an XML-valid preset name here (e.g. remove spaces, etc).  The proper name is stored in the file,
    ' but we need a valid tag name for this section, and we need it before doing subsequent processing.
    Dim xmlSafePresetName As String
    xmlSafePresetName = xmlEngine.getXMLSafeTagName(presetName)
    
    'Start by looking for this preset name in the file.  If it does not exist, abandon this load.
    If Not xmlEngine.doesTagExist("presetEntry", "id", xmlSafePresetName) Then
        readXMLSettings = False
        Exit Function
    End If
    
    'Iterate through each control on the form.  Check its type, then look for a relevant "Value" property in the
    ' saved preset file.
    Dim controlName As String, controlType As String, controlValue As String
    Dim controlIndex As Long
    
    'Some specialty user controls require us to parse out individual values from a lengthy param string
    Dim cParam As pdParamString
    
    Dim eControl As Object
    For Each eControl In Parent.Controls
        
        controlName = eControl.Name
        If InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
        controlType = TypeName(eControl)
        
        'See if an entry exists for this control; note that controls that are part of an array use a unique identifier of the type
        ' controlname:controlindex
        If controlIndex >= 0 Then
            controlValue = xmlEngine.getUniqueTag_String(controlName & ":" & controlIndex, "", , "presetEntry", "id", xmlSafePresetName)
        Else
            controlValue = xmlEngine.getUniqueTag_String(controlName, "", , "presetEntry", "id", xmlSafePresetName)
        End If
        If Len(controlValue) > 0 Then
        
            'An entry exists!  Assign out its value according to the type of control this is.
            Select Case controlType
            
                'Our custom controls all have a .Value property
                Case "sliderTextCombo", "textUpDown"
                    
                    Select Case eControl.SigDigits
                
                        Case 0
                            eControl.Value = CLng(controlValue)
                            
                        Case 1, 2
                            eControl.Value = CDblCustom(controlValue)
                                            
                    End Select
                    
                Case "smartCheckBox"
                    eControl.Value = CLng(controlValue)
                
                Case "smartOptionButton"
                    eControl.Value = CBool(controlValue)
                
                'Color pickers have a .Color property
                Case "colorSelector"
                    eControl.Color = CLng(controlValue)
                
                'Intrinsic VB controls may have different names for their value properties
                Case "HScrollBar", "VScrollBar"
                    eControl.Value = CLng(controlValue)
                    
                Case "ListBox", "ComboBox"
                    If CLng(controlValue) < eControl.ListCount Then
                        eControl.ListIndex = CLng(controlValue)
                    Else
                        If eControl.ListCount > 0 Then eControl.ListIndex = eControl.ListCount - 1
                    End If
                    
                Case "TextBox"
                    eControl.Text = controlValue
                    
                'PD's "smart resize" control has some special needs, on account of using multiple value properties
                ' within a single control.  Parse out those values from the control string.
                Case "smartResize"
                    Set cParam = New pdParamString
                    cParam.setParamString controlValue
                    
                    'Kind of funny, but we must always set the lockAspectRatio to FALSE in order to apply a new size
                    ' to the image.  (If we don't do this, the new sizes will be clamped to the current image's
                    ' aspect ratio!)
                    eControl.lockAspectRatio = False
                    
                    eControl.unitOfMeasurement = cParam.GetLong(4, MU_PIXELS)
                    eControl.unitOfResolution = cParam.GetLong(6, RU_PPI)
                    
                    eControl.imgDPI = cParam.GetLong(5, 96)
                    eControl.imgWidth = cParam.GetDouble(1, 1920)
                    eControl.imgHeight = cParam.GetDouble(2, 1080)
                    
                    Set cParam = Nothing
            
            End Select

        End If
        
    Next eControl
    
    'Allow the user to retrieve any of their custom preset data from the file
    curPresetEntry = xmlSafePresetName
    RaiseEvent ReadCustomPresetData
    
    'We have now filled all controls with their relevant values from the XML file.
    readXMLSettings = True
    
    'Enable previews
    allowPreviews = True
    
    'If the control has finished loading, request a preview update
    If controlFullyLoaded Then RaiseEvent RequestPreviewUpdate
    
End Function

'Search the preset file for all valid presets.  This sub doesn't actually load any of the presets - it just adds their
' names to the preset combo box.
Private Sub findAllXMLPresets()

    cmbPreset.Clear

    'The XML engine will do most the heavy lifting for this task.  We pass it a String array, and it fills it with
    ' all values corresponding to the given tag name and attribute.
    Dim allPresets() As String
    If xmlEngine.findAllAttributeValues(allPresets, "presetEntry", "id") Then
    
        Dim i As Long
        For i = 0 To UBound(allPresets)
            Dim presetToAdd As String
            presetToAdd = " " & xmlEngine.getUniqueTag_String("fullPresetName", , , "presetEntry", "id", allPresets(i))
            cmbPreset.AddItem presetToAdd, i
        Next i
    
    End If
    
    'When finished, clear any active text in the combo box
    cmbPreset.Text = ""

End Sub

'This beautiful little function comes courtesy of coder Merri:
' http://www.vbforums.com/showthread.php?536960-RESOLVED-how-can-i-see-if-the-object-is-array-or-not
Private Function InControlArray(Ctl As Object) As Boolean
    InControlArray = Not Ctl.Parent.Controls(Ctl.Name) Is Ctl
End Function
