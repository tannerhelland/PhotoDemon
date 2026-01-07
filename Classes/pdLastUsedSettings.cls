VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLastUsedSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Auto-Save Form Settings custom class
'Copyright 2013-2026 by Tanner Helland
'Created: 14/September/13
'Last updated: 18/October/21
'Last update: unify settings strategy with pdLastUsedSettings
'
'This class is a spin-off of the (far more involved) command bar user control.  Its purpose is
' to assist with saving and loading the values of all controls on a given form.  When the parent
' form is loaded, this control will search for an XML file with the form's last-used settings.
' If one isn't found, this class will do nothing, but if one IS found, it will restore all
' controls to their last-used values.
'
'As with the command bar user control, hook events are provided so that the parent form can
' easily read/write custom data as well.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Just like the command bar, this user control provides two hooks that the parent can use to
' save/write extra preset data.  These events are raised automatically at load and unload time.
Public Event AddCustomPresetData()
Public Event ReadCustomPresetData()

'XML handling (used to save/load presets) is handled through a specialized class
Private m_Params As pdSerialize

'Each instance of this control must be tied to a unique form.  The form's name is used to
' generate a unique key at initialization time, and that key is used to set/retrieve preset
' data from the central preset file.
Private m_ParentID As String

'If a form needs to read or write custom preset data, we use these variables to store all
' custom data supplied to us.
Private m_numUserPresets As Long
Private m_userPresetNames() As String
Private m_userPresetData() As String

Private m_ParentForm As Object

'BEFORE DOING ANYTHING WITH THIS CLASS, the parent form needs to call this function and
' pass itself as a reference. This class relies heavily on direct access to its parent form -
' so don't forget to do this!  (Similarly, the parent control must pass "Nothing" to this
' same function somewhere before its "Unload" event fires, or circular references will
' cause problems.)
Friend Sub SetParentForm(ByRef ownerForm As Object)
    Set m_ParentForm = ownerForm
End Sub

'The parent form must manually request the restoration of all previous values.  We do this to give the parent time
' to initialize things like drop-downs and other objects.
Friend Function LoadAllControlValues(Optional ByVal customSettingsOnly As Boolean = False) As Boolean
    
    'Failsafe only
    If (m_ParentForm Is Nothing) Then Exit Function
    
    'In previous versions of this control, each instance used a unique presets file inside the /Data/Presets folder.
    ' This was problematic, because unique disk accesses for each main window panel adds up, slowing program launch time.
    
    'To remedy this, v7.0 moved to a new system, where individual last-used values are all cached inside a single
    ' XML file.  This file is managed by the user preferences engine.
    If (m_Params Is Nothing) Then Set m_Params = New pdSerialize
    m_ParentID = Replace$(m_ParentForm.Name, "Form", vbNullString, , , vbTextCompare) & "-presets"
    
    Dim presetPacket As String
    If UserPrefs.GetDialogPresets(m_ParentID, presetPacket) Then
        m_Params.SetParamString presetPacket
    Else
        If (LenB(Trim$(m_ParentForm.Caption)) <> 0) Then m_Params.AddParam "ToolDescription", Trim$(m_ParentForm.Caption)
    End If
    
    'The XML object is now primed and ready for use.  Look for last-used control settings, and load them if available.
    ReadXMLSettings customSettingsOnly
    
End Function

Friend Function SaveAllControlValues(Optional ByVal customSettingsOnly As Boolean = False) As Boolean
    
    FillXMLSettings customSettingsOnly
    
    If UserPrefs.IsReady Then
        
        Dim tmpString As String
        tmpString = m_Params.GetParamString()
        
        'To keep the XML file looking tidy, remove any leading and trailing linebreaks
        Do While (Left$(tmpString, 2) = vbCrLf)
            tmpString = Right$(tmpString, Len(tmpString) - 2)
        Loop
        
        Do While (Right$(tmpString, 2) = vbCrLf)
            tmpString = Left$(tmpString, Len(tmpString) - 2)
        Loop
        
        'Send the tidied-up string to the preset manager
        UserPrefs.SetDialogPresets m_ParentID, vbCrLf & tmpString & vbCrLf
        
    End If
    
End Function

'Fill the m_Params object with the current values of all controls on our parent form.  (Note that some objects (like labels)
' do not support last-used settings, by design.)
Private Sub FillXMLSettings(Optional ByVal customSettingsOnly As Boolean = False)
    
    'Failsafe only
    If (m_ParentForm Is Nothing) Then Exit Sub
    
    'If a dialog only wants us to store custom settings, skip the automatic control loop
    If (Not customSettingsOnly) Then
        
        'Iterate through each control on the form.  Check its type, then determine a relevant way of tracking its "value".
        Dim controlName As String, controlType As String, controlValue As String
        Dim controlIndex As Long
        
        Dim eControl As Object
        For Each eControl In m_ParentForm.Controls
            
            controlName = eControl.Name
            If VBHacks.InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
            controlType = TypeName(eControl)
            controlValue = vbNullString
                
            'We only want to write out the value property of relevant controls.  Check that list now.
            Select Case controlType
            
                'Many PD custom controls have a .Value property
                Case "pdSlider", "pdCheckBox", "pdRadioButton", "pdSpinner", "pdTitle", "pdScrollBar", "pdButtonToolbox"
                    controlValue = Str$(eControl.Value)
                
                'List-type objects have an index
                Case "pdButtonStrip", "pdButtonStripVertical"
                    controlValue = Str$(eControl.ListIndex)
                
                Case "pdListBox", "pdListBoxView", "pdListBoxOD", "pdListBoxViewOD", "pdDropDown", "pdDropDownFont"
                    controlValue = Str$(eControl.ListIndex)
                
                'Custom values, including those returned as strings
                Case "pdColorSelector", "pdColorWheel", "pdColorVariants"
                    controlValue = Str$(eControl.Color)
                
                Case "pdBrushSelector"
                    controlValue = eControl.Brush
                    
                Case "pdPenSelector"
                    controlValue = eControl.Pen
                    
                Case "pdGradientSelector"
                    controlValue = eControl.Gradient
                    
                Case "pdTextBox"
                    controlValue = eControl.Text
                    
                Case "pdResize"
                    controlValue = eControl.GetCurrentSettingsAsXML()
                
                Case "pdHistory"
                    controlValue = eControl.GetHistoryAsString()
                    
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
            
            'If this control has a valid value property, add it to the XML file
            If (LenB(controlValue) <> 0) Then
            
                'If this control is part of a control array, we need to remember its index as well
                If (controlIndex >= 0) Then
                    m_Params.UpdateParam controlName & ":" & controlIndex, controlValue
                Else
                    m_Params.UpdateParam controlName, controlValue
                End If
                
            End If
        
        Next eControl
        
    End If
        
    'Reset the custom value tracker (used for dialogs with non-standard UI elements)
    m_numUserPresets = 0
    
    'Allow our parent to add any custom attributes here
    RaiseEvent AddCustomPresetData
    
    'If the user added any custom preset data, the m_numUserPresets value will have incremented
    If (m_numUserPresets > 0) Then
    
        'Loop through the user data, and add each entry to the XML file
        Dim i As Long
        For i = 0 To m_numUserPresets - 1
            m_Params.UpdateParam "custom:" & m_userPresetNames(i), m_userPresetData(i)
        Next i
    
    End If
    
    'We have now added all relevant values to the XML file.
    
End Sub

'This function is called when the user wants to add new preset data to the current preset
Friend Sub AddPresetData(ByRef presetName As String, ByRef presetData As String)
    
    'Increase the array size
    ReDim Preserve m_userPresetNames(0 To m_numUserPresets) As String
    ReDim Preserve m_userPresetData(0 To m_numUserPresets) As String

    'Add the entries
    m_userPresetNames(m_numUserPresets) = presetName
    m_userPresetData(m_numUserPresets) = presetData

    'Increment the custom data count
    m_numUserPresets = m_numUserPresets + 1
    
End Sub

'Does a preset exist in our list?  Useful for initialization behavior (e.g. creating
' a default preset value if one doesn't exist in the file)
Friend Function DoesPresetExist(ByVal presetName As String) As Boolean
    DoesPresetExist = m_Params.DoesParamExist("custom:" & presetName)
End Function

'This function is called when the user wants to read custom preset data from file
Friend Function RetrievePresetData(ByVal presetName As String, Optional ByVal returnValueIfMissing As String = vbNullString) As String
    RetrievePresetData = m_Params.GetString("custom:" & presetName, returnValueIfMissing)
End Function

'This sub will set the values of all controls on this form, using the values stored in the tool's m_Params object.
Private Function ReadXMLSettings(Optional ByVal customSettingsOnly As Boolean = False) As Boolean
    
    'Some dialogs may only want to store custom data.  If that's the case, skip the entire
    ' automatic settings load loop.
    If (Not customSettingsOnly) Then
        
        'Iterate through each control on the form.  Check its type, then look for a relevant "Value" property in the
        ' saved preset file.
        Dim controlName As String, controlType As String, controlValue As String
        Dim controlIndex As Long
        
        Dim eControl As Object
        For Each eControl In m_ParentForm.Controls
            
            controlName = eControl.Name
            If VBHacks.InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
            controlType = TypeName(eControl)
            
            'See if an entry exists for this control; note that controls that are part of an array use a unique identifier of the type
            ' controlname:controlindex
            If (controlIndex >= 0) Then
                controlValue = m_Params.GetString(controlName & ":" & controlIndex, vbNullString)
            Else
                controlValue = m_Params.GetString(controlName, vbNullString)
            End If
            
            If (LenB(controlValue) <> 0) Then
            
                'An entry exists!  Assign out its value according to the control type.
                Select Case controlType
                
                    'Our custom controls all have a .Value property
                    Case "pdSlider", "pdSpinner"
                        eControl.Value = CDblCustom(controlValue)
                    
                    Case "pdCheckBox", "pdRadioButton", "pdTitle", "pdButtonToolbox"
                        eControl.Value = CBool(controlValue)
                        
                    'Lists need to be checked to make sure we have enough entries to propery assign this value
                    Case "pdButtonStrip", "pdButtonStripVertical", "pdListBox", "pdListBoxView", "pdListBoxOD", "pdListBoxViewOD", "pdDropDown", "pdDropDownFont"
                        If (CLng(controlValue) < eControl.ListCount) Then
                            If (CLng(controlValue) > 0) Then eControl.ListIndex = CLng(controlValue) Else eControl.ListIndex = 0
                        Else
                            If (eControl.ListCount > 0) Then eControl.ListIndex = eControl.ListCount - 1 Else eControl.ListIndex = 0
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
                        
                    Case "pdTextBox"
                        eControl.Text = controlValue
                        
                    Case "pdResize"
                        eControl.SetAllSettingsFromXML controlValue
                        
                    Case "pdHistory"
                        eControl.SetHistoryFromString controlValue
                        
                    Case "pdPaletteUI"
                        eControl.CreateFromXML controlValue
                        
                    Case "pdRandomizeUI"
                        eControl.Value = controlValue
                
                End Select
    
            End If
            
        Next eControl
        
    End If
        
    'Allow the user to retrieve any of their custom preset data from the file
    RaiseEvent ReadCustomPresetData
    
    'We have now filled all controls with their relevant values from the XML file.
    ReadXMLSettings = True
    
End Function

Private Sub Class_Initialize()
    Set m_Params = New pdSerialize
End Sub

Private Sub Class_Terminate()
    Set m_ParentForm = Nothing
End Sub
