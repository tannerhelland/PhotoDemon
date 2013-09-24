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
'Copyright ©2012-2013 by Tanner Helland
'Created: 14/September/13
'Last updated: 14/September/13
'Last update: initial build
'
'This class is a spin-off of the (far more involved) command bar user control.  Its purpose is to assist with saving
' and loading the values of all controls on a given form.  When the parent form is loaded, this control will search
' for an XML file with the form's last-used settings.  If none is found, it will do nothing, but if one IS found, it
' will restore all controls to their last-used values.
'
'As with the command bar user control, hook events are provided so that the parent form can easily read/write its
' own custom data as well.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Just like the command bar, this user control provides two hooks that the parent can use to save/write extra
' preset data.  These events are raised automatically at load and unload time.
Public Event AddCustomPresetData()
Public Event ReadCustomPresetData()

'XML handling (used to save/load presets) is handled through a specialized class
Private xmlEngine As pdXML

'Each instance of this control lives on a unique tool dialog.  That dialog's name is stored here (automatically
' generated at initialization time).
Private parentToolName As String, parentToolPath As String

'If a form needs to read or write custom preset data , we use these variables to store all custom data supplied to us.
Private numUserPresetEntries As Long
Private userPresetNames() As String
Private userPresetData() As String
Private curPresetEntry As String

Private parentForm As Object

'BEFORE DOING ANYTHING WITH THIS CLASS, the parent form needs to call this function and pass itself as a reference.
' This class relies heavily on direct access to its parent form - so don't forget to do this!
Public Sub setParentForm(ByRef ownerForm As Object)
    Set parentForm = ownerForm
End Sub

'The parent form must manually request the restoration of all previous values.  We do this to give the parent time
' to initialize things like drop-downs and other objects.
Public Function loadAllControlValues() As Boolean

    'Start by looking for a preset file for this form
    Set xmlEngine = New pdXML
    parentToolName = Replace$(parentForm.Name, "Form", "", , , vbTextCompare)
    parentToolPath = g_UserPreferences.getPresetPath & parentToolName & ".xml"
    
    If FileExist(parentToolPath) Then
            
        'Attempt to load and validate the relevant preset file; if we can't, create a new, blank XML object
        If (Not xmlEngine.loadXMLFile(parentToolPath)) Or Not (xmlEngine.validateLoadedXMLData("toolName")) Then
            Message "This form's preset file may be corrupted.  A new preset file has been created."
            resetXMLData
        End If
            
    Else
        resetXMLData
    End If
        
    'Populate the preset combo box with any presets found in the file.
    'findAllXMLPresets
    
    'The XML object is now primed and ready for use.  Look for last-used control settings, and load them if available.
    readXMLSettings
    
End Function

Public Function saveAllControlValues() As Boolean

    'Write the current control values to the XML engine.  These will be loaded the next time the user uses this tool.
    fillXMLSettings
    xmlEngine.writeXMLToFile parentToolPath

End Function

'Reset the XML engine for this tool.  Note that the XML object SHOULD ALREADY BE INSTANTIATED before calling this function.
Private Function resetXMLData()

    xmlEngine.prepareNewXML "Tool preset"
    xmlEngine.writeBlankLine
    xmlEngine.writeTag "toolName", parentToolName
    xmlEngine.writeTag "toolDescription", Trim$(parentForm.Caption)
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "Everything past this point is preset values for the controls on this form."
    xmlEngine.writeBlankLine

End Function

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
    For Each eControl In parentForm.Controls
        
        controlName = eControl.Name
        If InControlArray(eControl) Then controlIndex = eControl.Index Else controlIndex = -1
        controlType = TypeName(eControl)
        controlValue = ""
            
        'We only want to write out the value property of relevant controls.  Check that list now.
        Select Case controlType
        
            'Our custom controls all have a .Value property
            Case "sliderTextCombo", "smartCheckBox", "smartOptionButton", "textUpDown"
                controlValue = CStr(eControl.Value)
            
            'Color pickers have a .Color property
            Case "colorSelector"
                controlValue = CStr(eControl.Color)
            
            'Intrinsic VB controls may have different names for their value properties
            Case "HScrollBar", "VScrollBar"
                controlValue = CStr(eControl.Value)
                
            Case "ListBox", "ComboBox"
            
                'Make sure the combo box is not the preset box on this control!
                controlValue = CStr(eControl.ListIndex)
                
            Case "TextBox"
                controlValue = CStr(eControl.Text)
        
        End Select
        
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
    
    Dim eControl As Object
    For Each eControl In parentForm.Controls
        
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
                            eControl.Value = CDbl(controlValue)
                                            
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
                    eControl.ListIndex = CLng(controlValue)
                    
                Case "TextBox"
                    eControl.Text = controlValue
            
            End Select

        End If
        
    Next eControl
    
    'Allow the user to retrieve any of their custom preset data from the file
    curPresetEntry = xmlSafePresetName
    RaiseEvent ReadCustomPresetData
    
    'We have now filled all controls with their relevant values from the XML file.
    readXMLSettings = True
    
End Function

'This beautiful little function comes courtesy of coder Merri:
' http://www.vbforums.com/showthread.php?536960-RESOLVED-how-can-i-see-if-the-object-is-array-or-not
Private Function InControlArray(Ctl As Object) As Boolean
    InControlArray = Not Ctl.Parent.Controls(Ctl.Name) Is Ctl
End Function
