VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdThemeColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Visual Theme Color List class
'Copyright 2016-2016 by Tanner Helland
'Created: 15/January/16
'Last updated: 15/January/16
'Last update: initial build
'
'Each individual PD control uses a unique list of colors.  Some of these colors may be modified by different settings
' or actions (e.g. enabled vs disabled, hovered vs active).
'
'This class exists to simplify the color retrieval process.  Each control maintains a list of needed colors as some
' kind of enum (meaning each color is assigned numerically).  During the initialization step, this class retrieves
' the string values matching those Enums from the master theme file, plugs in any missing values (e.g. not every
' color requires specialized "hovered" or "disabled" values), and fills a matching array with the complete list of
' required colors.
'
'This way, the rendering function inside each UC can access colors very quickly, regardless of changes to control state.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Number of colors stored by this list.  This must be specified by the caller during the initialization step.
Private m_NumOfColors As Long

'Color definition.  If a variation is missing, it will be replaced by the BaseColor value.  (As such, the BaseColor
' value must always be present in a color definition.)
Private Type pdThemeColor
    BaseColor As Long
    DisabledColor As Long
    ActiveColor As Long
    HoverColor As Long
    ActiveHoverColor As Long
End Type

'Actual color list.  This list is only modified during the initialization stage of the class, but it can be
' refreshed against a new theme if necessary.
Private m_ColorList() As pdThemeColor

'Name of the current object.  Because many controls share the same color names (e.g. "background", "text", etc),
' we differentiate between them by object name (e.g. "button-strip").
Private m_ObjectName As String

'Initialize a color list for a given object.  This preps a number of structures, but DOESN'T ACTUALLY RETRIEVE COLORS.
' (Colors must be retrieved individually.)
Public Function InitializeColorList(ByRef srcObjectName As String, ByVal numOfRequiredColors As Long) As Boolean
    
    'First, make sure the target object exists in the theme file.  If it doesn't, there's not much we can do.
    ' (Note that this step will fail inside the IDE; we have workarounds for this.)
    m_ObjectName = srcObjectName
    If g_IsProgramRunning Then
        InitializeColorList = g_Themer.VerifyThemeObject(srcObjectName)
    Else
        InitializeColorList = True
    End If
    
    'Prep the internal color list
    If numOfRequiredColors > 0 Then
        m_NumOfColors = numOfRequiredColors
        ReDim m_ColorList(0 To m_NumOfColors - 1) As pdThemeColor
    Else
        Debug.Print "WARNING!  You can't request zero colors from pdThemeColors.InitializeColorList()!"
    End If
    
End Function

'Load a given color from the active theme file.  Here's how this step works:
' - While inside the IDE:
'   - All values are ignored except the "IDE_FailsafeColor" value.
' - While running:
'   - We first look up the supplied color name inside the active theme file, using the current object's name as
'      the color namespace.
'       - If the color is found, we store it as the BaseColor value.
'       - If the color cannot be found, we fallback to the "Default" namespace and try again.
'   - Next, we search for the Disabled, Active, and Hover variants of the color.
'       - If a variant exists, it is stored in its respective position.
'       - If a variant does not exist, the BaseColor value is plugged in instead.
Public Function LoadThemeColor(ByVal colorIndex As Long, ByVal colorName As String, ByVal IDE_FailsafeColor As String) As Boolean
    
    LoadThemeColor = False
    
    'Before doing anything else, validate the color index
    If (colorIndex >= 0) And (colorIndex < m_NumOfColors) Then
    
        Dim colorLookupSuccessful As Boolean
        colorLookupSuccessful = False
        
        Dim baseColorValue As Long
        
        'If the program is running, try to retrieve a matching color from PD's central theme manager.
        If g_IsProgramRunning Then
        
            'Ask the central themer to look up this color.  It handles most the cumbersome work, like mapping color
            ' definition trees back to their source.
            Dim colorString As String
            colorString = g_Themer.LookUpColor(m_ObjectName, colorName)
            
            'If the color lookup was successful, colorString will be non-null.
            If Len(colorString) <> 0 Then
                
                'Attempt to convert the string into an RGB long
                colorLookupSuccessful = Colors.GetColorFromString(colorString, baseColorValue)
                If colorLookupSuccessful Then
                    
                    'Store the base color value.  We'll use this as the basis for all subsequent color calculations.
                    m_ColorList(colorIndex).BaseColor = baseColorValue
                    
                    'Because this color successfully exists, we are now going to check for disabled, active, hovered,
                    ' and active+hovered variants.  Each must be handled individually, and all are optional.  (If an
                    ' optional variant is missing, we'll just plug in the base color value, with the exception of
                    ' active+hovered which gets the active color by default.)
                    FillInColorVariant m_ObjectName, colorName, "disabled", m_ColorList(colorIndex).DisabledColor, baseColorValue
                    FillInColorVariant m_ObjectName, colorName, "active", m_ColorList(colorIndex).ActiveColor, baseColorValue
                    FillInColorVariant m_ObjectName, colorName, "hover", m_ColorList(colorIndex).HoverColor, baseColorValue
                    FillInColorVariant m_ObjectName, colorName, "activehover", m_ColorList(colorIndex).ActiveHoverColor, m_ColorList(colorIndex).ActiveColor
                    
                Else
                    pdDebug.LogAction "WARNING!  The themed color for " & m_ObjectName & ":" & colorName & " couldn't be converted into a valid RGB value."
                End If
                 
            Else
                pdDebug.LogAction "WARNING!  The themed color lookup for " & m_ObjectName & ":" & colorName & " failed."
            End If
        
        End If
        
        'If the program is not running, or if color lookup failed, fall back to the "IDE default" value
        If Not colorLookupSuccessful Then
            
            colorLookupSuccessful = Colors.GetColorFromString(IDE_FailsafeColor, baseColorValue)
            If colorLookupSuccessful Then
                With m_ColorList(colorIndex)
                    .BaseColor = baseColorValue
                    .DisabledColor = baseColorValue
                    .ActiveColor = baseColorValue
                    .HoverColor = baseColorValue
                End With
            Else
                Debug.Print "WARNING!  Your IDE failsafe color (" & IDE_FailsafeColor & ") is invalid.  Fix it!"
            End If
            
        End If
        
        'Return the success/fail value of the base color (since color variants are optional)
        LoadThemeColor = colorLookupSuccessful
        
    Else
        Debug.Print "WARNING!  You've requested an invalid color index from pdThemeColors.LoadThemeColor()!"
    End If

End Function

'Once a base color has been successfully retrieved, you can use this function to plug in any/all variant values
Private Sub FillInColorVariant(ByRef objectName As String, ByRef colorName As String, ByVal variantStringTag As String, ByRef dstColor As Long, ByVal failsafeColor As Long)
    
    'By default, assign the fallback color value.  If the rest of this function is successful, it will overwrite this
    ' value with a new, custom value.
    dstColor = failsafeColor
    
    Dim colorVariantName As String
    colorVariantName = colorName & "-" & variantStringTag
    
    Dim colorString As String
    colorString = g_Themer.LookUpColor(objectName, colorVariantName)
    
    If Len(colorString) <> 0 Then
        If Not Colors.GetColorFromString(colorString, dstColor) Then
            dstColor = failsafeColor
        End If
    End If
    
End Sub
