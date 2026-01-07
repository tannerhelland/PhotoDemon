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
'Copyright 2016-2026 by Tanner Helland
'Created: 15/January/16
'Last updated: 21/January/16
'Last update: continue fleshing out features
'
'Each individual PD control uses a unique list of colors.  Some of these colors may be modified
' by different settings or actions (e.g. enabled vs disabled, hovered vs active).
'
'This class exists to simplify color retrieval.  Each control maintains a list of required colors
' as some kind of enum (meaning each color is assigned numerically).  During initialization,
' this class retrieves the string values matching those Enums from PD's central theme file,
' plugs in any missing values (e.g. not every color requires specialized "hovered" or "disabled" values),
' then fills a matching array with the complete list of colors.
'
'This way, the rendering function inside each UC can access colors very quickly, regardless of changes
' to control state.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Number of colors stored by this list.  This must be specified by the caller during the initialization step.
Private m_NumOfColors As Long

'Actual color list.  This list is only modified during the initialization stage of the class, but it can be
' refreshed against a new theme if necessary.
Private m_ColorList() As PDThemeColor

'Name of the current object.  Because many controls share the same color names (e.g. "background", "text", etc),
' we differentiate between them by object name (e.g. "button-strip").
Private m_ObjectName As String

'Initialize a color list for a given object.  This preps a number of structures, but DOESN'T ACTUALLY RETRIEVE COLORS.
' (Colors must be retrieved individually.)
Friend Function InitializeColorList(ByRef srcObjectName As String, ByVal numOfRequiredColors As Long) As Boolean
    
    'First, make sure the target object exists in the theme file.  If it doesn't, there's not much we can do.
    ' (Note that this step will fail inside the IDE; a workaround is provided.)
    m_ObjectName = srcObjectName
    If PDMain.IsProgramRunning() And (Not g_Themer Is Nothing) Then
        InitializeColorList = g_Themer.VerifyThemeObject(srcObjectName)
    Else
        InitializeColorList = True
    End If
    
    'Prep the internal color list
    If (numOfRequiredColors > 0) Then
        m_NumOfColors = numOfRequiredColors
        ReDim m_ColorList(0 To m_NumOfColors - 1) As PDThemeColor
    Else
        InternalProblem "WARNING!  You can't request zero colors from pdThemeColors.InitializeColorList()!"
    End If
    
End Function

'Load a given color from the active theme file.  Here's how this step works:
' - While inside the designer:
'   - All values are ignored except the "IDE_FailsafeColor" value.
' - While running:
'   - We first look up the supplied color name inside the active theme file, using the current object's name as
'      the color namespace.
'       - If the color is found, we store it as the BaseColor value.
'       - If the color cannot be found, we fallback to the "Default" namespace and try again.
'   - Next, we search for the Disabled, Active, and Hover variants of the color.
'       - If a variant exists, it is stored in its respective position.
'       - If a variant does not exist, the BaseColor value is plugged in instead.
Friend Function LoadThemeColor(ByVal colorIndex As Long, ByRef srcColorName As String, ByRef IDE_FailsafeColor As String) As Boolean
    
    LoadThemeColor = False
    
    'Before doing anything else, validate the color index
    If (colorIndex >= 0) And (colorIndex < m_NumOfColors) Then
    
        Dim colorLookupSuccessful As Boolean
        colorLookupSuccessful = False
        
        Dim baseColorValue As Long
        
        'If the program is running, try to retrieve a matching color from PD's central theme manager.
        If PDMain.IsProgramRunning() And (Not g_Themer Is Nothing) Then
            
            'Start by seeing if the central themer has already cached a value for this color
            If g_Themer.RetrieveColorFromCache(m_ObjectName, srcColorName, m_ColorList(colorIndex)) Then
                colorLookupSuccessful = True
                
            'If this color does not exist in the central cache, we need to parse the original XML file for it
            Else
                
                'Ask the central themer to look up this color.  It handles most the cumbersome work, like mapping color
                ' definition trees back to their source.
                Dim colorString As String
                colorString = g_Themer.LookUpColor(m_ObjectName, srcColorName)
                
                'If the color lookup was successful, colorString will be non-null.
                If (LenB(colorString) <> 0) Then
                    
                    'Attempt to convert the string into an RGB long
                    colorLookupSuccessful = Colors.GetColorFromString(colorString, baseColorValue, ColorHex)
                    If colorLookupSuccessful Then
                        
                        'Store the base color value.  We'll use this as the basis for all subsequent color calculations.
                        m_ColorList(colorIndex).baseColor = baseColorValue
                        
                        'Because this color successfully exists, we are now going to check for disabled, active, hovered,
                        ' and active+hovered variants.  Each must be handled individually, and all are optional.  (If an
                        ' optional variant is missing, we'll just plug in the base color value, with the exception of
                        ' active+hovered which gets the *active* color by default.)
                        
                        'Note that this step was previously split into a separate function, but for performance reasons,
                        ' it's been moved in-line.
                        With m_ColorList(colorIndex)
                        
                            colorString = g_Themer.LookUpColor(m_ObjectName, srcColorName & "-disabled")
                            If (LenB(colorString) = 0) Then
                                .disabledColor = baseColorValue
                            Else
                                If (Not Colors.GetColorFromString(colorString, .disabledColor, ColorHex)) Then .disabledColor = baseColorValue
                            End If
                            
                            colorString = g_Themer.LookUpColor(m_ObjectName, srcColorName & "-active")
                            If (LenB(colorString) = 0) Then
                                .ActiveColor = baseColorValue
                            Else
                                If (Not Colors.GetColorFromString(colorString, .ActiveColor, ColorHex)) Then .ActiveColor = baseColorValue
                            End If
                            
                            colorString = g_Themer.LookUpColor(m_ObjectName, srcColorName & "-hover")
                            If (LenB(colorString) = 0) Then
                                .hoverColor = baseColorValue
                            Else
                                If (Not Colors.GetColorFromString(colorString, .hoverColor, ColorHex)) Then .hoverColor = baseColorValue
                            End If
                            
                            colorString = g_Themer.LookUpColor(m_ObjectName, srcColorName & "-activehover")
                            If (LenB(colorString) = 0) Then
                                .ActiveHoverColor = .ActiveColor
                            Else
                                If (Not Colors.GetColorFromString(colorString, .ActiveHoverColor, ColorHex)) Then .ActiveHoverColor = .ActiveColor
                            End If
                            
                        End With
                        
                        'With the color table successfully populated, cache its values in the central themer.  This spares
                        ' future requests for this color from having to repeat these steps from scratch.
                        g_Themer.AddColorToCache m_ObjectName, srcColorName, m_ColorList(colorIndex)
                        
                    Else
                        InternalProblem "WARNING!  Theme color for " & m_ObjectName & ":" & srcColorName & " couldn't be converted to a valid RGB value."
                    End If
                     
                Else
                    InternalProblem "WARNING!  Theme color lookup for " & m_ObjectName & ":" & srcColorName & " failed."
                End If
            
            End If
            
        End If
        
        'If the program is not running, or if color lookup failed, fall back to the "IDE default" value
        If (Not colorLookupSuccessful) Then
            
            colorLookupSuccessful = Colors.GetColorFromString(IDE_FailsafeColor, baseColorValue)
            If colorLookupSuccessful Then
                With m_ColorList(colorIndex)
                    .baseColor = baseColorValue
                    .disabledColor = baseColorValue
                    .ActiveColor = baseColorValue
                    .hoverColor = baseColorValue
                End With
            Else
                InternalProblem "WARNING!  Your IDE failsafe color (" & IDE_FailsafeColor & ") is invalid.  Fix it!"
            End If
            
        End If
        
        'Return the success/fail value of the base color (since color variants are optional)
        LoadThemeColor = colorLookupSuccessful
        
    Else
        InternalProblem "WARNING!  You've requested an invalid color index from pdThemeColors.LoadThemeColor()!"
    End If
    
End Function

'Once all colors have been populated, this function is used to return actual color values.  It's all lookup-table based,
' so feel free to use it inside actual rendering functions.
Friend Function RetrieveColor(ByVal colorIndex As Long, Optional ByVal enabledState As Boolean = True, Optional ByVal activeState As Boolean = False, Optional ByVal hoverState As Boolean = False) As Long

    'Before doing anything else, validate the color index
    If (colorIndex >= 0) And (colorIndex < m_NumOfColors) Then
    
        'Branch according to the passed control state.  The combination of values determines which color we return.
        If enabledState Then
            If activeState Then
                If hoverState Then
                    RetrieveColor = m_ColorList(colorIndex).ActiveHoverColor
                Else
                    RetrieveColor = m_ColorList(colorIndex).ActiveColor
                End If
            Else
                If hoverState Then
                    RetrieveColor = m_ColorList(colorIndex).hoverColor
                Else
                    RetrieveColor = m_ColorList(colorIndex).baseColor
                End If
            End If
        Else
            RetrieveColor = m_ColorList(colorIndex).disabledColor
        End If
    
    Else
        InternalProblem "WARNING!  You've requested an invalid color index from pdThemeColors.RetrieveColor()!"
    End If

End Function

Private Sub InternalProblem(ByRef errMsg As String)
    PDDebug.LogAction "WARNING!  pdThemeColors returned an error: " & errMsg
End Sub
