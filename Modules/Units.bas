Attribute VB_Name = "Units"
'***************************************************************************
'Unit Handling and Conversion Functions
'Copyright 2014-2026 by Tanner Helland
'Created: 10/February/14
'Last updated: 20/March/25
'Last update: load size presets from file on-demand and cache them here
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until February '14.  This module is now used to store all the random bits of unit conversion math required by the
' program.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Units of measurement, as used by PD (particularly the resize dialogs)
Public Enum PD_MeasurementUnit
    'Used only during validation steps; do *not* pass to any measurement functions or they will fail
    mu_Unknown = -1
    
    mu_Percent = 0
    mu_Pixels = 1
    mu_Inches = 2
    mu_Centimeters = 3
    mu_Millimeters = 4
    mu_Points = 5
    mu_Picas = 6
    [MU_MAX] = 6
End Enum

#If False Then
    Private Const mu_Unknown = -1, mu_Percent = 0, mu_Pixels = 1, mu_Inches = 2, mu_Centimeters = 3, mu_Millimeters = 4, mu_Points = 5, mu_Picas = 6, MU_MAX = 6
#End If

Public Enum PD_ResolutionUnit
    ru_PPI = 0
    ru_PPCM = 1
End Enum

#If False Then
    Private Const ru_PPI = 0, ru_PPCM = 1
#End If

'Used to query the OS to determine if metric or imperial units should be the default
Private Const LOCALE_USER_DEFAULT As Long = &H400&
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long

'PD loads the size template from file once, when a template UI list is first touched.
' After that first load, we store it persistently here, and other areas can access it on-demand.
Private m_TemplatesInitialized As Boolean, m_TemplateSizeFile As String, m_TemplateAspectFile As String

Public Type PD_SizeTemplate
    hSizeOrig As Single
    vSizeOrig As Single
    thisSizeUnit As PD_MeasurementUnit
    localizedName As String
End Type

'This module maintains a list of both aspect ratio *and* size presets
Private m_TemplateSize() As PD_SizeTemplate, m_TemplateAspect() As PD_SizeTemplate
Private m_numTemplateSize As Long, m_numTemplateAspect As Long

'Given a measurement in pixels, convert it to some other unit of measurement.  Note that at least two parameters are required:
' the unit of measurement to use, and a source measurement (in pixels, obviously).  Depending on the conversion, one of two
' optional parameters may also be necessary: a pixel resolution, expressed as PPI (needed for absolute measurements like inches
' or cm), and for percentage, an ORIGINAL value, in pixels, must be supplied.
Public Function ConvertPixelToOtherUnit(ByVal curUnit As PD_MeasurementUnit, ByVal srcPixelValue As Double, Optional ByVal srcPixelResolution As Double, Optional ByVal initPixelValue As Double) As Double

    Select Case curUnit
    
        Case mu_Percent
            If (initPixelValue <> 0) Then ConvertPixelToOtherUnit = (srcPixelValue / initPixelValue) * 100#
            
        Case mu_Pixels
            ConvertPixelToOtherUnit = srcPixelValue
            
        Case mu_Inches
            If (srcPixelResolution <> 0#) Then ConvertPixelToOtherUnit = srcPixelValue / srcPixelResolution
        
        Case mu_Centimeters
            If (srcPixelResolution <> 0#) Then ConvertPixelToOtherUnit = GetCMFromInches(srcPixelValue / srcPixelResolution)
            
        Case mu_Millimeters
            If (srcPixelResolution <> 0#) Then ConvertPixelToOtherUnit = GetCMFromInches(srcPixelValue / srcPixelResolution) * 10#
            
        Case mu_Points
            If (srcPixelResolution <> 0#) Then ConvertPixelToOtherUnit = (srcPixelValue / srcPixelResolution) * 72#
        
        Case mu_Picas
            If (srcPixelResolution <> 0#) Then ConvertPixelToOtherUnit = (srcPixelValue / srcPixelResolution) * 6#
    
    End Select

End Function

'Given a measurement in something other than pixels, convert it to pixels.  Note that at least two parameters are required:
' the unit of measurement that defines the source value, and the source value itself.  Depending on the conversion, one of two
' optional parameters may also be necessary: a resolution, expressed as PPI (needed to convert from absolute measurements like
' inches or cm), and for percentage, an ORIGINAL value, in pixels, must be supplied.  Note that in the unique case of percent,
' the "srcUnitValue" will be the percent used for conversion (as a percent, e.g. 100.0 for 100%).
Public Function ConvertOtherUnitToPixels(ByVal curUnit As PD_MeasurementUnit, ByVal srcUnitValue As Double, Optional ByVal srcUnitResolution As Double, Optional ByVal initPixelValue As Double) As Double

    'The translation function used depends on the currently selected unit
    Select Case curUnit
    
        Case mu_Percent
            ConvertOtherUnitToPixels = CDbl(srcUnitValue / 100#) * initPixelValue
        
        Case mu_Pixels
            ConvertOtherUnitToPixels = srcUnitValue
        
        Case mu_Inches
            ConvertOtherUnitToPixels = Int(srcUnitValue * srcUnitResolution + 0.5)
        
        Case mu_Centimeters
            ConvertOtherUnitToPixels = Int(GetInchesFromCM(srcUnitValue) * srcUnitResolution + 0.5)
            
        Case mu_Millimeters
            ConvertOtherUnitToPixels = Int(GetInchesFromCM(srcUnitValue / 10#) * srcUnitResolution + 0.5)
            
        Case mu_Points
            ConvertOtherUnitToPixels = Int((srcUnitValue / 72#) * srcUnitResolution + 0.5)
        
        Case mu_Picas
            ConvertOtherUnitToPixels = Int((srcUnitValue / 6#) * srcUnitResolution + 0.5)
        
    End Select
    
End Function

'Basic metric/imperial conversions for length
Public Function GetInchesFromCM(ByVal srcCM As Double) As Double
    GetInchesFromCM = srcCM * 0.393700787
End Function

Public Function GetCMFromInches(ByVal srcInches As Double) As Double
    GetCMFromInches = srcInches * 2.54
End Function

'Retrieve localized names for various measurement units
Public Function GetNameOfUnit(ByVal srcUnit As PD_MeasurementUnit, Optional ByVal getAbbreviatedForm As Boolean = False) As String

    If getAbbreviatedForm Then
    
        Select Case srcUnit
        
            Case mu_Percent
                GetNameOfUnit = "%"
            
            Case mu_Pixels
                GetNameOfUnit = g_Language.TranslateMessage("px")
            
            Case mu_Inches
                GetNameOfUnit = g_Language.TranslateMessage("in")
            
            Case mu_Centimeters
                GetNameOfUnit = g_Language.TranslateMessage("cm")
                
            Case mu_Millimeters
                GetNameOfUnit = g_Language.TranslateMessage("mm")
                
            Case mu_Points
                GetNameOfUnit = g_Language.TranslateMessage("pt")
            
            Case mu_Picas
                GetNameOfUnit = g_Language.TranslateMessage("pc")
            
        End Select
    
    Else

        Select Case srcUnit
        
            Case mu_Percent
                GetNameOfUnit = g_Language.TranslateMessage("percent")
            
            Case mu_Pixels
                GetNameOfUnit = g_Language.TranslateMessage("pixels")
            
            Case mu_Inches
                GetNameOfUnit = g_Language.TranslateMessage("inches")
            
            Case mu_Centimeters
                GetNameOfUnit = g_Language.TranslateMessage("centimeters")
                
            Case mu_Millimeters
                GetNameOfUnit = g_Language.TranslateMessage("millimeters")
                
            Case mu_Points
                GetNameOfUnit = g_Language.TranslateMessage("points")
            
            Case mu_Picas
                GetNameOfUnit = g_Language.TranslateMessage("picas")
            
        End Select
        
    End If

End Function

'Retrieve localized names for various measurement units
Public Function GetUnitFromName(ByVal srcName As String) As PD_MeasurementUnit

    Select Case srcName
        Case "%", "percent", g_Language.TranslateMessage("percent")
            GetUnitFromName = mu_Percent
        Case "px", "pixels", g_Language.TranslateMessage("px"), g_Language.TranslateMessage("pixels")
            GetUnitFromName = mu_Pixels
        Case "in", "inches", g_Language.TranslateMessage("in"), g_Language.TranslateMessage("inches")
            GetUnitFromName = mu_Inches
        Case "cm", "centimeters", g_Language.TranslateMessage("cm"), g_Language.TranslateMessage("centimeters")
            GetUnitFromName = mu_Centimeters
        Case "mm", "millimeters", g_Language.TranslateMessage("mm"), g_Language.TranslateMessage("millimeters")
            GetUnitFromName = mu_Millimeters
        Case "pt", "points", g_Language.TranslateMessage("pt"), g_Language.TranslateMessage("points")
            GetUnitFromName = mu_Points
        Case "pc", "picas", g_Language.TranslateMessage("pc"), g_Language.TranslateMessage("picas")
            GetUnitFromName = mu_Picas
        Case Else
            GetUnitFromName = mu_Unknown
    End Select
    
End Function

'Does *not* include validation enums (like mu_Unknown), by design
Public Function GetNumOfAvailableUnits() As Long
    GetNumOfAvailableUnits = MU_MAX
End Function

'Given a measurement, convert it to a display-friendly string with rounding and formatting consistently applied.
' (Note: the optional parameter "useRounding" only applies to PIXELS, because the incoming value is a float.)
Public Function GetValueFormattedForUnit(ByVal curUnit As PD_MeasurementUnit, ByVal srcValue As Double, Optional ByVal appendUnitAsText As Boolean = False, Optional ByVal useRounding As Boolean = True) As String
    
    Select Case curUnit
    
        Case mu_Percent
            GetValueFormattedForUnit = Format$(srcValue, "0.0#")
        
        Case mu_Pixels
            If useRounding Then srcValue = srcValue + 0.5
            GetValueFormattedForUnit = CStr(Int(srcValue))
        
        Case mu_Inches
            GetValueFormattedForUnit = Format$(srcValue, "0.0##")
        
        Case mu_Centimeters
            GetValueFormattedForUnit = Format$(srcValue, "0.0#")
            
        Case mu_Millimeters
            GetValueFormattedForUnit = Format$(srcValue, "0.0#")
            
        Case mu_Points
            GetValueFormattedForUnit = Format$(srcValue, "0.0#")
        
        Case mu_Picas
            GetValueFormattedForUnit = Format$(srcValue, "0.0#")
        
    End Select
    
    If appendUnitAsText Then
        If (curUnit = mu_Percent) Then
            Const PERCENT_SIGN As String = "%"
            GetValueFormattedForUnit = GetValueFormattedForUnit & PERCENT_SIGN
        Else
            GetValueFormattedForUnit = GetValueFormattedForUnit & " " & Units.GetNameOfUnit(curUnit, True)
        End If
    End If
        
End Function

'Given a measurement in pixels, convert it to some other unit of measurement.  Note that at least two parameters are required:
' the unit of measurement to use, and a source measurement (in pixels, obviously).  Depending on the conversion, one of two
' optional parameters may also be necessary: a pixel resolution, expressed as PPI (needed for absolute measurements like inches
' or cm), and for percentage, an ORIGINAL value, in pixels, must be supplied.
'
'(Note: the optional parameter "useRounding" only applies when converting some other unit to PIXELS.
Public Function GetValueFormattedForUnit_FromPixel(ByVal curUnit As PD_MeasurementUnit, ByVal srcPixelValue As Double, Optional ByVal srcPixelResolution As Double = 0#, Optional ByVal initPixelValue As Double = 0#, Optional ByVal appendUnitAsText As Boolean = False, Optional ByVal useRounding As Boolean = True) As String
    If (curUnit <> mu_Pixels) Then srcPixelValue = Units.ConvertPixelToOtherUnit(curUnit, srcPixelValue, srcPixelResolution, initPixelValue)
    GetValueFormattedForUnit_FromPixel = GetValueFormattedForUnit(curUnit, srcPixelValue, appendUnitAsText, useRounding)
End Function

'Returns TRUE if the current user's locale settings prefer METRIC (not imperial)
Public Function LocaleUsesMetric() As Boolean
    
    'Because GetLocaleInfo only returns a 1 or a 0 for metric vs non-metric, we don't need a large buffer.
    Dim sBuffer As String, sRet As String
    sBuffer = String$(4, 0)
    
    'From MSDN:
    ' System of measurement. The maximum number of characters allowed for this string is two, including a terminating null character.
    ' This value is 0 if the metric system (Systéme International d'Units, or S.I.) is used, and 1 if the United States system is used.
    Const LOCALE_IMEASURE As Long = &HD&
    sRet = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IMEASURE, StrPtr(sBuffer), Len(sBuffer))
    
    If (sRet > 0) Then
        LocaleUsesMetric = (CLng(Left$(sBuffer, sRet - 1)) = 0)
    Else
        LocaleUsesMetric = False
    End If
    
End Function

'If templates haven't been loaded yet, call this function - it will load them for you!
Public Sub InitializeSizeTemplates()
    
    'Look for the source file; if it doesn't exist, create it.
    ' Once created, load it.
    If (Not LoadTemplateFiles()) Then Exit Sub
    
    'We don't want to do this again
    m_TemplatesInitialized = True
    
End Sub

'Get a copy of the user's size and aspect ratio presets.
' Arrays are guaranteed precisely dimensioned to the number of available presets.
'
'Returns TRUE if at least one size and one aspect ratio preset exists.
Public Function GetCopyOfSizeAndAspectTemplates(ByRef dstSizes() As PD_SizeTemplate, ByRef dstAspects() As PD_SizeTemplate) As Boolean
        
    'Initialize the collection as needed
    If (Not m_TemplatesInitialized) Then InitializeSizeTemplates
        
    If (m_numTemplateSize > 0) Then
        ReDim dstSizes(0 To m_numTemplateSize - 1) As PD_SizeTemplate
    Else
        Erase dstSizes
    End If
    
    Dim i As Long
    For i = 0 To m_numTemplateSize - 1
        dstSizes(i) = m_TemplateSize(i)
    Next i
    
    If (m_numTemplateAspect > 0) Then
        ReDim dstAspects(0 To m_numTemplateAspect - 1) As PD_SizeTemplate
    Else
        Erase dstAspects
    End If
    
    For i = 0 To m_numTemplateAspect - 1
        dstAspects(i) = m_TemplateAspect(i)
    Next i
    
    GetCopyOfSizeAndAspectTemplates = (m_numTemplateSize > 0) And (m_numTemplateAspect > 0)
    
End Function

Private Function LoadTemplateFiles() As Boolean
    
    'Only load templates once per session (when first accessed)
    LoadTemplateFiles = m_TemplatesInitialized
    If LoadTemplateFiles Then Exit Function
    
    'Load the text from file into a pdStringStack object, or initialize the relevant text directly into the object
    Dim cSizePresets As pdStringStack, cAspectPresets As pdStringStack
    MakeTemplateFiles cSizePresets, cAspectPresets
    
    'We now have a stack of strings for both size presets and aspect ratio presets.
    ' Now we need to parse them into usable, strongly typed struct members.
    Const INIT_TEMPLATE_COUNT As Long = 16
    ReDim m_TemplateSize(0 To INIT_TEMPLATE_COUNT - 1) As PD_SizeTemplate
    ReDim m_TemplateAspect(0 To INIT_TEMPLATE_COUNT - 1) As PD_SizeTemplate
    m_numTemplateSize = 0
    m_numTemplateAspect = 0
    
    'Many UI elements display blank or special first entries in their template list.
    ' (This UI element will be selected when the user makes a custom or freeform selection.)
    ' This module doesn't apply those special entries.  It *only* loads defined templates from file.
    
    'Start with size templates
    Dim i As Long, tmpLine As String
    For i = 0 To cSizePresets.GetNumOfStrings() - 1
        
        tmpLine = Trim$(cSizePresets.GetString(i))
        
        Dim splitEntries() As String
        If (LenB(tmpLine) > 1) Then
            
            'Ignore comments (VB6 doesn't have a "continue for" statement)
            Const POUND_CHAR As String = "#"
            If (Left$(tmpLine, 1) = POUND_CHAR) Then GoTo NextEntry
            
            'Only one comma is required in a size line (e.g. "1920, 1080").
            ' There may also be a localized name and non-pixel measurement unit.
            Const COMMA_CHAR As String = ","
            splitEntries = Split(tmpLine, COMMA_CHAR, Compare:=vbBinaryCompare)
            If (UBound(splitEntries) >= 1) Then
                
                'Ensure we have room to store this entry
                If (UBound(m_TemplateSize) < m_numTemplateSize) Then ReDim Preserve m_TemplateSize(0 To m_numTemplateSize * 2 - 1) As PD_SizeTemplate
                
                Dim numSuccesses As Long, testFloat As Single
                numSuccesses = 0
                
                'Load dimensions and test each entry for correctness
                tmpLine = Trim$(splitEntries(0))
                If TextSupport.IsNumberLocaleUnaware(tmpLine) Then
                    testFloat = TextSupport.CDblCustom(tmpLine)
                    If (testFloat > 0!) Then
                        m_TemplateSize(m_numTemplateSize).hSizeOrig = testFloat
                        numSuccesses = numSuccesses + 1
                    End If
                End If
                
                tmpLine = Trim$(splitEntries(1))
                If TextSupport.IsNumberLocaleUnaware(tmpLine) Then
                    testFloat = TextSupport.CDblCustom(tmpLine)
                    If (testFloat > 0!) Then
                        m_TemplateSize(m_numTemplateSize).vSizeOrig = testFloat
                        numSuccesses = numSuccesses + 1
                    End If
                End If
                
                'Load any remaining entries if they exist
                If (numSuccesses = 2) Then
                    
                    m_TemplateSize(m_numTemplateSize).thisSizeUnit = mu_Pixels
                    If (UBound(splitEntries) >= 2) Then
                        m_TemplateSize(m_numTemplateSize).localizedName = Trim$(splitEntries(2))
                    Else
                        m_TemplateSize(m_numTemplateSize).localizedName = vbNullString
                    End If
                    
                    'Attempt to match unit text to name, if it exists
                    If (UBound(splitEntries) >= 3) Then
                        tmpLine = Trim$(splitEntries(3))
                        m_TemplateSize(m_numTemplateSize).thisSizeUnit = Units.GetUnitFromName(tmpLine)
                        If (m_TemplateSize(m_numTemplateSize).thisSizeUnit = mu_Unknown) Then
                            m_TemplateSize(m_numTemplateSize).thisSizeUnit = mu_Pixels
                            numSuccesses = 0
                        End If
                    End If
                    
                End If
                
                'We only require valid width/height to succeed (and store this template permanently)
                If (numSuccesses >= 2) Then m_numTemplateSize = m_numTemplateSize + 1
            
            '/end "line has at least one comma"
            End If
            
        '/end "line is non-empty"
        End If
        
NextEntry:
    Next i
    
    'Trim the final array size
    If (m_numTemplateSize > 0) Then ReDim Preserve m_TemplateSize(0 To m_numTemplateSize - 1) As PD_SizeTemplate
    
    'Repeat all the above steps, but for aspect ratio.  (These are easier because they are nameless and unit-less.)
    For i = 0 To cAspectPresets.GetNumOfStrings() - 1
        
        tmpLine = Trim$(cAspectPresets.GetString(i))
        If (LenB(tmpLine) > 1) Then
            
            'Ignore comments (VB6 doesn't have a "continue for" statement)
            If (Left$(tmpLine, 1) = POUND_CHAR) Then GoTo NextAspectEntry
            
            'Only one comma is required in a size line (e.g. "1920, 1080").
            ' There may also be a localized name and non-pixel measurement unit.
            splitEntries = Split(tmpLine, COMMA_CHAR, Compare:=vbBinaryCompare)
            If (UBound(splitEntries) >= 1) Then
                
                'Ensure we have room to store this entry
                If (UBound(m_TemplateAspect) < m_numTemplateAspect) Then ReDim Preserve m_TemplateAspect(0 To m_numTemplateAspect * 2 - 1) As PD_SizeTemplate
                numSuccesses = 0
                
                'Load dimensions and test each entry for correctness
                tmpLine = Trim$(splitEntries(0))
                If TextSupport.IsNumberLocaleUnaware(tmpLine) Then
                    testFloat = TextSupport.CDblCustom(tmpLine)
                    If (testFloat > 0!) Then
                        m_TemplateAspect(m_numTemplateAspect).hSizeOrig = testFloat
                        numSuccesses = numSuccesses + 1
                    End If
                End If
                
                tmpLine = Trim$(splitEntries(1))
                If TextSupport.IsNumberLocaleUnaware(tmpLine) Then
                    testFloat = TextSupport.CDblCustom(tmpLine)
                    If (testFloat > 0!) Then
                        m_TemplateAspect(m_numTemplateAspect).vSizeOrig = testFloat
                        numSuccesses = numSuccesses + 1
                    End If
                End If
                
                'Ignore anything left on the line
                
                'Aspect ratio only requires valid width/height to succeed (and store this template permanently)
                If (numSuccesses >= 2) Then m_numTemplateAspect = m_numTemplateAspect + 1
            
            '/end "line has at least one comma"
            End If
            
        '/end "line is non-empty"
        End If
        
NextAspectEntry:
    Next i
    
    'Trim the final array size
    If (m_numTemplateAspect > 0) Then ReDim Preserve m_TemplateAspect(0 To m_numTemplateAspect - 1) As PD_SizeTemplate
    
    'Only return success if we were able to generate some templates!
    LoadTemplateFiles = (m_numTemplateSize > 0) And (m_numTemplateAspect > 0)
    
End Function

Private Sub MakeTemplateFiles(ByRef dstSizes As pdStringStack, ByRef dstAspects As pdStringStack)

    'Try to load the file, if one exists
    Dim fileOK As Boolean: fileOK = False
    
    m_TemplateSizeFile = UserPrefs.GetPresetPath() & "Template_Sizes.txt"
    If Files.FileExists(m_TemplateSizeFile) Then
        
        Dim tmpString As String
        If Files.FileLoadAsString(m_TemplateSizeFile, tmpString, True) Then
            Set dstSizes = New pdStringStack
            dstSizes.CreateFromMultilineString tmpString, vbCrLf
            fileOK = (dstSizes.GetNumOfStrings() > 0)
        End If
    
    End If
    
    Dim tmpList As pdString, finalString As String
    
    'Sizes template file doesn't exist; create it anew
    If (Not fileOK) Then
        
        PDDebug.LogAction "Size template file not found; generating default one now..."
        
        'Manually populate a list of sizes
        Set tmpList = New pdString
        tmpList.AppendLine "1024, 768, XGA"
        tmpList.AppendLine "1280, 720, HD 720p"
        tmpList.AppendLine "1280, 768, WXGA"
        tmpList.AppendLine "1366, 768, FWXGA"
        tmpList.AppendLine "1600, 1200, UXGA"
        tmpList.AppendLine "1680, 1050, WSXGA+"
        tmpList.AppendLine "1920, 1080, Full HD 1080p"
        tmpList.AppendLine "1920, 1200, WUXGA"
        tmpList.AppendLine "2048, 1536, QXGA"
        tmpList.AppendLine "2560, 1600, WQXGA"
        tmpList.AppendLine "3840, 2160, 4K UHD"
        tmpList.AppendLine "3840, 2400, WQUXGA"
        tmpList.AppendLine "7680, 4320, 8K UHD"
        tmpList.AppendLine "16.5, 23.4, A2, in"
        tmpList.AppendLine "11.7, 16.5, A3, in"
        tmpList.AppendLine "8.3, 11.7, A4, in"
        tmpList.AppendLine "5.8, 8.3, A5, in"
        tmpList.AppendLine "4.1, 5.8, A6, in"
        tmpList.AppendLine "2.9, 4.1, A7, in"
        tmpList.AppendLine "8.5, 11, US Letter, in"
        tmpList.AppendLine "8.5, 14, US Legal, in"
        
        'Write it out to file
        finalString = tmpList.ToString()
        Files.FileDeleteIfExists m_TemplateSizeFile
        Files.FileSaveAsText finalString, m_TemplateSizeFile, True, True
        
        'Regardless of what happened with the file, create a matching list of lines from the
        ' stack we created.
        Set dstSizes = New pdStringStack
        dstSizes.CreateFromMultilineString finalString, vbCrLf
        
    End If
    
    'Repeat all the above steps, but with aspect ratios
    fileOK = False
    
    m_TemplateAspectFile = UserPrefs.GetPresetPath() & "Template_Aspects.txt"
    If Files.FileExists(m_TemplateAspectFile) Then
        
        If Files.FileLoadAsString(m_TemplateAspectFile, tmpString, True) Then
            Set dstAspects = New pdStringStack
            dstAspects.CreateFromMultilineString tmpString, vbCrLf
            fileOK = (dstAspects.GetNumOfStrings() > 0)
        End If
    
    End If
    
    'Aspect ratios template file doesn't exist; create it anew
    If (Not fileOK) Then
        
        'Manually populate a list of aspect ratios
        Set tmpList = New pdString
        tmpList.AppendLine "1, 1"
        tmpList.AppendLine "2, 3"
        tmpList.AppendLine "3, 5"
        tmpList.AppendLine "4, 6"
        tmpList.AppendLine "5, 7"
        tmpList.AppendLine "8, 10"
        tmpList.AppendLine "16, 9"
        tmpList.AppendLine "16, 10"
        tmpList.AppendLine "21, 9"
        
        'Write it out to file
        finalString = tmpList.ToString()
        Files.FileDeleteIfExists m_TemplateAspectFile
        Files.FileSaveAsText finalString, m_TemplateAspectFile, True, True
        
        'Regardless of what happened with the file, create a matching list of lines from the
        ' stack we created.
        Set dstAspects = New pdStringStack
        dstAspects.CreateFromMultilineString finalString, vbCrLf
        
    End If
    
End Sub
