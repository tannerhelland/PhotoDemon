Attribute VB_Name = "Units"
'***************************************************************************
'Unit Conversion Functions
'Copyright 2014-2025 by Tanner Helland
'Created: 10/February/14
'Last updated: 04/March/24
'Last update: helper function to retrieve OS user preference for metric vs imperial units
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
    Private Const mu_Percent = 0, mu_Pixels = 1, mu_Inches = 2, mu_Centimeters = 3, mu_Millimeters = 4, mu_Points = 5, mu_Picas = 6, MU_MAX = 6
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

Public Function GetNumOfAvailableUnits() As Long
    GetNumOfAvailableUnits = MU_MAX
End Function

'Given a measurement in pixels, convert it to some other unit of measurement.  Note that at least two parameters are required:
' the unit of measurement to use, and a source measurement (in pixels, obviously).  Depending on the conversion, one of two
' optional parameters may also be necessary: a pixel resolution, expressed as PPI (needed for absolute measurements like inches
' or cm), and for percentage, an ORIGINAL value, in pixels, must be supplied.
Public Function GetValueFormattedForUnit_FromPixel(ByVal curUnit As PD_MeasurementUnit, ByVal srcPixelValue As Double, Optional ByVal srcPixelResolution As Double = 0#, Optional ByVal initPixelValue As Double = 0#, Optional ByVal appendUnitAsText As Boolean = False) As String
    
    If (curUnit <> mu_Pixels) Then srcPixelValue = Units.ConvertPixelToOtherUnit(curUnit, srcPixelValue, srcPixelResolution, initPixelValue)
    
    Select Case curUnit
    
        Case mu_Percent
            GetValueFormattedForUnit_FromPixel = Format$(srcPixelValue, "0.0#")
        
        Case mu_Pixels
            GetValueFormattedForUnit_FromPixel = CStr(Int(srcPixelValue + 0.5))
        
        Case mu_Inches
            GetValueFormattedForUnit_FromPixel = Format$(srcPixelValue, "0.0##")
        
        Case mu_Centimeters
            GetValueFormattedForUnit_FromPixel = Format$(srcPixelValue, "0.0#")
            
        Case mu_Millimeters
            GetValueFormattedForUnit_FromPixel = Format$(srcPixelValue, "0.0#")
            
        Case mu_Points
            GetValueFormattedForUnit_FromPixel = Format$(srcPixelValue, "0.0#")
        
        Case mu_Picas
            GetValueFormattedForUnit_FromPixel = Format$(srcPixelValue, "0.0#")
        
    End Select
    
    If appendUnitAsText Then
        If (curUnit = mu_Percent) Then
            Const PERCENT_SIGN As String = "%"
            GetValueFormattedForUnit_FromPixel = GetValueFormattedForUnit_FromPixel & PERCENT_SIGN
        Else
            GetValueFormattedForUnit_FromPixel = GetValueFormattedForUnit_FromPixel & " " & Units.GetNameOfUnit(curUnit, True)
        End If
    End If
        
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
