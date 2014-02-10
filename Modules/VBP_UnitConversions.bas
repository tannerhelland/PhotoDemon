Attribute VB_Name = "Unit_Conversion_Functions"
'***************************************************************************
'Unit Conversion Functions
'Copyright ©2013-2014 by Tanner Helland
'Created: 10/February/14
'Last updated: 10/February/14
'Last update: abstracted measurement conversion code from smartResize UC to this standalone module
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until February '14.  This module is now used to store all the random bits of unit conversion math required by the
' program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'Units of measurement, as used by PD (particularly the resize dialogs)
Public Enum MeasurementUnit
    MU_PERCENT = 0
    MU_PIXELS = 1
End Enum

#If False Then
    Const MU_PERCENT = 0, MU_PIXELS = 1
#End If

'Given a measurement in pixels, convert it to some other unit of measurement.  Note that at least two parameters are required:
' the unit of measurement to use, and a source measurement (in pixels, obviously).  Depending on the conversion, one of two
' optional parameters may also be necessary: a pixel resolution, expressed as PPI (needed for absolute measurements like inches
' or cm), and for percentage, an ORIGINAL value, in pixels, must be supplied.
Public Function convertPixelToOtherUnit(ByVal UnitOfMeasurement As MeasurementUnit, ByVal srcPixelValue As Double, Optional ByVal srcPixelResolution As Double, Optional ByVal initPixelValue As Double) As Double

    Select Case UnitOfMeasurement
    
        Case MU_PERCENT
            convertPixelToOtherUnit = (srcPixelValue / initPixelValue) * 100
        
        Case MU_PIXELS
            convertPixelToOtherUnit = srcPixelValue
    
    End Select

End Function

'Given a measurement in something other than pixels, convert it to pixels.  Note that at least two parameters are required:
' the unit of measurement that defines the source value, and the source value itself.  Depending on the conversion, one of two
' optional parameters may also be necessary: a resolution, expressed as PPI (needed to convert from absolute measurements like
' inches or cm), and for percentage, an ORIGINAL value, in pixels, must be supplied.  Note that in the unique case of percent,
' the "srcUnitValue" will be the percent used for conversion (as a percent, e.g. 100.0 for 100%).
Public Function convertOtherUnitToPixels(ByVal UnitOfMeasurement As MeasurementUnit, ByVal srcUnitValue As Double, Optional ByVal srcUnitResolution As Double, Optional ByVal initPixelValue As Double) As Double

    'The translation function used depends on the currently selected unit
    Select Case UnitOfMeasurement
    
        'Percent
        Case MU_PERCENT
            convertOtherUnitToPixels = CDbl(srcUnitValue / 100) * initPixelValue
            
        'Pixels
        Case MU_PIXELS
            convertOtherUnitToPixels = srcUnitValue
            
    End Select
    
End Function
