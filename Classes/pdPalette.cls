VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Palette Container
'Copyright 2018-2018 by Tanner Helland
'Created: 16/January/18
'Last updated: 18/January/18
'Last update: forcibly remove duplicate entries from imported palettes
'
'This class currently has some non-obvious overlap with the Palettes module.  In the future, it would be nice to
' compartmentalize more palette-specific functionality within this class, then pass around class instances instead
' of bare RGBQuad arrays.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Sub CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal dstPointer As Long, ByVal srcPointer As Long, ByVal numOfBytes As Long)

Private m_PaletteFilename As String
Private m_PaletteName As String
Private m_PaletteColorCount As Long
Private m_PaletteColors() As RGBQuad

Friend Function GetPaletteColorCount() As Long
    GetPaletteColorCount = m_PaletteColorCount
End Function

Friend Function GetPaletteFilename() As String
    GetPaletteFilename = m_PaletteFilename
End Function

Friend Function GetPaletteName() As String
    GetPaletteName = m_PaletteName
End Function

Friend Sub SetPaletteName(ByRef newPaletteName As String)
    m_PaletteName = newPaletteName
End Sub

Friend Function CopyPaletteToArray(ByRef dstPalette() As RGBQuad) As Boolean

    CopyPaletteToArray = (m_PaletteColorCount > 0)
    
    If CopyPaletteToArray Then
        
        On Error Resume Next
        If (LBound(dstPalette) <> 0) Or (UBound(dstPalette) <> m_PaletteColorCount - 1) Then ReDim dstPalette(0 To m_PaletteColorCount - 1) As RGBQuad
        On Error GoTo 0
        
        CopyMemory_Strict VarPtr(dstPalette(0)), VarPtr(m_PaletteColors(0)), m_PaletteColorCount * 4
    
    End If

End Function

'Given a path to a supported palette file, return TRUE if the file can be successfully parsed for palette data; FALSE otherwise
Friend Function LoadPaletteFromFile(ByRef srcFile As String) As Boolean

    'Branch according to format; at present, only .gpl files are supported
    LoadPaletteFromFile = LoadPaletteFromGIMPFile(srcFile)
    
End Function

'Given a valid path to a GIMP-format .gpl file, return an array of RGBQuad entries
Private Function LoadPaletteFromGIMPFile(ByRef srcFile As String) As Boolean
    
    On Error GoTo InvalidPalette
    
    'GIMP palettes contain extra information that PD doesn't currently have a way to use.  We only parse the
    ' information we require, so you may need to look elsewhere for a full parsing.
    
    'Dump the full file into a string
    Dim rawFileString As String
    LoadPaletteFromGIMPFile = Files.FileLoadAsString(srcFile, rawFileString, True)
    
    If LoadPaletteFromGIMPFile Then
    
        'GIMP palette files always start with the text "GIMP Palette"
        If Strings.StringsEqual(Left$(rawFileString, 12), "GIMP Palette", True) Then
        
            'This appears to be a valid GIMP palette file.  Hypothetically, line order should be fixed,
            ' but we parse the file as if line order is *not* fixed.  Let me know if you encounter a file
            ' where this approach is invalid.
            
            'Prep our internal palette structs
            ReDim m_PaletteColors(0 To 255) As RGBQuad
            m_PaletteColorCount = 0
            
            'To simplify processing, split the string by lines.
            Dim fileLines As pdStringStack
            Set fileLines = New pdStringStack
            fileLines.CreateFromMultilineString rawFileString
            
            Const SPACE_CHAR As String = " "
            
            'Parse each line in turn
            Dim curLine As String
            Do While fileLines.PopString(curLine)
                
                'AFAIK, there is no formal GIMP spec for palette files.  As such, they can come in a variety
                ' of shapes and sizes.  One thing we want to standardize (to simplify parsing) is replacing
                ' tab chars with space chars; VB's lack of a generic "whitespace" identifier makes this choice
                ' the least of several evils.
                If (InStr(1, curLine, vbTab, vbBinaryCompare) <> 0) Then curLine = Replace$(curLine, vbTab, SPACE_CHAR, , , vbBinaryCompare)
                
                'Empty lines can be ignored
                If (LenB(Trim$(curLine)) = 0) Then
                    'Do nothing
                    
                'Comment lines start with a #; these can be completely ignored
                ElseIf Strings.StringsEqual(Left$(curLine, 1), "#", False) Then
                    'Do nothing
                
                'The palette name is stored on a line prefaced by "Name: "
                ElseIf Strings.StringsEqual(Left$(curLine, 5), "Name:", True) Then
                    m_PaletteName = Trim$(Right$(curLine, Len(curLine) - 5))
                
                'Color descriptor lines contain three numbers, separated by one or more spaces (as the columns
                ' are forcibly aligned).  Here are two examples of valid color lines:
                
                '232   0  50
                ' 26 130  38 ColorNameHere (occurs 6454)
                
                'Because of these variations in formatting, we have to search for colors in a somewhat complicated way.
                ElseIf (m_PaletteColorCount < 256) Then
                    
                    ' Start by looking for at least two spaces in the trimmed string (indicating at least three unique entries)
                    curLine = Trim$(curLine)
                    If (InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) <> InStrRev(curLine, SPACE_CHAR, -1, vbBinaryCompare)) Then
                    
                        'This string contains two spaces.  Extract the first string-delimited entry.
                        Dim targetColor As String, targetColorInt As Long
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        
                        'Attempt to convert this to a number; if it fails, that's okay; this is some kind of invalid line
                        ' and we can ignore further parsing.
                        On Error GoTo BadLineColor
                        targetColorInt = CLng(targetColor)
                        On Error GoTo 0
                        
                        'If we're still here, the color was parsed correctly.  Store it as the red component of the
                        ' current palette entry.
                        m_PaletteColors(m_PaletteColorCount).Red = targetColorInt
                        
                        'Trim the color we've parsed out of the string, then repeat the above steps
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare)))
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        targetColorInt = CLng(targetColor)
                        On Error GoTo 0
                        m_PaletteColors(m_PaletteColorCount).Green = targetColorInt
                        
                        '...and one last time, for the blue component.  Note that the resulting string may not
                        ' have a trailing space, so we forcibly add one to simplify the parser.
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare))) & SPACE_CHAR
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        targetColorInt = CLng(targetColor)
                        On Error GoTo 0
                        m_PaletteColors(m_PaletteColorCount).Blue = targetColorInt
                        
                        'GIMP palettes do not support alpha channels.  Forcibly set a value of 255.
                        m_PaletteColors(m_PaletteColorCount).Alpha = 255
                        
                        'If we made it all the way here, this line was successfully parsed for color data.
                        ' Advance the color count tracker and resume the line parser.
                        m_PaletteColorCount = m_PaletteColorCount + 1
                        
BadLineColor:
                    
                    End If
                    
                End If
            
            Loop
            
            'If we haven't errored out, consider this a valid parse
            LoadPaletteFromGIMPFile = True
            
            'On valid loads, check and remove any duplicate palette entries.  (As an ASCII format, GIMP makes
            ' no guarantees about this, but we don't want to waste time palette-matching duplicate entries.)
            FindAndRemoveDuplicates
            
        End If
    
    End If
    
    If LoadPaletteFromGIMPFile Then m_PaletteFilename = srcFile
    
    Exit Function

InvalidPalette:
    
    m_PaletteName = vbNullString
    m_PaletteColorCount = 0
    m_PaletteFilename = vbNullString
    LoadPaletteFromGIMPFile = False

End Function

'Many palette files are text-based, and nothing prevents problematic situations like duplicate colors.
' (In fact, some shitty formats *cough* Paint.NET *cough* require a *fiex* number of colors, which practically
' guarantees duplicate entries for small palettes.)  After loading a palette from file, I recommend calling this
' function to remove any duplicate palette entries.  It is *not* well-optimized (potentially O(n^2)) but seeing
' as PD only supports 8-bit palettes, I have not yet found it worth the trouble to optimize further.
Private Sub FindAndRemoveDuplicates()

    'Only palettes with multiple entries need to be searched for duplicates
    If (m_PaletteColorCount <= 1) Then Exit Sub
    
    Dim needsRemoval() As Byte
    ReDim needsRemoval(0 To 255) As Byte
    
    Dim i As Long, j As Long, targetValue As RGBQuad, duplicatesFound As Boolean
    For i = 0 To m_PaletteColorCount - 2
        
        'Array accesses are slow in VB; use a local value instead
        targetValue = m_PaletteColors(i)
        
        For j = i + 1 To m_PaletteColorCount - 1
            With m_PaletteColors(j)
                If (.Red = targetValue.Red) And (.Green = targetValue.Green) And (.Blue = targetValue.Blue) And (.Alpha = targetValue.Alpha) Then
                    needsRemoval(j) = 1
                    duplicatesFound = True
                End If
            End With
        Next j
        
    Next i
    
    'Remove all flagged entries
    If duplicatesFound Then
    
        Dim numRemoved As Long
        
        For i = 1 To m_PaletteColorCount - 1
            If (needsRemoval(i) = 0) Then
                m_PaletteColors(i - numRemoved) = m_PaletteColors(i)
            Else
                numRemoved = numRemoved + 1
            End If
        Next i
        
        m_PaletteColorCount = m_PaletteColorCount - numRemoved
        ReDim Preserve m_PaletteColors(0 To m_PaletteColorCount - 1) As RGBQuad
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Palette contained duplicate entries; color count reduced from " & CStr(m_PaletteColorCount + numRemoved) & " to " & CStr(m_PaletteColorCount) & " colors."
        #End If
        
    End If
    
End Sub
