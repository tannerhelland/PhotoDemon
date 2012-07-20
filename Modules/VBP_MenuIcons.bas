Attribute VB_Name = "Menu_Icons"
'***************************************************************************
'Menu Icon Handler
'©2011-2012 Tanner Helland
'Created: 24/June/12
'Last updated: 24/June/12
'Last update: Initial build
'
'Because VB6 doesn't provide a way to add icons/bitmaps to menus, an external method is required.
' I've chosen to use the excellent cMenuImage class by Leandro Ascierto.  Menu icons are extracted
' from a resource file (where they're stored in PNG format) and rendered to the menu at run-time.
'
' NOTE: Because the Windows XP version of this code utilizes potentially dirty subclassing,
' I have disabled menu icons while running in the IDE on Windows XP.  Compile the project to see icons.
' (Windows Vista and 7 use a different mechanism, so menu icons are enabled in the IDE.)
'
'***************************************************************************
  
Dim cMenuImage As clsMenuImage

Public Sub LoadMenuIcons()

    Set cMenuImage = New clsMenuImage

    With cMenuImage
    
        'Disable menu icon drawing if on Windows XP and uncompiled (to prevent subclassing crashes on unclean IDE breaks)
        If (Not .IsWindowVistaOrLater) And (IsProgramCompiled = False) Then Exit Sub
        
        .Init FormMain.HWnd, 16, 16
        
        .AddImageFromStream LoadResData("OPENIMG", "CUSTOM")    '0
        .AddImageFromStream LoadResData("OPENREC", "CUSTOM")    '1
        .AddImageFromStream LoadResData("IMPORT", "CUSTOM")     '2
        .AddImageFromStream LoadResData("SAVE", "CUSTOM")       '3
        .AddImageFromStream LoadResData("SAVEAS", "CUSTOM")     '4
        .AddImageFromStream LoadResData("CLOSE", "CUSTOM")      '5
        .AddImageFromStream LoadResData("BCONVERT", "CUSTOM")   '6
        .AddImageFromStream LoadResData("PRINT", "CUSTOM")      '7
        .AddImageFromStream LoadResData("SCANNER", "CUSTOM")    '8
        .AddImageFromStream LoadResData("DOWNLOAD", "CUSTOM")   '9
        .AddImageFromStream LoadResData("SCREENCAP", "CUSTOM")  '10
        .AddImageFromStream LoadResData("FRXIMPORT", "CUSTOM")  '11
        .AddImageFromStream LoadResData("UNDO", "CUSTOM")       '12
        .AddImageFromStream LoadResData("REDO", "CUSTOM")       '13
        .AddImageFromStream LoadResData("REPEAT", "CUSTOM")     '14
        .AddImageFromStream LoadResData("COPY", "CUSTOM")       '15
        .AddImageFromStream LoadResData("PASTE", "CUSTOM")      '16
        .AddImageFromStream LoadResData("CLEAR", "CUSTOM")      '17
        .AddImageFromStream LoadResData("PREFERENCES", "CUSTOM") '18
        .AddImageFromStream LoadResData("RESIZE", "CUSTOM")     '19
        .AddImageFromStream LoadResData("ROTATECW", "CUSTOM")   '20
        .AddImageFromStream LoadResData("ROTATECCW", "CUSTOM")  '21
        .AddImageFromStream LoadResData("ROTATE180", "CUSTOM")  '22
        .AddImageFromStream LoadResData("FLIP", "CUSTOM")       '23
        .AddImageFromStream LoadResData("MIRROR", "CUSTOM")     '24
        .AddImageFromStream LoadResData("PDWEBSITE", "CUSTOM")  '25
        .AddImageFromStream LoadResData("FEEDBACK", "CUSTOM")   '26
        .AddImageFromStream LoadResData("ABOUT", "CUSTOM")      '27
        .AddImageFromStream LoadResData("FITWINIMG", "CUSTOM")  '28
        .AddImageFromStream LoadResData("FITONSCREEN", "CUSTOM") '29
        .AddImageFromStream LoadResData("TILEHOR", "CUSTOM")    '30
        .AddImageFromStream LoadResData("TILEVER", "CUSTOM")    '31
        .AddImageFromStream LoadResData("CASCADE", "CUSTOM")    '32
        .AddImageFromStream LoadResData("ARNGICONS", "CUSTOM")  '33
        .AddImageFromStream LoadResData("MINALL", "CUSTOM")     '34
        .AddImageFromStream LoadResData("RESTOREALL", "CUSTOM")     '35
        .AddImageFromStream LoadResData("OPENMACRO", "CUSTOM")  '36
        .AddImageFromStream LoadResData("RECORD", "CUSTOM")     '37
        .AddImageFromStream LoadResData("RECORDSTOP", "CUSTOM") '38
        .AddImageFromStream LoadResData("BUG", "CUSTOM") '39
        .AddImageFromStream LoadResData("FAVORITE", "CUSTOM") '40
        
        
        'File Menu
        .PutImageToVBMenu 0, 0, 0       'Open Image
        .PutImageToVBMenu 1, 1, 0       'Open recent
        .PutImageToVBMenu 2, 2, 0       'Import
        .PutImageToVBMenu 3, 4, 0       'Save
        .PutImageToVBMenu 4, 5, 0       'Save As...
        .PutImageToVBMenu 5, 7, 0       'Close...
        .PutImageToVBMenu 6, 9, 0       'Batch conversion
        .PutImageToVBMenu 7, 11, 0      'Print
        
        '--> Import Sub-Menu
        'NOTE: the specific menu values will be different if the scanner plugin (eztw32.dll) isn't found.
        If ScanEnabled = True Then
            .PutImageToVBMenu 8, 0, 0, 2       'Scan Image
            .PutImageToVBMenu 9, 3, 0, 2       'Download Image
            .PutImageToVBMenu 10, 5, 0, 2      'Capture Screen
            .PutImageToVBMenu 11, 7, 0, 2      'Import from FRX
        Else
            .PutImageToVBMenu 9, 0, 0, 2       'Download Image
            .PutImageToVBMenu 10, 2, 0, 2      'Capture Screen
            .PutImageToVBMenu 11, 4, 0, 2      'Import from FRX
        End If
        
        'Edit Menu
        .PutImageToVBMenu 12, 0, 1      'Undo
        .PutImageToVBMenu 13, 1, 1      'Redo
        .PutImageToVBMenu 14, 2, 1      'Repeat Last Action
        .PutImageToVBMenu 15, 4, 1      'Copy
        .PutImageToVBMenu 16, 5, 1      'Paste
        .PutImageToVBMenu 17, 6, 1      'Empty Clipboard
        .PutImageToVBMenu 18, 8, 1      'Program Preferences
        
        'Image Menu
        .PutImageToVBMenu 19, 0, 2      'Resize
        .PutImageToVBMenu 20, 0, 2, 3     'Rotate Clockwise (rotate submenu)
        .PutImageToVBMenu 21, 1, 2, 3     'Rotate Counter-clockwise (rotate submenu)
        .PutImageToVBMenu 22, 2, 2, 3     'Rotate 180 (rotate submenu)
        .PutImageToVBMenu 23, 4, 2      'Flip
        .PutImageToVBMenu 24, 5, 2      'Mirror
        
        'Macro Menu
        .PutImageToVBMenu 36, 0, 5     'Open Macro
        .PutImageToVBMenu 37, 2, 5     'Start Recording
        .PutImageToVBMenu 38, 3, 5     'Stop Recording
        
        'Window Menu
        .PutImageToVBMenu 28, 0, 6     'Fit Window to Image
        .PutImageToVBMenu 29, 1, 6     'Fit on Screen
        .PutImageToVBMenu 30, 3, 6     'Tile Horizontally
        .PutImageToVBMenu 31, 4, 6     'Tile Vertically
        .PutImageToVBMenu 32, 5, 6     'Cascade
        .PutImageToVBMenu 33, 6, 6     'Arrange Icons
        .PutImageToVBMenu 34, 8, 6     'Minimize All
        .PutImageToVBMenu 35, 9, 6     'Restore All
        
        'Help Menu
        .PutImageToVBMenu 40, 0, 7     'Donate
        .PutImageToVBMenu 25, 2, 7     'Visit the PhotoDemon website
        .PutImageToVBMenu 26, 3, 7     'Submit Feedback
        .PutImageToVBMenu 39, 4, 7     'Submit Bug
        .PutImageToVBMenu 27, 6, 7     'About PD
    
    End With

End Sub
