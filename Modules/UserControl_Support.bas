Attribute VB_Name = "UserControl_Support"
'***************************************************************************
'Helper functions for various PhotoDemon UCs
'Copyright 2014-2015 by Tanner Helland
'Created: 06/February/14
'Last updated: 20/August/15
'Last update: start migrating various UC-inspecific functions here
'
'Many of PD's custom user controls share similar functionality.  Rather than duplicate that functionality across
' all controls, I've tried to collect reusable functions here.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The tab key presents a particular problem when intermixing API windows and VB controls.  VB (obviously) ignores the
' API windows entirely, and PD further complicates this by sometimes mixing API windows and VB controls on the same UC.
' To avoid this disaster, we manage our own tab key presses using an automated system that sorts controls from top-left
' to bottom-right, and automatically figures out tab order from there.
'
'To make sure the automated system works correctly, some controls actually raise a TabPress event, which their parent
' UC can use to cycle control focus within the UC.  When focus is tabbed-out from the last control on the UC, the UC
' itself can then notify the master TabHandler to pass focus to an entirely new control.
Public Enum PDUC_TAB_BEHAVIOR
    TabDefaultBehavior = 0
    TabRaiseEvent = 1
End Enum

#If False Then
    Private Const TabDefaultBehavior = 0, TabRaiseEvent = 1
#End If
