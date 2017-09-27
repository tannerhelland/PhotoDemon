Attribute VB_Name = "ColorPicker"
'***************************************************************************
'Color Picker Tool Manager
'Copyright 2017-2017 by Tanner Helland
'Created: 25/September/17
'Last updated: 27/September/17
'Last update: wrap up initial build
'
'At present, this module is just a thin wrapper to the toolpanel_ColorPicker form (where the *real* fun happens).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Current mouse/pen input values.  These are blindly relayed to us by the canvas, and it's up to us to perform
' any special tracking calculations.
'IMPORTANT NOTE: these coordinates have already been translated into the *image* coordinate space.
Private m_MouseDown As Boolean
Private m_MouseX As Single, m_MouseY As Single

'Similar to other tools, the color picker is notified of all mouse actions that occur while it is selected.
' At present, however, it only triggers a fill when the mouse is actually clicked.  (No action is taken on
' move events, unless the mouse button is down.)
Public Sub NotifyMouseXY(ByVal mouseButtonDown As Boolean, ByVal imgX As Single, ByVal imgY As Single, ByRef srcCanvas As pdCanvas)
    
    'Before doing anything else, cache the mouse coordinates in case we need them in the future
    Dim isFirstStroke As Boolean, isLastStroke As Boolean
    isFirstStroke = (Not m_MouseDown) And mouseButtonDown
    isLastStroke = m_MouseDown And (Not mouseButtonDown)
    
    m_MouseDown = mouseButtonDown
    m_MouseX = imgX
    m_MouseY = imgY
    
    'Because this tool is largely UI-centric, start by forwarding the mouse position to the toolbox itself.
    ' It will handle any required on-screen updates.
    toolpanel_ColorPicker.NotifyCanvasXY m_MouseDown, m_MouseX, m_MouseY, srcCanvas
    
End Sub
