VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSVG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon SVG (Scalable Vector Graphics) Container and Parser
'Copyright 2022-2022 by Tanner Helland
'Created: 02/February/22
'Last updated: 02/February/22
'Last update: start work on initial build
'
'This class (and its associated pdSVG- child classes) handle SVG image parsing.
' All code has been custom-built for PhotoDemon, with a special emphasis on parsing performance.
'
'At present, work is focused solely on import support.  Export support may be added later.
'
'As with all 3rd-party formats, SVGs have many features that don't have direct analogs in PhotoDemon.
' Such features are generally extracted by this class, but they will not "appear" in the final loaded
' image.  My ongoing goal is to expand support in this class as various SVG features are implemented
' in PD itself.
'
'Finally, there is always a possibility that a particular file format may prove too cumbersome to
' fully support. I make no guarantee that this class will ever reach full integration into PhotoDemon.
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work off the
' official SVG spec at this URL (link good as of February 2022):
' https://www.w3.org/TR/SVG2/
'
'Because the modern SVG spec is so complex (with elements like animation that don't really matter to
' an editor like PhotoDemon), I've primarily concentrated my efforts on supporting core SVG features
' as described in the "Tiny SVG" spec (link good as of February 2022):
' https://www.w3.org/TR/SVGTiny12/
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

