VERSION 5.00
Begin VB.Form FormHotkeys 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Keyboard shortcuts"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   5250
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "FormHotkeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Customizable hotkeys dialog
'Copyright 2024-2024 by Tanner Helland
'Created: 09/September/24
'Last updated: 09/September/24
'Last update: initial build
'
'This dialog allows the user to customize hotkeys.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_Menus() As PD_MenuEntry, m_NumOfMenus As Long

Private Sub cmdBar_OKClick()
    
    'TODO
    
End Sub

Private Sub cmdBar_ResetClick()
    'TODO
End Sub

Private Sub Form_Load()
    
    'Retrieve a copy of all menus (including hierarchies and attributes) from the menu manager
    m_NumOfMenus = Menus.GetCopyOfAllMenus(m_Menus)
    
    
    'TODO
    
    'Apply custom themes
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Interface.ReleaseFormTheming Me
End Sub

