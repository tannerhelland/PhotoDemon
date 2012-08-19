VERSION 5.00
Begin VB.Form FormSoftwareUpdate 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " This version of PhotoDemon is out-of-date"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSoftwareUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Automatic Software Updater (note: it doesn't do the actual updating, it just CHECKS for updates!)
'Copyright ©2000-2012 by Tanner Helland
'Created: 19/August/12
'Last updated: 19/August/12
'Last update: initial build
'
'Interface for notifying the user that a new version of PhotoDemon is available for download.  This code is simply the
' notification part; the actual update checking is handled within the SoftwareUpdater module.
'
'Note that this code interfaces with the .INI file so the user can opt to not check for updates and never be
' notified again. (FYI - this option can be enabled/disabled from the 'Edit' -> 'Program Preferences' menu.)
'
'***************************************************************************

Option Explicit

