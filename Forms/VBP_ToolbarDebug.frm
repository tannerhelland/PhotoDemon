VERSION 5.00
Begin VB.Form toolbar_Debug 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Debug"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   54
   ScaleMode       =   2  'Point
   ScaleWidth      =   138.75
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrDebug 
      Interval        =   1000
      Left            =   2280
      Top             =   600
   End
   Begin PhotoDemon.pdLabel lblDIB 
      Height          =   195
      Index           =   0
      Left            =   75
      Top             =   75
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   370
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCustomForeColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblDIB 
      Height          =   195
      Index           =   1
      Left            =   75
      Top             =   375
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   370
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCustomForeColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblDIB 
      Height          =   195
      Index           =   2
      Left            =   75
      Top             =   675
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   370
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCustomForeColor=   -1  'True
   End
End
Attribute VB_Name = "toolbar_Debug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Debug Window
'Copyright ©2013-2014 by Tanner Helland
'Created: 30/October/14
'Last updated: 30/October/14
'Last update: initial build
'
'As part of the 6.6 release, I'd like to optimize as much of PD's UI code as possible.  There are a lot of custom UC
' elements in these newer builds, and I want to make sure any memory leaks are caught early in the development cycle.
' Also, some of the older UCs (including slider) use very poor buffering strategies, so a lot of unnecessary temp
' DIBs are created during drawing.
'
'To that end, I've created this small debug window to help me track creation and destruction of certain objects
' throughout the program's lifecycle.  This will hopefully help me detect any problematic areas, and cut down on the
' amount of memory PD churns on UC work.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub Form_Load()
    
    #If DEBUGMODE = 1 Then
        tmrDebug.Enabled = True
        Call tmrDebug_Timer
    #Else
        tmrDebug.Enabled = False
    #End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrDebug.Enabled = False
End Sub

'Every second, update the on-screen labels with a report of the global DIB counter variables
Private Sub tmrDebug_Timer()

    lblDIB(0).Caption = "DIBs created: " & g_DIBsCreated
    lblDIB(1).Caption = "DIBs destroyed: " & g_DIBsDestroyed
    lblDIB(2).Caption = "DIBs active: " & CStr(g_DIBsCreated - g_DIBsDestroyed)
        
End Sub
