VERSION 5.00
Begin VB.Form FormThemeEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Theme editor"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10290
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
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   686
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdLabel pdLabelTitle 
      Height          =   285
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   503
      Caption         =   "Unicode text box for testing:"
      FontSize        =   12
      Layout          =   2
   End
   Begin PhotoDemon.pdTextBox pdTextBox1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3836
   End
End
Attribute VB_Name = "FormThemeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

