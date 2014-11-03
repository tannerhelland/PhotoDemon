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
   Begin PhotoDemon.pdTextBox pdTextBox1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2143
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "current theme path:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2115
   End
End
Attribute VB_Name = "FormThemeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

