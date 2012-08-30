VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Layer class
'Copyright ©2000-2012 by Tanner Helland
'Created: 29/August/12
'Last updated: 29/August/12
'Last update: initial build
'
'The time has come to break PhotoDemon's reliance on picture boxes.  The new plan is to implement an honest-to-God
' DIB class that will manage all image data in memory.  If built correctly, this thing will provide a whole host of
' benefits, including but not limited to:
'
'- Ability to load large images without problems
'- Alpha channels
'- High bit depths
'- Much faster
'- Eventual implementation of layers and adjustment layers
'
'Those are some serious benefits.
'
'Implementation will be slow and steady at first.  A rough roadmap is as follows:
'- Get image loading working.  This will require a lot of plumbing to already be in place, but I can't do anything else
' until I can get image data into the class.
'- Get viewport rendering working.  If I can't see the image, I can't work on it.
'- Experiment with tying the class's data into GetImageData, which is already well-implemented throughout PD.
'- Replace PicWidthL and PicHeightL with width/height values provided by this class.
'- Remove all direct intervention with ActiveForm.BackBuffer and force it through this class.  That means no unmanaged
' changing of width, height, etc - it has to come through here.
'
'This class will be a subset of pdImage().  Initially there will only be one layer (the image's data), but in the future,
' assuming this class is written well, a pdImage object could theoretically store many layers.
'
'Let the fun begin.
'
'***************************************************************************
