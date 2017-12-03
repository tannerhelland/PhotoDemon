If you want to change PhotoDemon's resource file, you first need to modify the PD_icons.rc 
resource script.  Then you need to recompile the .RES file using Microsoft's rc.exe and 
rcdll.dll files.  Finally, you must manually insert a valid manifest to the resource file 
using the instructions at this link:
 http://www.vbforums.com/showthread.php?606736-VB6-XP-Vista-Win7-Manifest-Creator

I cannot provide more support than this, as PhotoDemon is only designed to work with its 
original resource file.  Modified resource files may causes crashes, errors, or instability,
so fiddle with them at your own risk.