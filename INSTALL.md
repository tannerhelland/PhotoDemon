Installing PhotoDemon
=====================

PhotoDemon is a portable application.  **It does not require installation.**  

To use PhotoDemon, simply [download the latest release](https://github.com/tannerhelland/PhotoDemon/releases), extract the full contents of the .zip file to your PC, then double-click `PhotoDemon.exe`.

System Requirements
===================

PhotoDemon supports all modern versions of Windows (Windows XP through the latest Windows 11 builds).

Beyond this, it has only one fixed hardware requirement: a minimum display resolution of at least 1024x768 pixels.  

As with any software, a faster processor and plenty of free RAM and disk space provides a better experience, but current PhotoDemon builds run very well on XP-era processors with 5400 RPM HDDs and < 1 GB of RAM.  (Yes, I still test these configurations!)

Besides `PhotoDemon.exe`, the program also requires access to an `App/PhotoDemon/Plugins` subfolder.  This folder ships with PhotoDemon and it contains a number of essential 3rd-party libraries.  PhotoDemon will not run if this folder is missing or broken.

If PhotoDemon fails to start on your PC, please ensure that the `App/PhotoDemon/Plugins` subfolder is intact.  Nearly all startup problems are caused by ancient ZIP software (e.g. WinZip) failing to extract PhotoDemon's folder tree correctly.  If you don't see an `/App` subfolder, or if you see any .dll files in the base PhotoDemon folder, please re-extract PhotoDemon and its dependencies using a different ZIP tool.

Building PhotoDemon
===================

PhotoDemon is written in Visual Basic 6.0.  Building it is as simple as:

1) Load PhotoDemon.vbp into the VB6 IDE
2) Click File > Make PhotoDemon.exe
3) Click OK

Your VB6 copy should be completely up-to-date, with the latest SP6 update(s) installed.  No support is provided for other configurations.

Feel free to modify most settings in the project's compile options, with the following caveats:

1) PhotoDemon is *extremely* slow when compiled to P-Code.  **Compile to native code only.**
2) Do **not** enable the `Assume No Aliasing` advanced optimization.  PhotoDemon uses many aliasing tricks to improve performance, and the `Assume No Aliasing` optimization will produce buggy code.  All other advanced optimizations can (and should) be enabled.
3) Optimizing for fast vs small code makes little difference.  Choose whichever option you like.

PhotoDemon doesn't reference any external OCX or ActiveX DLL files, so you *do not* need to run the VB6 IDE elevated when building or testing the project.  In fact, for security purposes I strongly recommend *not* running the VB6 IDE elevated when running or building open-source projects you have not manually vetted.

The OS used for compiling does not matter.  (For example, you can compile on Windows 11 and run the resulting PhotoDemon.exe file on XP without problems.)  PhotoDemon is primarily developed on current Windows 11 builds, and limited compatibility testing is still performed on XP, Windows 7, and Windows 10 PCs.  Vista and Windows 8/8.1 compatibility relies on user-submitted bug reports, as I no longer keep dedicated VMs around for testing.

Finally, despite being built in VB6, PhotoDemon never requires any special compatibility modes or other modifications.  (In fact, it may break if you apply compatibility shims.)  You should always simply build and run it as-is.
