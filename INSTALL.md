Installing PhotoDemon
=====================

PhotoDemon is a portable application.  **It does not require installation.**  Any modern version of Windows (XP through the latest Win 10 builds) is fully supported.

Besides `PhotoDemon.exe`, the program also requires access to an `App/PhotoDemon/Plugins` subfolder containing the 3rd-party libraries found at [the official PhotoDemon repo](https://github.com/tannerhelland/PhotoDemon).  Of particular importance are the [Zstandard](https://github.com/facebook/zstd), [lz4](https://github.com/lz4/lz4), [libdeflate](https://github.com/ebiggers/libdeflate) and [LittleCMS](https://github.com/mm2/Little-CMS) libraries.  PhotoDemon will not run if these libraries are missing or broken.

If you encounter problems starting PhotoDemon, please ensure that the `App/PhotoDemon/Plugins` subfolder is intact.  99% of startup problems are caused by ancient .zip software (e.g. WinZip) that fails to extract PhotoDemon's folder tree correctly.  If you don't see that plugin folder, or if you see a bunch of .dll files crammed into the base PhotoDemon folder, please re-extract PhotoDemon and its dependencies using the built-in Windows .zip manager.

Building PhotoDemon
===================

PhotoDemon is written in Visual Basic 6.0.  Building it is as simple as:

1) Load PhotoDemon.vbp into the VB6 IDE
2) Click File > Make PhotoDemon.exe
3) Click OK

Your VB6 copy should be completely up-to-date, with the latest SP6 update(s) installed.  No support is provided for other configurations.

Feel free to modify most settings in the project's compile options, with the following caveats:

1) PhotoDemon is *extremely* slow when compiled to P-Code.  Compile to **native code only.**
2) Do **not** enable the `Assume No Aliasing` advanced optimization.  PhotoDemon uses many aliasing tricks to improve performance, and the `Assume No Aliasing` optimization will produce buggy code.  All other advanced optimizations can (and should) be enabled.
3) Optimizing for fast vs small code makes little difference.  Choose whichever option you like.

PhotoDemon doesn't reference any external OCX or ActiveX DLL files, so you *do not* need to run the VB6 IDE elevated when building or testing the project.  In fact, for security purposes I strongly recommend *not* running the VB6 IDE elevated when working with open-source projects you have not manually vetted.

The OS used for compilation does not matter; for example, you can compile on Win 10 and run the resulting PhotoDemon.exe file on XP without problems (and vice versa).  PhotoDemon is primarily developed on current Win 10 builds, and limited compatibility testing is still performed on XP and Win 7 PCs.  Vista and Win 8/8.1 compatibility relies on user-submitted bug reports, as I no longer keep dedicated VMs around for testing.

Finally, despite being built in VB6, PhotoDemon never requires any special compatibility modes or other modifications.  (In fact, it may break if you apply compatibility shims.)  You should always simply build it and run it as-is.
