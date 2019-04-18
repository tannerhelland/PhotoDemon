Building PhotoDemon
===================

PhotoDemon is written in Visual Basic 6.0.  Building it is as simple as:

1) Load PhotoDemon.vbp into the VB6 IDE
2) Click File > Make PhotoDemon.exe
3) Click OK

Feel free to modify most settings in the project's compile options, with the following caveats:

1) PhotoDemon is extremely slow when compiled to P-Code.  Compile to native code only.
2) Do **not** enable the Assume No Aliasing advanced optimization.  PhotoDemon uses many aliasing tricks to improve performance, and the Assume No Aliasing optimization will produce buggy code.  All other advanced optimizations can (and should) be enabled.
3) Optimizing for fast vs small code makes little difference.  Choose whatever option you like.

PhotoDemon does not utilize any external OCX or ActiveX DLL files, so you *do not* need to run the VB6 IDE elevated when building or testing the project.  In fact, for security purposes I strongly recommend *not* running the VB6 IDE elevated when working with open-source projects you have not manually vetted.

The OS used for compilation does not matter; for example, you can compile on Win 10 and run the resulting PhotoDemon.exe file on XP without problems (and vice versa).  

PhotoDemon is primarily developed on current Win 10 builds, and limited compatibility testing is still performed on XP and Win 7 PCs. Vista and Win 8/8.1 compatibility relies on user-submitted bug reports, as I no longer keep dedicated PCs (or VMs) around for testing.

Finally, despite being built in VB6, PhotoDemon never requires any special compatibility modes or other modifications.  (In fact, it may break if you apply compatibility shims to it.)  You should simply build it and run it as-is.

Installing PhotoDemon
=====================

PhotoDemon is a portable application.  It does not require installation.  Any modern version of Windows (XP through the latest Win 10 builds) is fully supported.

Besides PhotoDemon.exe, the program also requires access to a App/PhotoDemon/Plugins folder with all plugins available from [the official PhotoDemon repo](https://github.com/tannerhelland/PhotoDemon).  Of particular importance are the [Zstandard](https://github.com/facebook/zstd), [lz4](https://github.com/lz4/lz4), [libdeflate](https://github.com/ebiggers/libdeflate) and [LittleCMS](https://github.com/mm2/Little-CMS) libraries.  PhotoDemon will not work if these libraries are missing or broken.

Zstandard, lz4, and libdeflate use the official binaries supplied by each project's authors, and your own binaries (default cdecl calling convention) can theoretically be dropped-in without problems.  LittleCMS is currently custom-built owing to some [serious](https://github.com/mm2/Little-CMS/issues/162) [bugs](https://github.com/mm2/Little-CMS/issues/179) in the official binary, and I cannot guarantee that a default LittleCMS build will work without manually fixing those linked issues; as such, please use the copy that ships with PhotoDemon unless you are comfortable fixing those issues yourself.