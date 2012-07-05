--EZTwain Classic README.TXT

EZTWAIN Classic (eztw32.dll) is a Windows DLL that provides an easy
interface to the TWAIN image acquisition protocol.

This software is unsupported, public domain software.
Originally written by Spike McLarty and donated to the TWAIN Working Group, hosted for many years by Dosadi, it is now made available by Atalasoft, Inc.
It is free in both senses of the word: No fees, no restrictions.

You can download the latest version of EZTwain Classic, including
C source code and MSVC 6 project files, from:
http://www.eztwain.com/eztwain1.htm

Atalasoft does not officially provide support for EZTwain Classic, but offers a forum for community support at http://www.atalasoft.com/products/dotimage/forums

Atalasoft offers a comprehensive .NET imaging SDK called DotImage:
http://www.atalasoft.com/products/dotimage/

a TWAIN scanning SDK called DotTwain:
http://www.atalasoft.com/products/dotimage/dottwain

and an expanded version of EZTwain Classic (this library) called EZTwain Pro:
http://www.eztwain.com/eztwain3.htm


--History: See \EZTWAIN1\VC\EZTWAIN.H

--Overview of EZTwain

EZTWAIN functions are provided at several levels of abstraction:  At the
highest level, a single call will acquire an image and either return a
handle to it (as a DIB - Device Independent Bitmap) or place it in the
Windows clipboard.  There is also a single function that does all the work
of the 'Select Source' File menu item.

At the lowest level, there are functions to send individual triplets to the
Source Manager or to the currently open Data Source.


EZTWAIN is specifically designed to simplify the use of TWAIN in three cases:

** When the development language is not a mainstream C/C++ compiler - for
   example, Visual Basic, Application Basic, database programming languages,
   interpreted languages, LabView, Pascal, LISP, etc.

   With these languages, using the standard TWAIN.DLL entry point directly
   may be difficult or even impossible.  I hope that EZTWAIN will bridge
   this gap.

** When the image acquisition requirements are very simple.  Many developers
   just want to 'grab an image' without spending days or weeks understanding
   TWAIN and its subtleties.  For these developers, the full TWAIN protocol
   is expensive overkill - EZTWAIN may meet their needs with a much
   smaller learning and programming effort.

** Even skilled C programmers with weeks or months to spend on a full
   TWAIN implementation can still become frustrated trying to get their code
   to the point where it actually 'does something'.  Calling the high-level
   EZTWAIN functions should produce results in an hour or two, and then the
   high-level calls can be refined into calls to functions at lower and
   lower levels as needed to achieve more exotic effects.


-- Contents of the EZTwain Classic package (eztw1.zip)

If you unzip the distribution package to the root of a drive using
the path information, you'll get the folder \EZTWAIN1 at the top,
and the following files inside:

README.TXT                      this file.
ACCESS\EZTWAIN.BAS              declaration file for MS Access
CLARION\EZTW32.LIB              object library for use with Clarion
CLARION\EZTWAIN.CLW             Clarion declarations (I think...)
CLARION\README.TXT              notes on using EZTwain from Clarion
CSHARP\EZTWAIN.CS               C# declarations
DBASE\EZTWAIN.H                 declarations for dBase
DELPHI\EZTWAIN.PAS              declarations for Delphi
LotusScript\eztwain.lss         declarations for Lotus Notes
PERL\EZTWAIN.PL                 declarations for Perl
PowerBASIC\eztwain.inc          declarations for PowerBASIC
PowerBuilder\eztwain.txt        declarations for PowerScript 10+
PowerBuilder\eztwain-pb9.txt    declarations for PowerScript 9
PROGRESS\EZTWAIN.I              declarations for Progress
VB.NET\EZTWAIN.VB               declarations for VB.NET
VB\EZTWAIN.BAS                  declarations for VB5/6
VC\EZTWAIN.C                    source code of EZTwain DLL
VC\EZTWAIN.H                    declarations for C
VC\RESOURCE.H                   resource defines
VC\TWAIN.H                      TWAIN standard header file
VC\TWERP.C                      Twerp sample application, in C
VC\TWERP.ICO                    Twerp application icon
VC\TWERP.RC                     Twerp application resources
VC\VC.DSW                       MSVC 6.0 workspace for Eztwain and Twerp
VC\EZTW32\EZTW32.DEF            export definition file for eztw32.dll
VC\EZTW32\EZTW32.DSP            MSVC project file for eztw32.dll
VC\RELEASE\EZTW32.DLL           pre-built release version of EZTwain DLL
VC\RELEASE\EZTW32.LIB           link library of eztw32.dll
VC\RELEASE\TWERP32.EXE          pre-built Twerp sample app
VC\TWERP32\TWERP32.DSP          MSVC 6 project file for Twerp
VFP\EZTWAIN.PRG                 declarations for Visual FoxPro


-- Using EZTwain Classic

The EZTWAIN interface was designed to have easy-to-remember names; to
require very few parameters to functions; and to not require the use of
any data structures.  For the high-level functions, passing 0 for all
parameters is a good way to start.  

If you are programming in C, examine EZTWAIN.H and TWERP.C to see
how the EZTWAIN functions are declared and called.  You may have to 
configure your system to find the EZTWAIN.H file during compilation,
the EZTW32.LIB file during linking, and the EZTWAIN.DLL or EZTW32.DLL,
during execution.

For Microsoft Visual C++, these issues are taken care of if you open the
included project files - all the necessary files are found automatically.


For other development platforms, you will have to study your documentation
to learn three things:

1. How to declare DLL functions to be called.

If your platform can do this, it will be covered in the documentation, and
almost any programming language for Windows has a way to do this.

2. How to link to the DLL.

For C, C++, and Clarion you use a link-library, EZTW32.LIB - note that
Clarion needs a different format, so there are two copies of this file
in the package.
Your documentation should have a section like 'Linking to DLLs'.

For dynamic languages like Basic, you typically specify the name of the
DLL in some way, at or near the code that declares the EZTWAIN functions.

3. How to access the DLL at run-time.

Our declaration files for various languages will usually find
eztw32.dll if it is either in System32, or if it is placed in the same
folder as the .EXE that runs the application.  It is not uncommon to
place eztw32.dll in the System32 folder (under the Windows home folder)
but this can cause problems if two applications are installed with
different versions of eztw32.dll - we recommend trying to install the
eztw32.dll in the application folder if that's possible.


-- How should I design TWAIN support into my program?

Start out by calling TWAIN_SelectImageSource - this call is simple, and
should be the easiest to debug.  Alternatively, you could put in code to
call TWAIN_IsAvailable.  Use the result of this call to disable, gray or
hide, or to enable your TWAIN facilities.

Then add the calls to TWAIN_AcquireNative or TWAIN_AcquireToClipboard.
Don't worry about passing proper Window handles at first.  Notice also
that TWAIN_AcquireNative will disable and re-enable your app window, once
you decide to pass it in.

Once you have Select Source and Acquire working in your program, then it's
just a matter of deciding how much more you want.  The simple behavior of
AcquireNative is attractive but limiting.  AcquireNative treats every
Data Source as modal - disabling & closing the source after a single image
is acquired.  To allow e.g. acquisition of multiple images, or
modeless Data Sources to function properly, you will have to dig into
the source code for EZTWAIN and make lower level calls, or adapt the
EZTWAIN code to create your own DLL.


-- What to do with problems, questions and suggestions

If you have a specific problem to report, or a question or suggestion for
the developers of EZTWAIN, visit the EZTwain forum at
http://www.atalasoft.com/products/dotimage/forums

If you can't find any helpful information in the existing posts, start
a new thread in the EZTwain Classic forum, or send an e-mail to:

				support@eztwain.com


January 27, 2011
Spike McLarty
Principal Engineer, Dosadi
www.dosadi.com
