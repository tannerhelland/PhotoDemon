# PhotoDemon 6.8 pre-alpha

![PhotoDemon Screenshot] (http://photodemon.org/images/PD_66_b1.jpg)

### PhotoDemon is a portable photo editor focused on performance and usability.  

It provides a comprehensive selection of photo editing tools in an 8 MB download.  It runs on any Windows machine (XP through Win 10 TP) and *does not* require installation.  It can easily be run from a USB stick or SD card.  Translations are currently provided for English and eight other languages.

PhotoDemon is completely open-source and available under a permissive BSD license.  Outside contributions from coders, designers, translators, and enthusiasts are always welcome.

For information on the most recent release, please visit:
http://photodemon.org

If you obtained PhotoDemon from its official GitHub repository, please note that the repository *does not contain a compiled EXE.*  If you don't have access to a VB6 compiler, you can download a compiled .exe (including language files and core plugins), updated nightly, from:
http://photodemon.org/downloads/nightly/PhotoDemon_nightly.zip

***

## What makes PhotoDemon unique?

### It is lightweight and completely portable
PhotoDemon is designed to be run as a standalone program. No installer is provided or required.  It does not touch the Windows registry, and aside from a temporary file folder – which you can specify in the Tools > Options dialog – it leaves no trace of itself on your hard drive.  Many people choose to run PhotoDemon from a USB drive.

### It integrates macro recording and batch processing
Complex editing actions can be automated by recording them as macros (similar to Office software).  Once recorded, a macro can be applied to other images.  Macros integrate with a built-in batch processing tool, so you can choose a saved macro and a folder or list of images, and PhotoDemon will apply the macro to every image automagically.

### It emphasizes usability
Most free, open-source image editors are usability nightmares. PhotoDemon tries not to be. The interface was built by designers (not engineers), and small touches like save/load presets on all tools, automatic last-used settings preservation, unlimited Undo/Redo, "Fade last effect", keyboard accelerators, effect previews, mouse wheel and X-button support, and descriptive menu icons make PhotoDemon easy to use for both novices and professionals.

### It provides a comprehensive selection of pro-grade features and tools
* Extensive file format support, including all major RAW formats
* Powerful selection tools, with support for antialiasing, feathering, and on-canvas sizing/moving
* Color-managed workflow, including full support for embedded ICC profiles
* Non-destructive editing for select features, including resizing and key adjustments (exposure, clarity, vibrance, etc)
* 2D transformations: advanced rescale operators (Sinc, Catmull-Rom, etc), content-aware scaling (seam carving), crop, rotate, shear, zoom, tiling
* Pro adjustment tools: levels, curves, HDR, white balance, split-toning, sepia, full-featured histogram, green screen, Wratten filters, and many more
* Filters and effects: perspective correction, edge detection, noise removal, content-aware blur, unsharp masking, lens diffraction, vignetting, film grain, and many more
* 100+ tools are provided in the current build.

### What doesn't PhotoDemon do?

* The current release (6.6) does not support text layers.  Text layers are planned for the next release.
* The current release (6.6) includes partial Unicode support.  Some features may not work with Unicode filenames, but work on this is ongoing.
* The current release (6.6) does not provide any on-canvas painting tools.  On-canvas paint tools are planned for the next release.
* The current release (6.6) may not integrate well with high-contrast Windows themes, or non-standard Windows themes.  Improvements to theming are planned for the next release.
* PhotoDemon isn't designed for OSes other than Microsoft Windows.  A compatibility layer like Wine (http://www.winehq.org/) may allow PhotoDemon to work on OSX, Linux, BSD, Solaris, or Maemo systems, but program stability and performance may suffer.

### How can I get involved? 
PhotoDemon is maintained by a single individual with a family to support.  The software is provided free-of-charge under a permissive open-source license, and no fees or money will ever be charged for its use.

That said, donations go a long way toward supporting the development of this powerful photo editing tool. If you would like to donate and support development, please visit:

http://photodemon.org/donate/

If you can't contribute monetarily to the project, here are other ways to help:
* Let me know if you find any bugs. Issues can be submitted via PhotoDemon's official bug tracker: https://github.com/tannerhelland/PhotoDemon/issues, or this dedicated feedback form: http://photodemon.org/about/contact/
* Are you a classic VB coder?  I'm always open to outside bug fixes and feature implementations from fellow VB6 enthusiasts.
* Tell friends, family, and other websites about PhotoDemon.
* Send me an email and let me know what you like (or dislike) about PhotoDemon. Get in touch at http://photodemon.org/about/contact/

### How is PhotoDemon and its source code licensed?

PhotoDemon is available under a BSD license.  In a nutshell, this allows you to use its source code in any application, commercial or otherwise, provided you supply proper attribution.  Proper attribution includes a notice of copyright and disclaimer of warranty.

A full copy of the BSD license is included below.  You can also learn more about the BSD license at the following location: http://creativecommons.org/licenses/BSD/

PLEASE NOTE: sections of PhotoDemon's source code were written by third-parties and may be subject to additional copyrights and licenses.  Documentation within a specific source code file supercedes the BSD license governing this project as a whole, so please review file headers prior to using any PhotoDemon source code in your own projects.

Questions regarding licensing should be directed to: http://photodemon.org/about/contact/

Full text of BSD license follows.

Copyright (c) 2015, Tanner Helland and Contributors.
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
* Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

### Who contributes to PhotoDemon's development?

PhotoDemon would not be possible without the help of many talented contributors, including...

* Helmut Kuerbiss for maintaining (and hand-editing!) the German language file
* Roy K for additional help with the German language file
* Peter Burn for help overhauling the Shadow/Highlight adjustment tool and many other helpful suggestions
* Plinio C Garcia for editing and improving the Spanish language file, and suggesting the Color Halftone filter
* Boban Gjerasimoski for catching and reporting many bugs
(https://www.behance.net/Boban_Gjerasimoski)
* Djordje Djoric for help fixing a number of issues with shortcut keys
(https://www.odesk.com/o/profiles/users/_~0181c1599705edab79/)
* Dirk Hartmann for additional help with the German language file
(http://www.taichi-zentrum-heidelberg.de)
* Raj Chaudhuri for multiple patches (including fixes for a number of long-standing issues) and a great deal of bug-testing 
(https://github.com/rajch)
* Will Stampfer for a comprehensive code review and multiple optimization and bug-fix patches
(https://github.com/epmatsw)
* Hans Nolte for many improvements to HDR image format handling
(https://github.com/hansnolte)
* Zhu JinYong for a great deal of Unicode testing and help
(http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=66273&lngWId=1)
* Dana Seaman for many excellent Unicode references
(http://www.cyberactivex.com/) 
* Olaf Schmidt for many code samples and forum discussions, particularly regarding Unicode handling
(http://www.vbrichclient.com)
* The portablefreeware.com team for helping me debug a nasty Windows 7 crash
(http://www.portablefreeware.com/forums/viewtopic.php?t=21652)
* Frans van Beers for many detailed bug reports and testing
 (https://plus.google.com/+FransvanBeers/)
* Frank Donckers for extensive help with the translation engine. Frank also maintains the Dutch and French language files.
(http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=2213335741)
* GioRock for the Italian language file and extensive debugging (http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=77440558266)
* Audioglider for the development of many Adjustment and Effect tools, including Channel Mixer, Vibrance, Exposure, Sunshine, Bilateral Smoothing, Lens Flare, and more (https://github.com/audioglider)
* Robert Rayment for detailed research and bug-testing  (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1)
* Rod Stephens and VB-Helper.com for a themable, multiline-supporting tooltip class (http://www.vb-helper.com/howto_multi_line_tooltip.html)
* Kroc of camendesign.com for the bluDownload library and debugging contributions (http://camendesign.com)
* chrfb of deviantart.com for the original version of PhotoDemon's icon ('Ecqlipse 2,' CC-BY-NC-SA-3.0) (http://chrfb.deviantart.com/art/quot-ecqlipse-2-quot-PNG-59941546)
* Juned Chhipa for the 'jcButton 1.7' customizable command button replacement control (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1)
* Steve McMahon for an excellent CommonDialog interface, accelerator key handler, and progress bar replacement (http://www.vbaccelerator.com/home/VB/index.asp)
* Floris van de Berg and Hervé Drolon for the FreeImage library, and Carsten Klein for the VB interface (http://freeimage.sourceforge.net/)
* Jason Bullen for a native-VB implementation of knot-based cubic spline interpolation (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1)
* Dosadi for the EZTW32 scanner/digital camera library (http://eztwain.com/eztwain1.htm)
* Jean-Loup Gailly and Mark Adler for the zLib compression library (http://www.winimage.com/zLibDll/index.html)
* Waty Thierry for insights regarding printer interfacing in VB (http://www.ppreview.net/)
* Manuel Augusto Santos for the original version of the 'Artistic Contour' algorithm (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26303&lngWId=1)
* LaVolpe for his automated VB6 Manifest Creator tool (http://www.vbforums.com/showthread.php?t=606736)
* Leandro Ascierto for a clean, lightweight class that adds PNGs to menu items (http://leandroascierto.com/blog/clsmenuimage/)
* Carles P.V., Avery, and Dana Seaman for their work on GDI+ usage in VB (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1)
* Mark James and famfamfam.com for the Silk icon set (CC-BY-2.5) (http://www.famfamfam.com/lab/icons/silk/)
* Yusuke Kamiyamane for the Fugue icon set (CC-BY-3.0) (http://p.yusukekamiyamane.com/)
* Everaldo and The Crystal Project for certain menu and button icons (LGPL) (http://www.everaldo.com/crystal/)
* The Tango Icon Library for certain menu and button icons (public-domain) (http://tango.freedesktop.org/Tango_Icon_Library)
* Phil Fresle for a native-VB implementation of SHA-2 hashing (http://www.frez.co.uk/vb6.aspx)
* Pornel Lesinski, Greg Roelofs, and Jef Poskanzer for the pngquant tool (http://pngquant.org/)
* Jerry Huxtable and JHLabs for an excellent reference on Distort-style filters (Apache 2.0) (http://www.jhlabs.com/ip/filters/index.html)
* Phil Harvey for the comprehensive ExifTool metadata handler (choice of GPL or Artistic License) (http://www.sno.phy.queensu.ca/~phil/exiftool/)
* Bernhard Stockmann for his many excellent GIMP tutorials (http://www.gimpusers.com/tutorials/colorful-light-particle-stream-splash-screen-gimp.html)
* Paul Bourke for references on miscellaneous image distortions (http://paulbourke.net/miscellaneous/)
* vbForums.com user dilettante for an asynchronous piping custom control and lightweight binary stream class (http://www.vbforums.com/showthread.php?660014-VB6-ShellPipe-quot-Shell-with-I-O-Redirection-quot-control)
* Tom Loos for additional Windows 8/8.1 testing
(http://www.designedbyinstinct.com)
* All those who have contributed patches, bug reports, and donations, with extra special thanks to: Mohammad Reza, A.G. Violette, Abhijit Mhapsekar, Allan Lima, Andrew Yeoman, Dave Jamison, Alfred Hellmueller.
