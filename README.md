# PhotoDemon v6.0 alpha

### A free, portable photo editor focused on performance and usability. Macro recording and batch processing are key features. English, Dutch, French, German, and Italian translations included.

### For more info, visit http://tannerhelland.com/photodemon

***

## What makes PhotoDemon preferable to other photo editing tools?

### Portable and lightweight
PhotoDemon is designed to be run as a standalone .exe. No installer is provided or required. PhotoDemon does not touch the Windows registry, and aside from a temporary file folder – which you can specify in the Options dialog – it leaves no trace of itself on your hard drive.

### Powerful macro and batch processor support
PhotoDemon provides full macro support. Simply click “Record Macro”, then perform as many actions as you’d like. When finished, save that macro to the hard drive (in human-readable XML format) so you can repeat it at any point in the future. Macros fully integrate with a built-in batch processing tool – simply choose a saved macro and a folder or list of images, and PhotoDemon will apply that macro to every image automagically.

### Emphasis on usability
Most free, open-source image editors are usability nightmares. PhotoDemon tries not to be. The interface was built with input from professional designers – not just software engineers – and small touches like unlimited Undo/Redo, "Fade last effect", keyboard accelerators, effect previews, mouse wheel and forward-back button support, and descriptive menu icons make PhotoDemon useful to novices and professionals alike.

### A comprehensive collection of image editing tools and filters
* Powerful selection tools, with full support for antialiasing, feathering, and on-canvas sizing/moving.
* 2D transformations: image resize, crop, autocrop, rotate, shear, isometric conversion.
* Color tools: levels, curves, white balance, shadow/highlight correction, grayscale, sepia, full-featured histogram (including equalization and stretching), green screen, Wratten filters, and many more.
* Filters and effects: context-aware blur, unsharp masking, edge detection, noise removal, lens diffraction, vignetting, perspective correction, sphere wrapping, film grain, and many more.
* 100+ tools are provided in the current build, plus a custom filter tool that allows you to construct your own convolution filters.

### What doesn't PhotoDemon do?

* PhotoDemon does not provide any on-canvas painting tools. These are on the roadmap, but they did not make the cut for this version.
* PhotoDemon does not provide advanced color management (ICC profiles). More specifically, it ignores embedded ICC profiles. Even MORE specifically, PhotoDemon relies on DIB sections via the Windows GDI, which default to the sRGB space - see http://technet.microsoft.com/en-us/query/ms536845
* PhotoDemon (probably) does not run on non-Windows operating systems. Wine (http://www.winehq.org/) finally added full DIB support in March 2012. Because PhotoDemon relies heavily on DIB sections, it may work on OSX, Linux, BSD, Solaris or Maemo systems with Wine v1.4 or later. However, should you choose to go down this route, you are effectively on your own.

## PhotoDemon FAQ

### How can I get involved? 
PhotoDemon is written and maintained by a single individual with a family to support.  The program is provided free-of-charge under an extremely permissive open-source license, and no fees or money will ever be charged for its use.

That said, donations go a long way toward supporting the development of this powerful photo editing tool. If you would like to donate and support development, please visit:

http://www.tannerhelland.com/donate/

While I can't make any promises, I have been known to give extra attention to feature requests from individuals who donate. 

If you can't contribute monetarily to the project, here are other ways to help:
* Let me know if you find any bugs. Issues can be submitted via PhotoDemon's github page: https://github.com/tannerhelland/PhotoDemon, or this dedicated PhotoDemon feedback form: http://www.tannerhelland.com/photodemon-contact/
* Are you a VB6 coder? I'm always open to outside bug fixes and feature implementations from fellow programmers.
* Tell friends, family, and other websites about PhotoDemon. If you know a site that tests or reviews image processing tools, email and ask if they've tried it.
* Send me an email and let me know how you use PhotoDemon. I love to hear from users. Get in touch at http://tannerhelland.com/contact

### How is PhotoDemon and its source code licensed?

PhotoDemon is released under a BSD license. You may read more about this license at the following location: http://creativecommons.org/licenses/BSD/. A full copy of this license is included at the bottom of this section.

Sections of this source code were written by third-parties and may be subject to additional licenses. Documentation within a specific source code file supercedes the BSD license governing this project as a whole.

Questions regarding licensing should be directed to: http://tannerhelland.com/contact

Full text of BSD license follows.

Copyright (c) 2013, Tanner Helland.
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
* Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

### Who has contributed to PhotoDemon's development?

PhotoDemon would not be possible without the following talented contributors.  Thank you to...
* Frank Donckers for extensive help with the translation engine. Frank also created the Dutch, German, and French language files (30,000 words in total!)
* GioRock for the Italian language file and additional translation debugging (http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=77440558266)
* audioglider for the Channel Mixer and Vibrance tools (https://github.com/audioglider)
* Robert Rayment for detailed research and bug-testing on a variety of features (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1)
* Rod Stephens and VB-Helper.com for a themable, multiline-supporting tooltip class (http://www.vb-helper.com/howto_multi_line_tooltip.html)
* Kroc of camendesign.com for the bluMouseEvents library (http://camendesign.com)
* chrfb of deviantart.com for PhotoDemon's icon ('Ecqlipse 2,' CC-BY-NC-SA-3.0) (http://chrfb.deviantart.com/art/quot-ecqlipse-2-quot-PNG-59941546)
* Juned Chhipa for the 'jcButton 1.7' customizable command button replacement control (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1)
* Steve McMahon for an excellent CommonDialog interface, accelerator key handler, and progress bar replacement (http://www.vbaccelerator.com/home/VB/index.asp)
* Floris van de Berg and Hervé Drolon for the FreeImage library, and Carsten Klein for the VB interface (http://freeimage.sourceforge.net/)
* Jason Bullen for a native-VB implementation of knot-based cubic spline interpolation (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1)
* Dosadi for the EZTW32 scanner/digital camera library (http://eztwain.com/eztwain1.htm)
* Jean-Loup Gailly and Mark Adler for the zLib compression library (http://www.winimage.com/zLibDll/index.html)
* Waty Thierry for many insights regarding printer interfacing in VB (http://www.ppreview.net/)
* Manuel Augusto Santos for the original version of the 'Artistic Contour' algorithm (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26303&lngWId=1)
* LaVolpe for his automated VB6 Manifest Creator tool (http://www.vbforums.com/showthread.php?t=606736)
* Leandro Ascierto for a clean, lightweight class that adds PNGs to menu items (http://leandroascierto.com/blog/clsmenuimage/)
* Carles P.V., Avery, and Dana Seaman for their work on GDI+ usage in VB (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1)
* Mark James and famfamfam.com for the Silk icon set (CC-BY-2.5) (http://www.famfamfam.com/lab/icons/silk/)
* Yusuke Kamiyamane for the Fugue icon set (CC-BY-3.0) (http://p.yusukekamiyamane.com/)
* Everaldo and The Crystal Project for menu and button icons (LGPL) (http://www.everaldo.com/crystal/)
* The Tango Icon Library for menu and button icons (public-domain) (http://tango.freedesktop.org/Tango_Icon_Library)
* Phil Fresle for a native-VB implementation of SHA-2 hashing (http://www.frez.co.uk/vb6.aspx)
* Adrian Pellas-Rice, Kornel Lesinski, Stuart Coyle, Greg Roelofs, and Anthony Dekker for the pngnq-s9 tool (http://sourceforge.net/projects/pngnqs9/)
* Jerry Huxtable and JHLabs for an excellent reference on Distort-style filters (Apache 2.0) (http://www.jhlabs.com/ip/filters/index.html)
* Phil Harvey for the comprehensive ExifTool metadata handler (choice of GPL or Artistic License) (http://www.sno.phy.queensu.ca/~phil/exiftool/)
* Bernhard Stockmann for his many excellent GIMP tutorials (http://www.gimpusers.com/tutorials/colorful-light-particle-stream-splash-screen-gimp.html)
* Paul Bourke for references on miscellaneous image distortions (http://paulbourke.net/miscellaneous/)
* All those who have contributed patches, bug reports, and donations, with extra special thanks to: Allan Lima, Zhu JinYong, Andrew Yeoman, Dave Jamison, Alfred Hellmueller.