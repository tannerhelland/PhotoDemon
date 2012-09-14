PhotoDemon v5.0b1 - Copyright (c) 2012 by Tanner Helland
=====================================================================================
A free, open-source photo editor and image processing tool.
http://tannerhelland.com/photodemon


What makes PhotoDemon preferable to other photo editing tools?
-------------------------------------------------------------------------------------
:: Portable and lightweight
	PhotoDemon is designed to be run as a standalone .exe.  No installer is provided
	or required.  An INI file is used to store program settings, and if no INI is
	found, PhotoDemon will generate one for you.  PhotoDemon does not touch the
	Windows registry, and aside from a temporary file folder – which you can specify
	in the Preferences menu – it leaves no trace of itself on your hard drive.

:: Powerful macro and batch conversion support
	PhotoDemon provides full macro support.  Simply hit “Record Macro”, then perform
	as many actions as you’d like.  When finished, save that macro to the hard drive
	so you can repeat it at any point in the future.  Macros fully integrate with a
	built-in batch conversion tool – simply choose a saved macro and a folder or list
	of images, and PhotoDemon will apply that macro to every image automagically.  For
	large batches (50+ images), PhotoDemon will give you a running estimate of
	time-to-completion.
		
:: Emphasis on usability
	Most free, open-source image editors are usability nightmares.  PhotoDemon tries
	not to be.  The interface was built with input from professional designers – not
	just software engineers – and small touches, like unlimited Undo/Redo, “Fade last
	effect,” keyboard accelerators, effect previews, mouse wheel and forward-back
	button support, and descriptive menu icons make PhotoDemon useful to novices and
	professionals alike.
	
:: A comprehensive selection of image editing tools and filters
	2D transformations: image resizing, rotation, isometric conversion.  Color tools:
	image levels, white balance, grayscale, sepia, color reduction, full-featured
	histogram (including equalization and stretching).  Filters: blur, sharpen, edge
	detection, solarize, despeckle, dilate/erode, diffuse, mosaic, and many more.
	50+ in the current build – and that’s not including a custom filter tool that
	allows you to build your own 5×5 convolution filters.

	
What doesn't PhotoDemon do?
-------------------------------------------------------------------------------------
:: Painting tools.
	PhotoDemon does not provide any painting tools.  It only supports actions and
	filters that operate on an entire image.

:: Alpha-channels (transparency) and high bit-depths
	Per its name, PhotoDemon is designed for use with photos.  It will happily import
	images with alpha channels or bit-depths greater than 16 million colors, but it
	will internally convert these images to true color (24-bit RGB) before operating
	on them, and it will only save images in non-alpha 8 or 24-bit color depths.  If
	you need alpha or deep color support, I'm afraid PhotoDemon is not the right tool
	for you.
		
:: Advanced color management (ICC profiles)
	PhotoDemon ignores embedded ICC profiles.  As a tool designed for consumers and
	hobbyists, it is unlikely to ever gain ICC profile support.  If color management
	is integral to your work, PhotoDemon is not the right tool for you.  (Note: if
	you're interested, PhotoDemon relies on DIB sections via the Windows GDI, which
	default to the sRGB space - http://technet.microsoft.com/en-us/query/ms536845)
		
:: Run on non-Windows operating systems...probably
	Wine (http://www.winehq.org/) finally added full DIB support in March 2012 (v1.4).
	Because PhotoDemon relies heavily on DIB sections, it may work on OSX, Linux, BSD,
	Solaris or Maemo systems with Wine v1.4 (or later) installed.  However, should you
	choose to go down this route, you are effectively on your own.  PhotoDemon's
	developer doesn't have the resources to support Wine in any official capacity.

		
		
Contents of this file:
=====================================================================================
[1] How to support PhotoDemon
[2] License
[3] Acknowledgements


[1] How to support PhotoDemon:
=====================================================================================

PhotoDemon is written and maintained by a single individual (with a family to support!).  
It is provided free-of-charge under an extremely permissive open-source license, and no
fees or money will ever be charged for its use.

That said, donations go a long way toward supporting the development of this powerful
image editing tool.  If you would like to donate and support PhotoDemon's development,
please visit:
http://www.tannerhelland.com/donate/

While I can't make any promises, I have been known to give extra attention to feature
requests from individuals who donate.

If you can't contribute monetarily to the project, here are other ways to help:

* Let me know if you find any bugs.  Issues can be submitted via PhotoDemon's github page:
  https://github.com/tannerhelland/PhotoDemon
  ...or this dedicated PhotoDemon feedback form:
  http://www.tannerhelland.com/photodemon-contact/
* Are you a VB6 fiend?  I'm always open to outside bug fixes and feature implementations
  from fellow VB6 programmers.
* Tell friends, family, and other websites about PhotoDemon.  If you know a site that
  tests or reviews image processing tools, email and ask if they've tried it.
* Send me an email and let me know how you use PhotoDemon.  I love to hear from users.
  Get in touch at:  tannerhelland.com/contact 

  
[2] License
=====================================================================================
PhotoDemon is Copyright (c) 2012 by Tanner Helland, www.tannerhelland.com

PhotoDemon is released under a BSD license. You may read more about this license at the following location: http://creativecommons.org/licenses/BSD/.  A full copy of this license is included at the bottom of this section.

Parts of this source code were written by third-parties and may be subject to additional licenses.  Documentation within a specific source code file supercedes the BSD license governing this project as a whole.

Questions regarding licensing should be directed to: www.tannerhelland.com/contact

Full text of BSD license follows.

Copyright (c) 2012, Tanner Helland
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

    - Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
    - Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


[3] Acknowledgements:
=====================================================================================

* Kroc of camendesign.com for many suggestions regarding UI design and organization
  (http://camendesign.com)
* chrfb of deviantart.com for PhotoDemon's icon ('Ecqlipse 2,' CC-BY-NC-SA-3.0)
  (http://chrfb.deviantart.com/art/quot-ecqlipse-2-quot-PNG-59941546)
* Juned Chhipa for the 'jcButton 1.7' customizable command button replacement control
  (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1)
* Steve McMahon for an excellent CommonDialog interface, accelerator key handler, and progress bar replacement
  (http://www.vbaccelerator.com/home/VB/index.asp)
* Floris van de Berg and Hervé Drolon for the FreeImage library, and Carsten Klein for the VB interface
  (http://freeimage.sourceforge.net/)
* Brad Martinez for the original implementation of VB binary file extraction
  (http://btmtz.mvps.org/gfxfromfrx/)
* Ron van Tilburg for a native-VB implementation of Xiaolin Wu's line antialiasing routine
  (http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=71370&lngWid=1)
* Jason Bullen for a native-VB implementation of knot-based cubic spline interpolation
  (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1)
* Paul Turcksin for code related to dynamic icon generation (used for MDI child icons in PhotoDemon)
  (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=60600&lngWId=1)
* Dosadi for the EZTW32 scanner/digital camera library
  (http://eztwain.com/eztwain1.htm)
* Jean-Loup Gailly and Mark Adler for the zLib compression library, and Gilles Vollant for the WAPI wrapper
  (http://www.winimage.com/zLibDll/index.html)
* Waty Thierry for many insights regarding printer interfacing in VB
  (http://www.ppreview.net/)
* Manuel Augusto Santos for original versions of the 'Enhanced 2-bit Color Reduction' and 'Artistic Contour' algorithms
  (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26303&lngWId=1)
* Johannes B for the original version of the 'Fog' algorithm
  (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42642&lngWId=1)
* LaVolpe for his automated VB6 Manifest Creator tool
  (http://www.vbforums.com/showthread.php?t=606736)
* Leandro Ascierto for a clean, lightweight class that adds PNGs to menu items
  (http://leandroascierto.com/blog/clsmenuimage/)
* Carles P.V., Avery, and Dana Seaman for their work on GDI+ usage in VB
  (http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1)
* Mark James and famfamfam.com for the Silk icon set (CC-BY-2.5)
  (http://www.famfamfam.com/lab/icons/silk/)
* Yusuke Kamiyamane for the Fugue icon set (CC-BY-3.0)
  http://p.yusukekamiyamane.com/
* Everaldo and The Crystal Project for menu and button icons (LGPL)
  (http://www.everaldo.com/crystal/)
* The Tango Icon Library for menu and button icons (public-domain)
  (http://tango.freedesktop.org/Tango_Icon_Library)
* All those who have contributed patches, bug reports, and donations, with extra special thanks to:
  * Herman Liu
  * Dave Jamison
  * Robert Rayment
  * Alfred Hellmueller
