# PhotoDemon 7.2 alpha

### PhotoDemon is a portable photo editor focused on performance and usability. 

1. [Overview](#overview)
2. [What makes PhotoDemon unique?](#what-makes-photodemon-unique)
3. [What's new in nightly builds](#whats-new-in-nightly-builds)
4. [Contributing](#contributing)
5. [Licensing](#licensing)

## Overview

![Screenshot](https://photodemon.org/media/PD_screenshot_master.jpg)

PhotoDemon provides a comprehensive photo editor in a 14 MB download.  It runs on any Windows machine (XP through Win 10) and it *does not* require installation.  Feel free to run it from a USB stick, SD card, or portable drive.

PhotoDemon is open-source and available under a permissive [BSD license](#licensing).  Contributors have translated the program into more than twenty languages.

You can support PhotoDemon's ongoing development [through Patreon](https://www.patreon.com/photodemon) or [with a one-time donation](https://photodemon.org/donate/).

New contributions from coders, designers, translators, and enthusiasts are always welcome.

* For information on the latest stable release, visit https://photodemon.org
* To download the program's source code, visit https://github.com/tannerhelland/PhotoDemon
* To download a nightly build of the latest source code, visit https://photodemon.org/download/

## What makes PhotoDemon unique?

### Lightweight and completely portable
No installer is provided or required.  Aside from a temporary folder – which you can specify from the Tools > Options menu – PhotoDemon leaves no trace on your hard drive.  Many users run PhotoDemon from a USB stick or portable drive.

### Integrated macro recording and batch processing
Complex editing actions can be recorded as macros (similar to Office software).  A built-in batch processor lets you apply macros to entire folders of images.

### Usability is paramount
Many open-source photo editors are usability nightmares.  PhotoDemon tries not to be.  Small touches like save/load presets on all tools, unlimited Undo/Redo, "Fade last action", keyboard accelerators, effect previews, mouse wheel and X-button support, and descriptive menu icons make it fast and easy to use.

### Pro-grade features and tools
* Extensive file format support, including Adobe Photoshop files (PSD) and all major camera RAW formats
* Color-managed workflow, including full support for embedded ICC profiles
* Non-destructive editing for select features, including resizing, rotation, and common adjustments
* 2D transformations: advanced rescale operators (Sinc, Catmull-Rom, etc), content-aware scaling (seam carving), crop, straighten, shear, zoom
* Adjustment tools: levels, curves, HDR, shadow/highlight recovery, white balance, Wratten filters, and many more
* Filters and effects: perspective correction, edge detection, noise removal, content-aware blur, unsharp masking, green screen, lens diffraction, vignetting, and many more
* More than 200 tools are provided in the current build.

### Limitations

* PhotoDemon isn't designed for operating systems other than Microsoft Windows.  A compatibility layer like Wine (http://www.winehq.org/) may allow it to work on OSX, Linux, or BSD systems, but program stability and performance may suffer.

## What's new in nightly builds

[Current nightly builds](https://photodemon.org/download/) offer the following improvements over the [last stable release](https://photodemon.org/2017/11/28/photodemon-7-0-release.html):

- Comprehensive support for [Adobe Photoshop (PSD) files](https://photodemon.org/2019/02/20/psd-support-now-available.html) and their open-source equivalent, [OpenRaster (ORA) files](https://www.openraster.org/)
- Comprehensive support for [animated PNG and GIF files](https://github.com/tannerhelland/PhotoDemon/issues/278)
- New [custom-built PNG engine](https://github.com/tannerhelland/PhotoDemon/commit/8206ae38831bc095afa49556420bbb7d5c15778f) with a fully integrated color-mangement pipeline; the engine also [auto-optimizes PNGs losslessly](https://github.com/tannerhelland/PhotoDemon/commit/10c78b3cc12c7e99af49d1667f5d8887b99a054c) for maximum file size reductions, with additional options for lossy quantization (similar to [pngquant](https://pngquant.org/)).
- Main UI support for viewing animated images
- New best-in-class [gradient tool](https://www.patreon.com/posts/photodemons-new-26199115) and gradient editor
- New clone stamp tool
- Main UI now provides [a search bar](https://www.patreon.com/posts/photodemon-now-26904685) for locating features and tools
- Main UI now provides [on-canvas rulers](https://www.patreon.com/posts/canvas-rulers-to-19178070)
- New on-canvas [measure tool](https://www.patreon.com/posts/how-to-use-new-7-20466383) with support for auto-straightening the image
- New [Effects > Render menu](https://www.patreon.com/posts/photodemon-7-2-29679659) with Clouds and Fibers tools.
- Macros can now be [automatically created from session history](https://github.com/tannerhelland/PhotoDemon/issues/265)
- [Many](https://github.com/tannerhelland/PhotoDemon/issues/244) [improvements](https://github.com/tannerhelland/PhotoDemon/issues/243) to [keyboard](https://github.com/tannerhelland/PhotoDemon/commit/730f2ebe7a8121e7c5c633ce7b3ff7aea01dc273) [navigation](https://github.com/tannerhelland/PhotoDemon/issues/277)
- All-new [digital palette features](https://www.patreon.com/posts/how-to-use-new-7-19823148), including import/export support for all major palette file formats (Adobe PhotoShop, PaintShop Pro, Paint.NET, GIMP, JASC)
- Color selector now provides a [palette UI mode](https://github.com/tannerhelland/PhotoDemon/commit/904b1c6d5b72a9e4488648f50bcebe6bb51a2080) for selecting colors from a palette file
- Right-side UI panels are now user-resizable
- Improved [auto-correct and auto-enhance tools](https://github.com/tannerhelland/PhotoDemon/commit/1800489ce2f59277833b2eebd5319139ab7050cc)
- Input boxes now support [simple math equations](https://github.com/tannerhelland/PhotoDemon/issues/263)
- Selection tools now support [locked aspect ratios](https://github.com/tannerhelland/PhotoDemon/commit/d263e3bb3777db27ae1953bd15b25d299b96fc08) for easier cropping
- New Layer > Split menu for automatically splitting layers into individual images, or merging separate images into a single layered image
- Disk I/O tasks have been moved to a new [memory-mapped interface](https://en.wikipedia.org/wiki/Memory-mapped_file), improving performance when e.g. loading/saving image files
- Saved presets on all tools can now be edited, deleted, and rearranged from within the tool UI
- Users on stable builds can now invoke PD's internal debug tool, and the debugger will also auto-start (if user preferences allow) when it detects a program crash
- Numerous bug-fixes, memory reductions, and performance improvements

For a full list of changes, please consult [the commit log](https://github.com/tannerhelland/PhotoDemon/commits/master).

## Contributing

PhotoDemon is primarily supported by an [ongoing Patreon campaign](https://www.patreon.com/photodemon). Donating through Patreon comes with extra benefits, like monthly tutorials and updates on new PhotoDemon features, and an interactive area where you can submit feature requests. To learn more, visit [PhotoDemon’s Patreon page](https://www.patreon.com/photodemon):

PhotoDemon's lone developer is also extremely grateful for one-time donations.  A secure donation page is available at [photodemon.org/donate](https://photodemon.org/donate/).  Thank you!

If you are interested in contributing in other ways (language translations, bug reports, pull requests, etc), please [create a new issue at GitHub](https://github.com/tannerhelland/PhotoDemon/issues).  A full list of (wonderful!) contributors is available in [AUTHORS.md](https://github.com/tannerhelland/PhotoDemon/blob/master/AUTHORS.md).

## Licensing

PhotoDemon is BSD-licensed.  This allows you to use its source code in any application, commercial or otherwise, if you supply proper attribution.  Proper attribution includes **a notice of copyright** and **disclaimer of warranty**.

PhotoDemon uses some 3rd-party open-source libraries.  These libraries are found in the /App/PhotoDemon/Plugins folder.  These libraries have their own licenses, separate from PhotoDemon.

Full licensing details are available in [LICENSE.md](https://github.com/tannerhelland/PhotoDemon/blob/master/LICENSE.md).
