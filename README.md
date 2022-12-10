## Download

| Stable (9.0) | Nightly (9.2-a) | Source code |
| :----------: | :-------------: | :---------: |
| [Download ZIP (14 MB)](https://github.com/tannerhelland/PhotoDemon/releases/download/v9.0/PhotoDemon-9.0.zip) | [Download ZIP (14 MB)](https://github.com/tannerhelland/PhotoDemon/releases/download/PhotoDemon-nightly/PhotoDemon-nightly.zip) | [Download ZIP (17 MB)](https://github.com/tannerhelland/PhotoDemon/archive/main.zip) |

## About PhotoDemon

PhotoDemon is a portable photo editor.  It is 100% free and [100% open-source](https://github.com/tannerhelland/PhotoDemon/blob/main/README.md#licensing).  

1. [Overview](#overview)
2. [What makes PhotoDemon unique?](#what-makes-photodemon-unique)
3. [What's new in nightly builds](#whats-new-in-nightly-builds)
4. [Contributing](#contributing)
5. [Licensing](#licensing)

## Overview

![Screenshot](https://photodemon.org/media/images/photodemon_9.0.png)

PhotoDemon provides a comprehensive photo editor in a 14 MB download.  It runs on any Windows PC (XP through Win 11) and it *does not* require installation.  You can run it from a USB stick, SD card, or portable drive.

PhotoDemon is open-source and available under a permissive [BSD license](#licensing).  Contributors have translated the program into more than a dozen languages.

You can support PhotoDemon's ongoing development [through Patreon](https://www.patreon.com/photodemon) or [with a one-time donation](https://photodemon.org/donate/).

New contributions from translators, coders, designers, and enthusiasts are always welcome.

* For information on the latest stable release, visit https://photodemon.org
* To download a nightly build (built from the latest source code), visit https://photodemon.org/download/
* To download PhotoDemon's source code, visit https://github.com/tannerhelland/PhotoDemon

## What makes PhotoDemon unique?

### Lightweight and completely portable
No installer is provided or required.  Aside from a temporary folder – which you can specify in the `Tools > Options` menu – PhotoDemon leaves no trace on your hard drive.  Many users run PhotoDemon from a USB stick or microSD card.

### Integrated macro recording and batch processing
Complex editing actions can be recorded as macros (similar to Office software).  A built-in batch processor lets you apply macros to entire folders of images.

### Usability is paramount
Many open-source photo editors are usability nightmares.  PhotoDemon tries not to be.  Small touches like real-time effect previews, save/load presets on all tools, unlimited Undo/Redo, "Fade last action", keyboard accelerators, mouse wheel and X-button support, and descriptive icons make it fast and easy to use.

### Pro-grade features and tools
* Extensive file format support, including Adobe Photoshop (PSD), Corel PaintShop Pro (PSP), GIMP (XCF), and all major camera RAW formats
* Advanced multi-layer support, including editable text layers and non-destructive layer modifications 
* Color-managed workflow, including full support for embedded ICC profiles
* On-canvas tools: digital paintbrushes, clone and pattern brushes, advanced selection tools, interactive gradients, and more
* Adjustment tools: levels, curves, HDR, shadow/highlight recovery, white balance, and many more
* Filters and effects: perspective correction, edge enhancement, noise removal, content-aware fill and resize, unsharp masking, gradient and palette mapping, and many more
* More than 200 tools are provided in the current build.

### Limitations

* PhotoDemon isn't designed for operating systems other than Microsoft Windows.  A compatibility layer like [Wine](http://www.winehq.org/) may allow it to work on macOS, Linux, or BSD systems, but these configurations are not officially supported.
* Due to its portable nature, PhotoDemon is only available as a 32-bit application.  (This means it cannot load or save images larger than ~2 GB in size.)

## What's new in nightly builds

![GitHub last commit](https://img.shields.io/github/last-commit/tannerhelland/PhotoDemon?style=flat-square)  ![GitHub commits since latest release](https://img.shields.io/github/commits-since/tannerhelland/PhotoDemon/latest?style=flat-square&color=light-green)

[Current nightly builds](https://photodemon.org/download/) offer the following improvements over the [last stable release](https://photodemon.org/2022/09/08/photodemon-9-0.html).

### File formats

- Comprehensive import and export support for [JPEG XL images](https://en.wikipedia.org/wiki/JPEG_XL), including full support for all color models in both lossy and lossless modes.
- Import support for [satellite topography (HGT) images](https://www2.jpl.nasa.gov/srtm/faq.html#data)
- Icon (ICO) export now provides [much higher-quality downsampling](https://github.com/tannerhelland/PhotoDemon/commit/6c3dc5ae7b33791d3cb2c7611409679f3a4c3e40) and a new `use merged image` option allows you to automatically generate icon frames from a merged multi-layer image.
- Windows metafiles (EMF, WMF) now provide an import dialog where you can choose custom rasterization dimensions.

### Image and Layer tools

- [The Advanced Text Tool supports new features](https://github.com/tannerhelland/PhotoDemon/pull/431), including justified text alignment, custom fill + stroke order, and new antialiasing settings.
- The `[Effects > Transform > Perspective]` tool now supports [custom forshortening values in both x- and y-directions](https://github.com/tannerhelland/PhotoDemon/commit/27f6d12242fad25e14b0226831d88fdd4ee7dc31).

### User interface 

- [Improved Windows XP support](https://github.com/tannerhelland/PhotoDemon/commit/8b339413e4604a568c829df9f42e52aacd786d51), including better coverage of 3rd-party libraries that have XP-specific limitations.

### Batch processing

- Batch conversion of [SVG images to raster formats](https://github.com/tannerhelland/PhotoDemon/commit/13c466f1aaef58afe623a56f47da6b3975541329) is now supported.
- Batch conversion of [Windows metafiles (EMF, WMF)](https://github.com/tannerhelland/PhotoDemon/commit/18812e6ef7d552b3da6ce430cbb0613316e8e63e) is now supported.
- A new "import size override" allows you to specify custom dimensions for vector images (SVG, EMF, WMF) involved in a batch process

### Other

For a full list of changes, [visit the project's commit log](https://github.com/tannerhelland/PhotoDemon/commits/main).

## Contributing

Ongoing PhotoDemon development is made possible by donations from users.

My [Patreon campaign](https://www.patreon.com/photodemon) is one way to donate. Donating through Patreon comes with extra benefits, like in-depth updates on new PhotoDemon features. To learn more, visit [PhotoDemon’s Patreon page](https://www.patreon.com/photodemon).

I am also extremely grateful for one-time donations.  A secure donation page is available at [photodemon.org/donate](https://photodemon.org/donate/).  **Thank you!**

If you can contribute in other ways (language translations, bug reports, pull requests, etc), please [create a new issue at GitHub](https://github.com/tannerhelland/PhotoDemon/issues).  A full list of (wonderful!) contributors is available in [AUTHORS.md](https://github.com/tannerhelland/PhotoDemon/blob/main/AUTHORS.md).

## Licensing

PhotoDemon is BSD-licensed.  This allows you to use its source code in any application, commercial or otherwise, if you supply proper attribution.  Proper attribution includes a **notice of copyright** and **disclaimer of warranty**.

PhotoDemon uses some 3rd-party open-source libraries.  These libraries are found in the /App/PhotoDemon/Plugins folder.  These libraries have their own licenses, separate from PhotoDemon.

Full licensing details are available in [LICENSE.md](https://github.com/tannerhelland/PhotoDemon/blob/main/LICENSE.md).
