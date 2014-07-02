
http://pngquant.org

##Usage

- batch conversion of multiple files: `pngquant 256 *.png`
- Unix-style stdin/stdout chaining: `… | pngquant 16 | …`

To further reduce file size, you may want to consider [optipng](http://optipng.sourceforge.net) or [ImageOptim](http://imageoptim.pornel.net).

###BATCH FILES
Originally produced by Thomas Rutter, now with minor amendments by BJ.

You can drag and drop your 24-bit PNGs onto these batch files as an easy way
to process them without messing around with the command line. Dithered and
non-dithered optimised copies of the source PNGs will be created in the same
directory as the originals. The batch files must remain in the same directory
as pngquant.exe.

##Options

See `pngquant -h` for full list.

###`--quality min-max`

`min` and `max` are numbers in range 0 (worst) to 100 (perfect), similar to JPEG. pngquant will use the least amount of colors required to meet or exceed the `max` quality. If conversion results in quality below the `min` quality the image won't be saved (if outputting to stdin, 24-bit original will be output) and pngquant will exit with status code 99.

    pngquant --quality=65-80 image.png

###`--ext new.png`

Set custom extension (suffix) for output filename. By default `-or8.png` or `-fs8.png` is used. If you use `-ext .png -force` options pngquant will overwrite input files in place (use with caution).

###`--speed N`

Speed/quality trade-off from 1 (brute-force) to 10 (fastest). The default is 3. Speed 10 has 5% lower quality, but is 8 times faster than the default.

###`--iebug`

Workaround for IE6, which only displays fully opaque pixels. pngquant will make almost-opaque pixels fully opaque and will avoid creating new transparent colors.

###`--version`

Print version information to stdout.

###`-`

Read image from stdin and send result to stdout.

###`--`

Stops processing of arguments. This allows use of file names that start with `-`. If you're using pngquant in a script, it's advisable to put this before file names:

    pngquant $OPTIONS -- "$FILE"


#COPYRIGHT AND LICENSES

Improved PNGQuant is
- Copyright (C) 1989, 1991 by Jef Poskanzer
- Copyright (C) 1997, 2000, 2002 by Greg Roelofs; based on an idea by Stefan Schneider.
- Copyright (C) 2009-2012 by Kornel Lesinski.
** Permission to use, copy, modify, and distribute this software and its
** documentation for any purpose and without fee is hereby granted, provided
** that the above copyright notice appear in all copies and that both that
** copyright notice and this permission notice appear in supporting
** documentation.  This software is provided "as is" without express or
** implied warranty.

libpng:
* Copyright (c) 1998-2009 Glenn Randers-Pehrson
* (Version 0.96 Copyright (c) 1996, 1997 Andreas Dilger)
* (Version 0.88 Copyright (c) 1995, 1996 Guy Eric Schalnat, Group 42, Inc.)
- Full license is supplied in png.h file, but here is an excerpt:
* The PNG Reference Library is supplied "AS IS".  The Contributing Authors
* and Group 42, Inc. disclaim all warranties, expressed or implied,
* including, without limitation, the warranties of merchantability and of
* fitness for any purpose.  The Contributing Authors and Group 42, Inc.
* assume no liability for direct, indirect, incidental, special, exemplary,
* or consequential damages, which may result from the use of the PNG
* Reference Library, even if advised of the possibility of such damage.
*
* Permission is hereby granted to use, copy, modify, and distribute this
* source code, or portions hereof, for any purpose, without fee, subject
* to the following restrictions:
*
* 1. The origin of this source code must not be misrepresented.
*
* 2. Altered versions must be plainly marked as such and
* must not be misrepresented as being the original source.
*
* 3. This Copyright notice may not be removed or altered from
*    any source or altered source distribution.
*
* The Contributing Authors and Group 42, Inc. specifically permit, without
* fee, and encourage the use of this source code as a component to
* supporting the PNG file format in commercial products.  If you use this
* source code in a product, acknowledgment is not required but would be
* appreciated.

zlib:
- Copyright (C) 1995-2005 Jean-loup Gailly and Mark Adler
* This software is provided 'as-is', without any express or implied
* warranty.  In no event will the authors be held liable for any damages
* arising from the use of this software.
*
* Permission is granted to anyone to use this software for any purpose,
* including commercial applications, and to alter it and redistribute it
* freely, subject to the following restrictions:
*
* 1. The origin of this software must not be misrepresented; you must not
* claim that you wrote the original software. If you use this software
* in a product, an acknowledgment in the product documentation would be
* appreciated but is not required.
* 2. Altered source versions must be plainly marked as such, and must not be
* misrepresented as being the original software.
* 3. This notice may not be removed or altered from any source distribution.
