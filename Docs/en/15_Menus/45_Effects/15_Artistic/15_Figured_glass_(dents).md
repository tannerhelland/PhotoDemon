The "Figured glass" effect projects an image through an imaginary sheet of [figured glass](http://en.wikipedia.org/wiki/Architectural_glass#Rolled_plate_.28figured.29_glass).

![figured glass dialog](/docimg/figured_glass.jpg)

For information about standard parts of the tool interface (such as the preview area on the left, or the colored bar with buttons along the bottom), please refer to the [general tool window article](/en/Interface_and_Toolboxes/Individual_Tool_Windows).

### Options

- **Scale** 
    The size of individual distortions, on a scale of 0-100 percent.
- **Turbulence** 
    The depth of individual distortions.
- **If pixels lie outside the image...** 
    This tool moves pixels from their original location, and sometimes they may move completely outside image boundaries.  PhotoDemon provides a number of different ways to deal with such pixels:
	- **Clamp to the nearest available pixel**
	  Use the color of the nearest valid pixel.  This may cause streaking lines of color in the final image.
	- **Reflect them across the nearest edge**
	  Take the pixel's distance from the image, and "reflect" that distance across the nearest image edge.  This causes a mirror-like effect along boundaries.
	- **Wrap them around the image**  
	  Take the pixel's distance from the image, and wrap it around the opposite side.  This causes a tile-like effect along boundaries.
	- **Erase them**
	  Force pixels to the current background color (black, by default).
	- **Ignore them**
	  Do not process pixels that fall outside the image.  This may cause portions of the original image to show through.
- **Render emphasis**
    This tool provides a choice between slow, high-quality rendering and fast, low-quality rendering.  (*Quality* is always recommended, but *Speed* may be helpful on very old PCs.)

	
* * *
	
### Developer Notes	

The "Figured glass" filter uses [Perlin Noise](http://en.wikipedia.org/wiki/Perlin_noise) to calculate a custom displacement map for the image.  PhotoDemon uses a heavily modified version of an algorithm first made available by [Jerry Huxtable](http://www.jhlabs.com/index.html), and a heavily modified version of a Perlin Noise class first made available by [Steve McMahon](http://www.vbaccelerator.com/home/VB/Code/vbMedia/Algorithmic_Images/Perlin_Noise/article.asp).

Displacement is calculated according to the formula:

        pNoiseCache = PI_DOUBLE * Perlin_Noise(x / fxScale, y / fxScale) * fxTurbulence
        
		srcX = x + Sin(pNoiseCache) * fxScale
        srcY = y + Cos(pNoiseCache) * fxScale * fxTurbulence
			
Special handling is required when Scale = 0.

A few notes:

- Using Cos() for y displacement causes the image to move vertically, even when no turbulence has been applied.  There are a few ways to avoid this.  In versions pre-6.2, PhotoDemon used Sin() for both x and y.  This had the unfortunate side-effect of stretching all pixels along the x=y axis.  In 6.2, Cos() was added back in, and the turbulence parameter is now multiplied by the displacement, which prevents vertical displacement if turbulence is zero.  (Thanks to Robert Rayment for testing and discussion on this point!)

- Sin() and Cos() are not required for the filter to work, but they add a pleasant "roundness" to the displacement.

- To simplify inputs for the user, the Scale parameter is used for many aspects of the distortion.  Technically, xScale, yScale, and Amount could all be handled separately.

- Jerry's original implementation pre-computed lookup tables for the sin/cos values.  The lookup tables held 256 entries, and Perlin Noise (multiplied by 127 and added to 127) was used to look up table values.  In 6.2, I rewrote PhotoDemon's method without lookup tables, to improve interpolation quality.  This slowed the filter slightly, but a fresh batch of profiling in the Perlin Noise class offset any speed losses.  (In fact, the new version is actually faster than the old one.)