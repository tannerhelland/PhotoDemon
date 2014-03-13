PhotoDemon tool windows follow a standard pattern:

1. [Preview area](#preview)
2. [Tool options](#options)
3. [Command bar](#commandbar)

(image of tool window, with regions marked as 1 - Preview area, 2 - Tool options, 3 - Command bar)

### <a name="preview"></a>1: Preview area

The preview area of a tool displays a live sample of how the image would be affected by the current tool options.  

All previews provide a link for toggling between the original image, and a preview of the tool effect.  This can be helpful for seeing subtle differences caused by the current tool.

Some tools may provide additional preview options:

- A **zoom button** in the bottom-right.  
    This button allows you to switch between:
	- A preview of the full image, shrunk to fit the window
	- A section of the window at 100% zoom.  Because most images do not fit in the preview area when viewed at 100%, you are allowed to click-drag the image to move it around.
	
	Some tools do not allow you to switch out of *fit full image* mode.  This is because that tool only works if it has access to the entire image (*e.g.* the "rotate image" tool).
	
- Set tool options by **clicking the preview window**.
    This feature will vary from tool to tool, but common features include:
	
	- **Click to set center point**.  Some tools, like the Vignetting style effect, operate around a "center point", which you can set by clicking anywhere on the preview window.
	- **Click to select color**.  Some tools, like the Green screen tool, allow you to specify a target color.  In these tools, you can select the target color by clicking the preview image.  (If this mode is available, hovering the mouse over the image will cause the preview effect to temporarily disappear, so you can select a color from the original image.)  
	- Other options may be available in certain tool dialogs; refer to individual tool documentation for details.
	
### <a name="options"></a>2: Tool options	
	
Each tool window will provide one or more tool options.  Options come in many shapes and sizes, including scroll bars, color selectors, drop-down lists, and more. 

The first time you use a tool, all options will be set to a predetermined default value.  After you have used a tool, options will default to your last-used settings.

For information on a specific tool's options and what they control, visit the help page for that specific tool.

### <a name="commandbar"></a>3: Command Bar

All tool dialogs provide an advanced **command bar** at the bottom of the screen.  In addition to standard OK and Cancel buttons, the command bar provides a few extra features.  Starting from the left, these features are:

(screenshot of command bar, with buttons labeled from left-to-right)

1. [Reset](#reset)
2. [Randomize](#randomize)
3. [Preset List](#presets)
4. [Save preset](#savepreset)

#### <a name="reset"></a>1: Reset

After playing with a tool's options, it may be desirable to reset them to their default values.  Use this button to reset all options to their default settings.

Because PhotoDemon automatically remembers your last-used settings for all tools, the Reset option can also be helpful for ignoring your last-used settings in favor of the default tool parameters.

(Note: the Reset button will not modify any saved preset values.)

#### <a name="randomize"></a>2: Randomize

If a tool has many options, it may be difficult to know how the various options interact.  The Randomize button can help.  Each time it is clicked, it will set all options on the form to a random value.  By clicking it repeatedly, you can see how the tool works under many different settings.

After randomizing values one or more times, you can use the Reset button (the left-most button) to return the tool to its default state. 

#### <a name="presets"></a>3, <a name="savepreset"></a>4: Preset list and Save preset

Sometimes you may find that you frequently use the same set of values for a given tool.  As an example, on the Resize Tool, perhaps you frequently resize images to 1920x1080, because that is the resolution of your monitor.

To save you the trouble of constantly re-entering frequently used values, you can save them as a Preset.  After setting all tool options to the values you desire, enter a name for the preset in the Preset box (such as "1920x1080 wallpaper" for the example above), then click the "Save Preset" button.  The current values will then be available at any point in the future by clicking that entry in the Preset drop-down box.

By default, PhotoDemon will always save one preset for you, called "last-used settings".  This preset will automatically appear after you use a tool for the first time, and it will always be updated with your most recently used values for that tool.  *Do not save your own preset under this name*, as it will be replaced with your actual last-used settings as soon as the dialog is closed.


* * *

###Technical Notes

- Preset data is saved to the Data/Presets subfolder.  
- XML format is used for all presets.  
- Each dialog has its own preset file, which is named using the internal PhotoDemon name for the dialog (computed dynamically at run-time).  
- If the same dialog is used for multiple tools (for example, the Median dialog is used for Median, Erode, and Dilate), extra identifiers are automatically added to the XML filename.
- You can manually edit the XML contents of Preset files.  This can be helpful when converting preset data from other software to PhotoDemon format.  Just make sure to close all tags properly, and to use proper data types to avoid conversion errors at run-time.  For numeric preferences in particular, you must exercise caution, as invalid values may cause out-of-range or divide-by-zero errors (among other problems).  Do not assume that a given tool will automatically clip input values for you!
- Preset files' XML tags should be self-explanatory.  Each control on a dialog is treated as a tag, while its value is written between tags in an appropriate format.  Some special controls - like the Curves dialog or PhotoDemon's internal Resize control - are hooked using special command bar functions.  These tools write their data to file using PD's parameter string format, where each unique parameter is delimited by the | character.
- There is no limit to the number of presets that can be saved.  
- At present, there is no way to delete saved presets, short of editing preset XML files by hand.  I hope to remedy this in the future.