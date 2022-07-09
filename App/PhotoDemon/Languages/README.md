## About PhotoDemon localization

Thank you for your interest in PhotoDemon localization.

This README provides a quick overview of PhotoDemon's localization process.  Importantly, there are **critical instructions** you must follow to avoid losing changes to these language files.

Please read this entire document before editing PhotoDemon language files.

### Quick overview

The XML files in this folder (`/App/PhotoDemon/Languages`) are PhotoDemon's "official" language files.  They ship with all PhotoDemon downloads.

When PhotoDemon self-updates, these XML files are automatically replaced with newer versions from the update server.  For nightly builds in particular, language files are frequently updated, so this replacement process could happen at any time.

Because the language files in this folder can be replaced at any time, **you MUST NOT edit them directly**.

Instead, any modified language files should be saved to the `/Data/Languages` folder.  The `/App` subfolder is reserved for PhotoDemon itself, but the `/Data` subfolder is your user data folder.  Files in the `/Data` folder will not be erased by PhotoDemon updates.

So please, please, please remember to place any modified language files in the `/Data/Languages` folder, **not** this `/App/PhotoDemon/Languages` folder.

### Editing an existing language file in any text editor

Editing existing language files is easy.

**Short version:**

PhotoDemon's language files are plain-text XML-like files.  You can edit them in any text editor.  After making changes, send the updated file to me and I will add it to PhotoDemon!

**Long version:**

1) Copy the language file you want to edit into the `/Data/Languages` folder, and modify its filename to something like `German_new.xml`.
2) Open the language file in any text editor.  For beginners, I recommend the free Notepad++ app: https://notepad-plus-plus.org/
3) Inside the language file, you will see a collection of tags.  Tags are special text enclosed by angle brackets, with names like `<langid>` or `<author>`.  Do not edit text within `<` and `>` characters.  Only translate text *between* tags - for example, if you see `<translation>text goes here</translation>` you can freely edit the `text goes here` portion.
4) Ensure the `<langid>` text at the top of the file matches the two-letter ISO language code (https://en.wikipedia.org/wiki/List_of_ISO_639-1_codes) and two-letter ISO country code (https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2).  PhotoDemon uses these codes to auto-suggest languages to new users (by matching the codes to the language and country code of a user's PC).
5) You can also modify the language version to match the PhotoDemon version you are working on, and the language status (e.g. "complete" or "in-progress") and author tags to reflect their current state.
6) All phrases in the program are stored as `<original>` and `<translation>` pairs.  You *must not* modify any text inside `<original></original>` tags.  PhotoDemon requires the `<original>` text to precisely match on-screen interface elements.
7) You can freely modify text inside `<translation>` tags.  Empty tags (`<translation></translation>`) are common for newly added features.  You can use this to quickly locate phrases that need to be translated, by searching for the text `<translation></translation>`.
8) Where possible, PhotoDemon tries to use similar naming conventions to other popular photo editors, like Adobe Photoshop and GIMP.  If a menu command in PhotoDemon does not make sense, you might see what text Photoshop or GIMP uses for their corresponding feature.  (Later on, I will describe an automated way to do this.)
9) Sometimes, PhotoDemon needs to insert dynamic text at run-time.  This is usually a number or percent, like `"Step 1 of 3"`.  Dynamic text is flagged with a `%` prefix, so the example phrase `"Step 1 of 3"` would appear as `<original>Step %1 of %2</original>` in the language file.  In your translation, place the same `%1` and `%2` markers wherever they make sense.
10) When you are satisfied with your changes, save your work and contact me.  Filing a pull request or new issue on GitHub and attaching your updated language file is the fastest way to submit changes: https://github.com/tannerhelland/PhotoDemon .  If you are uncomfortable using GitHub, alternate means of contacting me are available here: https://photodemon.org/about/ 

### Starting a new language file

Starting a new PhotoDemon language file is exactly the same as editing an existing one.  The only difference is that instead of copying an existing language file into the `/Data/Languages` folder, you will instead copy PhotoDemon's master English text file from `/App/PhotoDemon/Languages/Master/MASTER.xml` to the `/Data/Languages` folder.  Rename the file with an appropriate language name, then follow the steps given above in `Editing an existing PhotoDemon language file`.

### Using PhotoDemon's built-in Tools > Language Editor menu

Some translators prefer to use PhotoDemon's built-in `Tools > Language editor` tool.  It can simplify the language editing process, especially for beginners.

PhotoDemon's Language Editor tool is largely self-explanatory, but two items deserve extra explanation.  

#### Automatic translation suggestions (via the online DeepL service)

On the second page of the editor, there is a box titled `(optional) free DeepL.com API key for translation suggestions`.  

DeepL.com (https://www.deepl.com/translator) is a free, high-quality translation service.  DeepL is not connected to or affiliated with PhotoDemon in any way, but PhotoDemon's Language Editor can interface with DeepL's website to automatically suggest translations for you.  This is especially helpful for longer phrases, like error messages, which can be tedious to translate manually.

Like most online services, DeepL requires you to set up a free user account before using their translation service.  As of June 2022, the "Free" button at this link is the fastest way to open a new account:

https://www.deepl.com/pro-api?cta=header-pro-api/

After setting up an account, you will receive a unique API key.  Copy and paste that key into the matching box in PhotoDemon's Language Editor, and PhotoDemon will start auto-suggesting translations for you.  PhotoDemon will also save the pasted API key to your user preferences file, so you do not need to enter it again.

Again, this feature **is not required** to edit PhotoDemon translations.  It is simply provided as a convenience for those who want it.

If other automated translation services are more appropriate for your language, please contact me and I'll see if I can add them to the tool as well.

#### Comparing translations to other open-source software

Where possible, PhotoDemon tries to use the same terminology as other popular photo editors.  This reduces the learning curve for new users and makes it easier to switch between PhotoDemon and other software.

To help achieve this, the second page of PhotoDemon's Language Editor provides a box titled `(optional) 3rd-party language file (.po) to compare`.  Click the `...` button to select a language file from any other software.  For example, GIMP's language files are freely downloadable from this link:

https://gitlab.gnome.org/GNOME/gimp/-/tree/master/po 

...while Krita's language files are freely downloadable from this link:

https://websvn.kde.org/trunk/l10n-kf5/

(Click your desired language, then `messages > krita > krita.po > download`.)

If you provide a translation file from another app, PhotoDemon will also display that app's translations on the translation panel (when available).  This can be helpful for common photo editing terms like menu and tool names that appear in both programs.

While this feature is very helpful, I must request that you *do not* blindly copy text from other software's files.  These translations are typically copyright by their original authors and should be used as a reference only.  PhotoDemon's Language Editor will not allow you to bulk-copy these translations, by design.  In particular, I provide the GIMP and Krita links only as a helpful reference for challenging phrases, *not* as a way to steal anyone else's translation work.

And again, this feature is 100% optional and you do not need to use it to translate PhotoDemon text.  It only exists as a convenience for those who want it.

### Conclusion

Thank you again for helping me improve PhotoDemon's language files.  If you have any other questions, please contact me.  My contact info is available here:

http://photodemon.org/about/contact/

I am very excited to merge your localization work into the project!

Kind regards,

Tanner Helland

(PhotoDemon developer)

Last modified: 09 July 2022