## About PhotoDemon localization

Thank you for your interest in PhotoDemon's language support.

This README provides a quick overview of PhotoDemon's localization process.  Importantly, it provides critical instructions to prevent you from losing any updates you make to PhotoDemon's language files.

Please read this entire document before editing any files in this folder.

### Quick overview

The XML files in this folder (`/App/PhotoDemon/Languages`) are PhotoDemon's "official" language files.  They are included in all PhotoDemon downloads.

When PhotoDemon applies an automatic update, these files will be replaced with newer versions from the PhotoDemon update server.  For nightly builds in particular, language files are frequently updated, so this replacement process could happen at any time.

Because these language files can be replaced at any time, you **MUST NOT** edit them directly.

Instead, modified language files must be saved to the `/Data/Languages` folder.  The `/App` subfolder is reserved for PhotoDemon itself, but the `/Data` subfolder is your user data folder.  Files in the `/Data` folder will never be erased by PhotoDemon updates.

So please, please, please remember to place any modified language files in the `/Data/Languages` folder, **not** the `/App/PhotoDemon/Languages` folder.

### Editing an existing language file

Editing an existing language file is easy!  

1) Copy the language file you want to edit into the `/Data/Languages` folder, and modify its filename to something like `German_new.xml`.
2) Open the language file in any text editor.  Many translators use the free Notepad++ app: https://notepad-plus-plus.org/
3) Inside the language file, you will notice a few tags at the top of the file.  These tags have names like `<langid>` or `<author>`.  
4a) Ensure the `<langid>` matches the two-letter ISO language code (https://en.wikipedia.org/wiki/List_of_ISO_639-1_codes) and two-letter ISO country code (https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2) that you want to provide.  PhotoDemon uses these codes to auto-suggest languages to new users, by matching these codes to the language and country code of each user's PC.  
4b) (Optional)  You can also modify the language version (to ideally match the PhotoDemon version you are working on), language status (e.g. "complete" or "in-progress"), and author tags, as relevant.
5) All phrases in the program are stored as `<original>` and `<translation>` pairs.  You *must not* modify any text inside `<original></original>` tags.  PhotoDemon requires these tags to precisely match on-screen interface elements - that's how it locates specific translations inside each file.
6) You can freely modify any text between `<translation>` and `</translation>` tags.  Empty tags (e.g. translation tags with no text between them) have not yet been translated.  This is common for newly added features.  You can use this to quickly locate text that needs to be translated, by searching for `<translation></translation>`.
7) Where possible, PhotoDemon tries to match naming conventions with other popular photo editors, like Adobe Photoshop or GIMP.  If a menu command in PhotoDemon does not make sense, you might check what text GIMP uses for their corresponding feature: https://gitlab.gnome.org/GNOME/gimp/-/tree/master/po (click on any language to see the GIMP translations for that language).  I must also request that you *do not* blindly copy text from other software's translations.  I provide the GIMP link only as a reference for comparison on challenging phrases, *not* as a way to steal their translation work.
8) Sometimes, PhotoDemon needs to insert dynamic text at run-time.  This is usually a number or percent, like `"Step 1 of 3"`.  Dynamic text is flagged in translations using a % prefix, so the example phrase `"Step 1 of 3"` would appear as `<original>Step %1 of %2</original>` in a PhotoDemon language file.  In your translation, place the same `%1` and `%2` markers in your translation wherever they make sense.
9) When you are satisfied with your changes, save your work and contact me.  Filing a pull request or new issue on GitHub and attaching your updated file is the fastest way to get your changes merged into the main application: https://github.com/tannerhelland/PhotoDemon .  If you are afraid of GitHub, alternate means of contacting me are available here: https://photodemon.org/about/ 

### Starting a new language file

Starting a new PhotoDemon language file is exactly the same as editing an existing one.  The only difference is that instead of copying an existing language file into the `/Data/Languages` folder, you will instead copy PhotoDemon's master English text file from `/App/PhotoDemon/Languages/Master/MASTER.xml` into the `/Data/Languages` folder.  Rename the file with an appropriate language name, then follow the steps given in `Editing an existing PhotoDemon language file`, above.

### PhotoDemon's built-in Tools > Language Editor menu

Some translators prefer to use PhotoDemon's built-in `Tools > Language editor` tool.  It can simplify the language editing process, especially for beginners.

PhotoDemon's Language Editor tool is largely self-explanatory, but one item deserves extra explanation.  On the first page of the editor, there is a box titled `(optional) free DeepL.com API key for translation suggestions`.  

DeepL.com (https://www.deepl.com/translator) is a free, high-quality translation service.  It is not connected to or affiliated with PhotoDemon in any way, but PhotoDemon's Language Editor can interface with DeepL.com to automatically suggest translations for you.  This is especially helpful for longer phrases, like error messages, which can be tedious to translate manually.

Like most online services, DeepL requires you to set up a free user account before using their translation service.  As of June 2022, the "Free" button at this link is the fastest way to setup a new, free account:

https://www.deepl.com/pro-api?cta=header-pro-api/

After setting up an account, you will be provided with a unique API key.  Simply copy and paste that key into the matching box in PhotoDemon's Language Editor, and PhotoDemon can now auto-suggest translations for you.  PhotoDemon will also save the pasted API key to your user preferences file, so you only need to enter it once.

Again, this feature **is not required** to edit PhotoDemon translations.  It is simply provided as a convenience for those who want it.  I initially added it to help me update language files that no longer have active contributors, and because I use it so frequently I thought it might also be useful to others.  If other automated translation services are more appropriate for your language, please contact me and I'll see if I can add them to the tool as well.

### Conclusion

Thank you again for helping me improve PhotoDemon's language files.  If you have any other questions, please contact me directly.  My contact info is available here:

http://photodemon.org/about/contact/

Kind regards,
Tanner Helland
(PhotoDemon developer)

Last modified: 28 June 2022