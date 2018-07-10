# SSMS Theme Editor

## Description:
The SSMS Theme Editor is a simple and clean theme editor for SSMS (SQL Server Management Studio).  It allows you to load, edit and save theme files in vssettings format. A preview is provided that displays instant feedback as to how the changes will be displayed in SSMS. The settings are trimmed down to what is actually important for SSMS and broken into three categories.

Three default themes are provided:
- Obsidian.DarkScheme.vssettings
- Solarized.DarkScheme.vssettings
- VisualStudio.DarkScheme.vssettings: based off the default Visual Studio dark theme

More styles can be downloaded from https://studiostyl.es/

### Screenshot:

![Example Screenshot](https://raw.githubusercontent.com/tcartwright/SSMSThemeEditor/master/SSMSThemeEditor/images/screenshot.png)

## How to Use:

- New: Starts a new theme, completely blank. 
- Load: Loads an existing theme that can be modified as desired.
- Save: Saves all changes to a vssettings file. All colors that have not been changed and are still blank will not write to the resulting file.

Blank colors will show as an X to signify they have not been set for that setting. You can right click an individual color to clear it out and set it back to transparent, or right click a tab to clear out all the colors for that tab.

Each setting has a foreground, or a background color button, and in some cases both. If the foreground or background buttons are not available that is because those settings should not have that value. As you change the colors the preview will be kept in synch with the changes you make. 

You can press CTRL-Z to undo changes, or CTRL-Y to redo changes. The application will store up to 50 changes.

Once you are done making any changes to your theme you can then save the theme file and import it into SSMS.

## Downloads:

[Latest Release](https://github.com/tcartwright/SSMSThemeEditor/releases/latest) 

