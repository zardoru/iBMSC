pBMSC
=====
pBMSC is a modified version of uBMSC (which is a modified version of iBMSC) with a primary focus on quality of life functionalities such as keyboard shortcuts.
See README.md.old for original iBMSC README file.

# Changes
Listed in the order added.
## Bugfixes
* Added keybindings for lane D1-D8. See **Keyboard Shortcuts** for more information.
* Fixed the search function such that notes on lane A8 and D8 are now searchable.
* Fixed the mirror function such that notes between A1 and D8 are reflected locally.
* Fixed the Statistic Label not including notes between D1-D8. Statistic window still not fixed.
* Fixed the total note count on the toolbar.
* Reorganized the sidebar so you can tab between textboxes properly (mostly).
## Functionality
* Added Random and S-Random. For S-random, note overlapping can occur.
* Added a display for recommended #TOTAL.
* The application now saves the option "Disable Vertical Moves".
* Changed the temporary bms file extension from .bms to .bmsc.
* Added advanced statistics (Ctrl+Shift+T).
* Added custom color for a specified range of notes. XML must be typed out manually for now. The added window currently does nothing.
* Removed restriction for drag and dropping files, as well as opening files. Mainly for opening bms template files, not tested thoroughly.
* Added note search function (goto measure except it's goto note). One note per VPosition only.
* Added sort function. Selected notes are sorted based on their VPosition and Value.
## Keyboard shortcuts
* Changed keybinding to allow note placement between D1 and D8.
  * Numpad keys are now assigned to 2P lanes.
  * QWERTYUI keys are also assigned to 2P lanes.
  * 1 to 7 are now assigned to A2 to A8, and 8 is now assigned to A1.
  * Ctrl+1 to Ctrl+8 are now assigned to D1-D8.
* Added Save As keyboard shortcut (Ctrl+Alt+S)
* Added recent bms keyboard shortcuts (Alt+1 to Alt+5)
* Added shortcuts for toggling lanes.
  * Alt+B - BPM lane
  * Alt+S - Stop lane
  * Alt+C - Scroll lane
  * Alt+G - BGA/Layer/Poor
* Added shortcuts for the panel splitter (Alt+Left and Alt+Right).
* Added Options shortcut
  * F9 - Player Options
  * F10 - General Options
  * F12 - Visual Options
* Added advanced statistics (Ctrl+Shift+T)
* Added keyboard shortcuts for previewing and replacing keysounds in the Sounds List (Spacebar to preview, enter to replace).
