µBMSC
=====
µBMSC is a modified version of iBMSC to add features and clean up the iBMSC code, fix bugs and so on.
See README.md.old for original iBMSC README file.

Changes in this fork
=====
* Changed keybinding to allow note placement between D1 and D8.
  * Numpad keys are now assigned to 2P lanes.
  * QWERTYUI keys are also assigned to 2P lanes.
  * 1 to 7 are now assigned to A2 to A8, and 8 is now assigned to A1.
  * Ctrl+1 to 8 are now assigned to D1-D8.
* Fixed the search function such that notes on lane A8 and D8 are now searchable.
* Fixed the mirror function such that notes between A1 and D8 are reflected locally.
* Added Random and S-Random. For S-random, note overlapping can occur.
* Fixed the Statistic Label not including notes between D1-D8. Statistic window still not fixed.
* Added a display for recommended #TOTAL.
* The application now saves the option "Disable Vertical Moves".
* Changed the temporary bms file extension from .bms to .bmsc.
* Allows selecting all file types when opening files.
* Fixed the total note count on the toolbar.
* Added advanced statistics (Ctrl+Shift+T).
* Added custom color for a specified range of notes. XML must be typed out manually for now.
* Removed restriction for drag and dropping files. Mainly for opening bms template files, not tested thoroughly.
* Added note search function (goto measure except it's goto note).
* Added sort function. Selected notes are sorted based on their VPosition and Value.
* Added keyboard shortcuts for previewing and replacing keysounds in the Sounds List (Spacebar to preview, enter to replace).
