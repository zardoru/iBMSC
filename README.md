µBMSC
=====
µBMSC is a modified version of iBMSC to add features and clean up the iBMSC code, fix bugs and so on.
See README.md.old for original iBMSC README file.

Changes in this fork
=====
* Changed keybinding to allow note placement between D1 and D8.
  * Numpad keys are now assigned to 2P lanes.
  * QWERTYUI keys are also assigned to 2P lanes.
  * 1 to 7 are now assigned to A1 to A7, and 8 is now assigned to A8.
  * Ctrl+1 to 7 are now assigned to D1-D8.
* Fixed the search function such that notes on lane A8 and D8 are now searchable.
* Fixed the mirror function such that notes between A1 and D8 are reflected locally.
* Added Random and S-Random. For S-random, note overlapping can occur.
* Fixed the Statistic Label not including notes between D1-D8. Statistic window still not fixed.
* Added a display for recommended #TOTAL.
* The application now saves the option "Disable Vertical Moves".
* Changed the temporary bms file extension from .bms to .bmsc.
* Allows selecting all file types when opening files
* Fixed the total note count on the toolbar
