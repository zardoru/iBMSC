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


Check appveyor for automated builds.
[![Build status](https://ci.appveyor.com/api/projects/status/m7iygj9sje2yqf43?svg=true)](https://ci.appveyor.com/project/zardoru/ibmsc)
