v1.4.2 Changelog
- Bugfix: QB TB clean now handles non-numeric accounts (eg, "1100F") properly.
- Changed: QB TB Clean now asks to include $0 balances (instead of exclude).

v1.4.2 Changelog
- Bugfix: QB TB clean now copies account numbers properly.
- Bugfix: Automerge only unmerges if a merged cell is the only selected cell.
- Added: Replace functions macro v2 (much more stable).

v1.4.1 Changelog
- Added: New default Word template.

v1.4.0 Changelog
- Bugfix: Fixed hang on load of spreadsheet.
- Bugfix: Eliminated annoying popup when removing hyperlinks.
- Bugfix: Made compare values function work again!
- Added: Support for QB Online TB clean. Also made the cleaned TB easier to read.
- Added: Add hyperlinks now bolds the text and makes the whole cell clickable.
- Added: Automerge now unmerges the cells if clicked again.
- Added: Trivial materiality now calculated properly.
- Removed: Removed update checking functionality.

v1.3.0 Changelog (completed 12/18/14, released 7/24/15)
 - replace functions functionality no longer shows popup dialog
 - fixed replace functions logic errors & added more error checking
 - hyperlinks now reference named ranges, to remain independent of sheet changes (row/col add/rem)
 
 v1.2.2 Changelog
 - redid materiality header calculations (rounding on PM & trivial)

 v1.2.1 Changelog
 - fixed error in replace functions logic
 - reorganized header in add header function (organization name on top)
 - reworked materiality in add header function (changed calculation of PM & trivial, renamed threshold)
 - added progress status to gap analysis and Benford test

 v1.2.0 Changelog
 - added materiality calculations option to add header button
 - added ability to define header width on header add (width of selection)
 - added option to ignore $0 balances on QB TB clean
 - added Benford test
 - added ability to hyperlink across tabs
 - added automerge functionality (merge, word wrap, text align, resets formatting)
 - added gap analysis functionality
 - added compare values functionality
 - added ability to replace all nonstandard Excel formulas with their values
 - added auto update functionality
