# copy-named-range-references
This is a small tool to copy worksheet-scoped named range references from one worksheet to another.

The main subroutine:

deletes all of the worksheet-scoped names on the new worksheet,

loops through all of the worksheet-scoped names on the old worksheet, and then generates new name references and new range references for each name, and adds the corresponding name and range to the new sheet.


The most imporant part though is the function that generates the new references. It uses a small regex pattern to change the references from one worksheet to another. It does catch that excel uses two single quotes in place of one single quote when placing a name of a worksheet into a range reference (e.g. "gary'sWorksheet" shows up in references as "gary''sWorksheet"). I'm not sure if there are other edge/corner cases that this doesn't catch.
