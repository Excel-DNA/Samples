# ArrayResizer sample

This code shows one way to implement automatic resizing of array results from UDFs for Excel versions that don't support dynamic arrays.

The new Excel Dynamic Arrays functionality is a much more robust and complete solution to this problem, and is discussed in the tutorial here: https://github.com/Excel-DNA/Tutorials/blob/master/SpecialTopics/DynamicArrays

To run the examples, add the .cs or .vb source code into a new or existing Excel-DNA add-in project, or take a copy of ExcelDna.xll / ExcelDna64.xll, rename to ArrayResizer.xll and then put it next to the `ArrayResizer.dna` and `ArrayResizer.cs` source files.
One the add-in is running in Excel, you should have the following functions to test:
* `dnaMakeArray` - return an array with no resizing
* `dnaMakeArrayDoubles` - return an array of doubles with no resizing
* `dnaResize` - function that takes any input array and resizes the result as an array formula covering a resized result regions
* `dnaResizeDoubles` - function that takes any input array of doubles and resizes the result as an array formula covering a resized result regions - might be faster than `Resize`
* `dnaMakeArrayAndResize` - the equivalent of `=Resize(MakeArray())`.
* `dnaMakeArrayAndResizeDoubles` - the equivalent of `=ResizeDoubles(MakeArrayDoubles())`.
* `dnaSupportsDynamicArrays` - Hidden function that detects whether the runnin instance of Excel supports dynamic arrays (and then skips the auto-resizing behaviour)
