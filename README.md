# Ghostscript-Pdf-Mark
Add Markings To a .pdf vector filed based on layers and rotation of a line inside the pdf. 
Insert a text from a variable with given Versal, Font and color, or insert external .eps drawing into the pdf.

The Script is setup to continously surveil a folder for input.

It takes and reads the .pdf file and scans for layers with the relevant names. It then reads the to and from coordinates of the lines on the layers and calculates
rotation and direction for the text to be inserted.

When an .eps file is inserted it is setup to add it on top of other graphics and a set distance from the text (Specific use case in this example). It can be changed.
