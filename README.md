# FontFinder
Find all font types in use in Word, Excel or PowerPoint files

Tested in Office 2007

## Usage
Either use the example Macro Word document
or
* open VBA editor in MS Word
* add the class modules
  * cFile
  * cFolder
  * cTyper
* add the modules
  * modFiles
  * modFonts
* set references (Tools > References)
  * Microsoft Office 12 Object Library
  * Microsoft Excel 12 Object Library
  * Microsoft PowerPoint 12 Object Library
  * Microsoft Word 12 Object Library
  * Microsoft Scripting Runtime
  
  In the Immediate Window, run "PrintFontsInFolder [type], [path]".
  
  If you leave path empty, the program will use the path of the active document.
