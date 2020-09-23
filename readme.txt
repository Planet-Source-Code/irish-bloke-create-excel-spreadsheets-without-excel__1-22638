CODE UPLOADED MY GERRY MC DONNELL
http://go.to/wdwtbam
=================================

http://www.vb-helper.com/HowTo/doexcel.zip

	Purpose
Write Excel files directly without using Excel

	Method
See the code and other text files.

Thanks to Dan Gardner (nambit@hotmail.com).

	Disclaimer
This example program is provided "as is" with no warranty of any kind. It is
intended for demonstration purposes only. In particular, it does no error
handling. You can use the example in any form, but please mention
www.vb-helper.com.
-----------

* What does this do?

This VB sample demonstrates one way to write Excel files nativly from VB.


* Why would I want to do that?

Because sometimes you don't want the hassle of loading Excel as an 
automation object, playing with the values and writing the file.
Sometimes you want to run some sort of daemon process which creates Excel
files on a box that doesn't have Excel. 


* Why would *you* want to do that?

I needed to send information to my customers requlary, and they were asking
for Excel files far too much. I could get away with sending CSV text for only
so long. I don't have Excel on the box running my daemon, so I decided to write
the files directly as binary, rather than paying for Excel.


* Which version of Excel does it use?

The class creates Excel 2 format files


* How do I use it?

Add ExcelFile.cls to your project

Dim EX1 as New ExcelFile
EX1.OpenFile "test.xls"
EX1.EWriteString 1,1,"test"
EX1.EWriteInteger 1,2,100
EX1.CloseFile


The parameters to EWriteString and EWriteInteger are row, column and data. Row
and Col start at 1,1 for cell A1.


* What about Bold, Italic and all that Jazz?

I didn't need them

I've only done text and integer fields, because that's all I needed. If 
someone fancies fixing my code and adding support for these things then please 
feel free.


* Where can I learn more?

I've included the two documents which I used as reference in this archive


* Your code is horrible

Yes, I know, but it started as a playing around kludge, and didn't get much better.
Again, if someone wants to fix the code so it looks better then feel free. It's so
short that it doesn't really matter to me.


* Contacts?

You can mail me at nambit@hotmail.com
There's no web address

* Licence?

Far too much trouble for 30 lines of code. Do with it what you will, but if you 
use it for anything really neat then let me know.


This document was written 17 Jun 1999
