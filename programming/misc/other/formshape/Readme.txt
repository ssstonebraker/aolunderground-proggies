Readme for VB Shaped Form Creator (VBSFC) 8/28/98
=================================================

Update 9/1/98:
~~~~~~~~~~~~~~
Version 2 released, with Edge Tracing command.  See help file for details.

Disclaimer:
~~~~~~~~~~~
This software is provided on an "as is" basis.  Nothing should go wrong,
but if it does, it is not my fault.  You use this software at your own risk.

License:
~~~~~~~~
This program is licensed for non-commercial use only.  If you intend to sell
your program, I think it is only fair you pay for mine (only 20$).  See
"Registration" in the help file for details on how to register online or by
mail.

Installation:
~~~~~~~~~~~~~
After you have downloaded VBSFC.zip, extract all the contents into the
folder where you want it installed, for exampe "c:\program files\accessories
\Visual Basic Shaped Form Creator\".  There should be 4 files; this one
(Readme.txt), the main executable (VB Shaped Form Creator.exe), the help
file (VBSFC.hlp) and the help contents (VBSFC.cnt).

This program requires the VB4 runtimes (VB40032.dll and OlePro32.dll) and the
two common control ocx's ComCtl32.ocx and ComDlg32.ocx.

If you are missing any of these files they can be downloaded from the my
website at http://www.comports.com/AlexV/VBSFC.html

Uninstallation:
~~~~~~~~~~~~~~~
Run VBSFC and choose "Remove All Settings" from the "Help" menu.  This will
remove all the programs registry entries.  You can then delete all the files
you installed.

WalkThrough 10 step easyguide thingy.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

This was requested, so I am putting it in.  90% of you won't need this, so
don't bother reading it if you don't.

10 steps to a Shaped Hello World example
========================================

1:  Load VB Shaped form creator.
2:  Draw your shape.  For example, stick an oval so that it just touches the top and left of the drawing area.  You can put it elsewhere if you want, but it is only wasting space.
3:  Save it as a form.  Put it somewhere you will remember.
(optional step --
3b:  Export it as a bitmap.  Put it in the same place as the form
--)
4:  Load VB.  You should start with a blank project.
5:  Remove the Form1 file (as we will not use it)
6:  Add a file to the project.  In VB4 this is done by right clicking on the project box and choosing "Add File".  Select the form you saved in step 3.
7:  Show the form (if not shown already)
(optional step --
7b:  Set the form's "picture" property to the picture you exported in step 3b.  This gives you
a guideline as to what bits of the form are displayed when it is run.
--)
8:  Put a label control and a button contol on the form.  Change the label caption to nothing, and the button caption to "Press Me" (or similar).  If you have done the optional steps, then you can make sure that both controls are within the coloured area (red by default), then remove the form's picture (as we don't need it any more).
9:  Double click the "Press Me" button to set code.  Type into the sub the line:   label1.caption = "Hello World"
10:  Run the project.  You will be told that you have not specified a startup form, and asked to do so.  Specify the "ShapedForm" form, and run the project.


------------------------------------------------------------------------------
Thank you for reading this document.

Email the author at:	 			AlexV@ComPorts.com
VB Shaped Form Creator website:		www.ComPorts.com/AlexV/VBSFC.html
Other programs by the same author:	www.ComPorts.com/AlexV

