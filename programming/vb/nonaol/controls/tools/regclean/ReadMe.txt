README.TXT for
RegClean 4.1a (build 7364.1)                   March 13, 1998
-----------------------------------------------------------------

HOW TO INSTALL
--------------

Copy version 4.1a build 7364.1 of RegClean.EXE to any folder on your machine.

HOW TO RUN
----------

Double-click on the RegClean icon.

RegClean will start by displaying a progress dialog. While this
is shown, it will load a copy of the parts of the registry that
it is going to check, and perform the actual scanning. Depending
on how much information is in the registry and the speed of your
CPU, this will take anywhere from about 30 seconds to 30 minutes.

If you have many entries in your registry, there may be times
when RegClean appears to have stopped working. In fact, RegClean
may appear completely halted whenever it is checking remote or
removable drives.

Once these progress meters are gone, you will be prompted for the
next action. You can do two things at this point:

1. Exit RegClean.

   If RegClean did not find any errors in your registry, or, if you do    not want RegClean to fix the errors that it found, click the Cancel
button.

2. Allow RegClean to fix errors that it found.
   
   Clicking the Fix Errors button will cause RegClean to remove the           entries with errors it found in the registry. You will see a          progress meter while it does this. Sometimes, the progress meter may       stop momentarily, but it should resume after a few seconds. When the       meter is gone, RegClean is done. Press Exit to end RegClean.

   Clicking the Fix Errors button also creates an UNDO.REG file in the       folder where RegClean was run. The file will be titled "UNDO    computer yyyymmdd hhmmss.REG," where computer is the name of your       machine, yyyymmdd is the date, and hhmmss is the time. If at any    point you would like to "undo," or put back what RegClean removed    from your registry, double click the UNDO.REG file.


WHAT REGCLEAN DOES
------------------

RegClean analyzes Windows Registry keys stored in a common
location in the Windows Registry. It finds keys that contain
erroneous values, and after recording those entries in the
Undo.Reg file, it removes them from the Windows Registry.

WHAT REGCLEAN DOESN'T DO
------------------------

RegClean does not fix every known problem with the registry. It
does not fix a "corrupt" registry; it only fixes problems with
some of the entries that are in a normal registry.

It is very possible that RegClean will not correct a problem that
you have encountered. RegClean will leave any entries in the
registry that it doesn't understand or could possibly be correct.

----------------------------------------
GENERAL ISSUES
--------------

 - OLEAUT32.DLL
 - DOESN'T SOLVE PROBLEMS/MAKES NEW PROBLEMS
 - CAN'T UNDO THE REG FILE

----------------------------------------

FIX: REQUIRES UPDATED OLEAUT32

This error occurred with an earlier version of RegClean. In order for RegClean to work, you must have the updated version of the OLEAUT32.DLL. If your system has Internet Explorer version 2.0 or lower, installing Internet Explorer version 3.0 or higher will update the OLEAUT32.DLL. If you are using Microsoft Windows NT 3.51 with SP 4 or lower, you will need to install SP 5 in order to update the OLEAUT32.DLL. These files are also available for download below. If you experience the below symptoms when using RegClean 4.1a, please follow the resolution instructions.

SYMPTOMS
--------

Two messages boxes appear. One box says:

   "REGCLEAN.EXE is linked to missing export OLEAUT32.DLL:421"

while the next message box says:

   "A device attached to the system is not correctly functioning"

RESOLUTION
----------

You need to install the update to the OLE Automation system
libraries. These files are contained in the executable
OADIST.EXE, which is included in the download. (NOTE: If you are
using Window NT 3.51, you must use OADIST2Z.EXE instead; read the
KB article shown below.)

Alternately, you can find these files either at:

   http://support.microsoft.com/support/kb/articles/Q164/5/29.asp

 -or-

   ftp://ftp.microsoft.com/Softlib/MSLFILES/oadist.exe
   ftp://ftp.microsoft.com/Softlib/MSLFILES/oadist2z.exe

You can use most web browsers to download this file from either
location. This file is also on the Microsoft Download Library
BBS.

MORE INFORMATION
----------------

OADIST.EXE is a self-extracting Cabinet that will install the
Automation libraries on your system. You just need to download
the file OADIST.EXE and double-click to run it. It will prompt
you to be sure that you want to proceed. Answering "yes" will
update the Automation libraries on your computer.

This self-extracting Cabinet only works on Windows NT 4.0 and
Windows 95. If you are on Windows NT 3.51, you will need to get
OADIST2Z.EXE instead, which is still being published. Both
contain the same files. However, the OADIST2Z.EXE is a self-
extracting ZIP file with a README.TXT file that contains
installation instructions.

OADIST.EXE is 490096 bytes.
OADIST2Z.EXE is 559499 bytes.

----------------------------------------

PROBLEM: IT DOESN'T SOLVE MY REGISTRY PROBLEMS

 -or-

IT JUST CREATES MORE PROBLEMS

(status updated: 30 Dec 97)

Running RegClean may, in a few cases, cause other problems, such
as causing part of the Microsoft Network viewer to stop
functioning, or causing other programs to stop functioning. If
this happens to you, simply Undo the changes RegClean made, by
double-clicking on the last UNDO.REG file.

These types of problems are very rare, but it's a good idea to
keep your last UNDO.REG file for at least a few days or so.

Microsoft will continue improving RegClean to reduce the
frequency of problems like these.

----------------------------------------

PROBLEM: CAN'T UNDO THE UNDO.REG FILE

(status updated: 18 Mar 97)

SYMPTOM
-------

Windows displays several error message boxes when you try to
double-click the UNDO.REG file to undo the changes RegClean made.

RESOLUTION
----------

This problem is unrelated to RegClean. It's a problem with the
Associated Program for .REG files.

To correct this problem:

1. Go to any Explorer window, click the View menu, and select
   Options.

2. In the Options dialog, select the File Types tab.

3. Scroll down in the "Registered file types:" list until you
   find the entry called "Registration Entries."

4. Double-click this item or click on the Edit button.

5. In the Edit File Type dialog, select the "Merge" entry, and
   either double-click this item or click on the Edit button.

6. In the "Editing action for type: Registration Entries" dialog
   box, make sure that the text in the "Application used to
   perform action:" field has the following entry, including the
   double quotes:

      regedit.exe "%1"

7. Click the OK buttons to dismiss all three dialog boxes.

8. You should be able to double-click on the UNDO.REG file now.

----------------------------------------

Thanks for using RegClean! We hope you find this utility useful.

 - The RegClean Team

----------------------------------------

REGCLEAN IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
MICROSOFT DISCLAIMS ALL WARRANTIES, EITHER EXPRESS OR IMPLIED,
INCLUDING THE WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A
PARTICULAR PURPOSE. IN NO EVENT SHALL MICROSOFT CORPORATION OR
ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER INCLUDING
DIRECT, INDIRECT, INCIDENTAL, CONSEQUENTIAL, LOSS OF BUSINESS
PROFITS OR SPECIAL DAMAGES, EVEN IF MICROSOFT CORPORATION OR ITS
SUPPLIERS HAVE BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY
FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES SO THE FOREGOING
LIMITATION MAY NOT APPLY.

-----------------------------------------------------------------

