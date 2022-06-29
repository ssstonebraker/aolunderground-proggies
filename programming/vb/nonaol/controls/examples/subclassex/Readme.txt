Subclass - Visual Basic subclassing control
Copyright (c) 1997 SoftCircuits Programming (R)
Redistributed by Permission.

This package includes a VB-authored subclassing control. A
subclassing control allows Visual Basic programs to detect messages
for which Visual Basic does not provide direct support for. This
controls is documented below. This package also includes a number of
subclassing example programs. These examples are unzipped to the
Examples directory (see Examples.txt for additional information).

The Subclass OCX is freeware that you can use freely with your own
programs. Any portion of the sample programs may also be incorporated
into your own applications. However, you may only distribute
Subclass.ocx as a) part of your own application that uses this control
or b) within this complete and unmodified package (i.e., you may
distribute the entire Subclass.zip file).

In addition, the source code to the control is also included. However,
to benefit everyone who uses this control, there are a number of
limitations as to how that source code can be used. See Subclass.txt,
which is in the same zip file as the Subclass source, for more
information on using the control's source code.

This example program was provided by:
 SoftCircuits Programming
 http://www.softcircuits.com
 P.O. Box 16262
 Irvine, CA 92623

======================================================================

Using Subclass.ocx
------------------

Note that you must register Subclass.ocx with the system registry in
order to use it. To do this, select the Project/Components command
within Visual Basic. Then check SoftCircuits Subclass Control from the
list. NOTE: This control was not designed for, nor tested with, Visual
Basic 4.

Properties
----------
hWnd		Handle of the window to be subclassed. Must be set at run time.

Messages	Specifies which messages you want to detect. Must be set at run
		time. The Messages property operates like an array and allows you
		to specify multiple messages like this:

		Subclass1.Messages(WM_MENUSELECT) = True 'Detect this message
		Subclass1.Messages(WM_SIZE) = True       'Detect this message
		Subclass1.Messages(WM_PAINT) = False     'Don't detect msg (default)

Methods
-------
CallWndProc	Invokes the original window procedure. Invoke only from within
		the WndProc event.

Events
------
WndProc		This event is invoked when any of the specified messages are sent
		to the specified window. The arguments are specific to the message
		sent (see the Windows API documentation for details). The Result
		argument is the value returned to Windows. This argument is 0 by
		default.
