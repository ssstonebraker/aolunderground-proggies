AOL6LIST ActiveX DLL README
developed by huma

AOL6LIST.dll is an ActiveX DLL developed in Visual Basic 6 SP4. If you plan on using AOL6LIST.dll 
in your vb projects, you must register the dll first using RegSvr32.exe and set a reference to the dll
in your visual basic projects pointing to the dll.

Keep in mind that if you plan on distributing a program that uses the AOL6LIST.dll you must 
include AOL6LIST.dll and it must be registered on the user's computer before it can be used.
It also require's the Visual Basic 6.0 Runtime files

Usage:

The CAol6List object only has one method and two properties. 

The arguments for the GetScreennamesFromList method are:

	hWnd: Window handle of a list on AOL.
	listType: You have two options, lteListbox for listboxes, lteCombobox for comboboxes. Default is lteListbox
	accountLength: Buffer length of a screen name. This argument is optional and is best left alone.
	accountCount: Screenname count to return. This argument is optional and is best left alone.

The return value is boolean TRUE/FALSE indicating if it was successful. If there are no screen names that were added, then
a FALSE value is returned.

The Listcount property returns the number of screen names in the list. An empty list minus 1 will return -1. You would 
use the listcount property like a regular listbox. which means the total number of items in the list is always 
minus one. e.g. 

dim index as long

	For index = 0 to clsAol.Listcount - 1          '<|------- minus 1;	where clsAol is an instance of the CAOL6List class object
	Next

The Item property returns the screen name from the list. The argument is:
	index: the item number of the screen name

debug.print clsAol.Item(0) 'where clsAol is an instance of the CAOL6List class object

In Visual Basic, select the Project, References to bring up the References dialog. If AOL6LIST.dll is registered, 
select 'AOL6LIST by Huma' from the list and press the OK button to close the dialog. If it is not registered, press 
the Browse button and find the location of AOL6LIST on your computer and select it to register the dll.

Once you have a reference to the dll, you need to create an instance of the CAol6List object to use.
The code below will not work without first defining the window handle (hWnd) and the list type (listType)

Dim clsAol As CAol6List 'define the variable to bind to the object.
Dim index As Long 'current item index for iteriating a list

Set clsAol = New CAol6List 'create an instance of the object

    'once the object has been created, we can call the GetScreennamesFromList method
    'from the object. If the value returns "TRUE" then you can iteriate through the list
    If clsAol.GetScreennamesFromList(hWnd, lteListbox) = True Then
        'iteriates through each item in the list and prints the screen name to the debug window
        For index = 0 To clsAol.ListCount - 1
            Debug.Print clsAol.Item(index)
        Next
    End If

Set clsAol = Nothing 'destroy the object



or

Dim clsAol As Object 'define the variable to perform a late bind to the object
Dim index As Long 'current item index for iteriating a list
Dim listType As Long 'value can be 0 or 1

Set clsAol = CreateObject("AOL6LIST.CAol6List") 'create an instance of the object

    'once the object has been created, we can call the GetScreennamesFromList method
    'from the object. If the value returns "TRUE" then you can iteriate through the list
    If clsAol.GetScreennamesFromList(hWnd, listType) = True Then
        'iteriates through each item in the list and prints the screen name to the debug window
        For index = 0 To clsAol.ListCount - 1
            Debug.Print clsAol.Item(index)
        Next
    End If

Set clsAol = Nothing 'destroy the object








http://welcome.to/huma
http://direct.at/huma/
Nai (h4ma@yahoo.com)