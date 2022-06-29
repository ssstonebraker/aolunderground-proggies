********************************************************************************
 CSCEXFTP Extended FTP control					

 DISCLAIMER:  This is a 'use at your own risk' control.  I haven't had any reason
 to believe this control doesn't work exactly as it should, so chances are you
 won't have problems with it.  If you do, send me an email.
********************************************************************************

The CSCEEXFTP control is designed to be a FREE, easy-to-use replacement for the currently
available FTP controls.  It was designed for use with Visual Basic 6.0.

Requirements:

VB 5.0 or 6.0 for developers (VB6 runtime needed)


Properties:

DropBehavior
-This property controls whether or not the control does automatic uploads or not.  If the DropBehavior
 property is set to cscAutomatic, any file or group of files dragged on to the control will be
 automatically uploaded to the connected ftp server.  During the upload, the control will generate
 a status form with a percentage and a progress bar.

ShowStatus
-When this property is true, any Asynchronous upload you perform when calling the PutFile method
 will invoke the automatic status form.

TransferType
-This controls whether files are transferred in binary or ASCII modes

Picture
-Since this control supports drag and drop uploading, this property allows you to choose a picture 
 for the face of the control that the user can drop files on.

Methods:

Connect
-Connects to an ftp site when you call this method and pass in a Server name, username, password,
 and ftp syntax.  ie: CSCEXFTP1.Connect("ftp.microsoft.com","Anonymous","guest",Active)

DisConnect
-Disconnects from an ftp site

GetFile
-Downloads a file Synchronously from an FTP site when the local and remote file names are 
 specified.  Note that to get a file from a directory other than the root, add the directory
 and a "\" to the remote file name.  Asynchrounous downloads are not supported in this version

PutFile
-Uploads a file Synchronously or Asynchronously to the connected FTP site.  To put a file in a subdirectory,
 prefix the remote file name with the directory and a "\".



***************************************************************************************************
SPECIAL NOTE:

This control uses Microsofts WININET API from wininet.dll.  I't hasn't been tested on anything but
IE 4.0.  If you're running IE 3.0, or 5.0, and experience problems, send me an email
***************************************************************************************************


Author's Comments:

I designed this control a while ago, and am releasing it now for free instead of rounding out
the feature set I wanted it to have.  If you would like to see some specific functionality that
the control does not posess, send me an email, and I'll see if i can't accommodate your requests.
All comments and requests should be sent to:

FVandervelde@coyotecorp.com OR rockyb@sympatico.ca

Thanks,

Fred Vandervelde