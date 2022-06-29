Thank you for downloading ActiveZipper. I hope you find
it as useful for you as it was for me. Here are the 
instructions on how to use the control.


To load the control from VB4/32 either you can press CTRL+T or
right click on the toolbar full of controls and select 'Controls'.

After that just place it on the form of your choice.

ActiveZipper gives you only two properties. OutPutFile and SourceFile.
The SourceFile is what you want to compress, the OutPutFile is what will 
be the result of the compressed file (in its new filename). After you have
filled in those two properties, just one statement triggers compression:
ActiveZipper1.Compress This will compress the SourceFile. ActiveZipper WILL
check to see if the file is present, if not, it just won't compress, also
it will NOT start compressing while ActiveZipper is busy. It's a precautionary feature.

To decompress, put in the compressed filename in the SourceFile property and the
new file name in the OutPutFile property and just use this statement: ActiveZipper1.Decompress

PLEASE NOTE when it has finished compressing/decompressing, ActiveZipper will trigger a 'Completed'
event.

DUE TO SECURITY. IT WILL NOT COMPRESS/DECOMPRESS TO VITAL SYSTEM FILES. SUCH AS:
WIN.INI
SYSTEM.INI
SYSTEM.DAT
CONFIG.SYS
AUTOEXEC.BAT

ACTIVEZIPPER PRO RELEASED! NEW VERSION! 1.0f 1094!
Features such as...
1. Multi-file archive support (winzip like)
2. PKzip/PKunzip support!
3. Password protect archives
4. 32bit file encryption, encrypts a 163kb file in 8 seconds!
5. CRC Checking to ensure integrity
6. Custom compression level and speeds
7. Updated compression/decompression algorithm!
8. Very easy to program for it!

Download it at:
http://www.iessoft.com/download/azpzip.zip
Visit the ActiveZipper Pro homepage at:
http://www.iessoft.com/scripts/activezipperpro.asp

COPYRIGHT (C) 1996-1998 iNTERFACE ENTERPRISES
Visit iNTERFACE Enterprises on the WWW at: http://www.iessoft.com