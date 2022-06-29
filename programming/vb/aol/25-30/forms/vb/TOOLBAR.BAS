Option Explicit

' Declare a global variable to hold a count of the number of buttons on the toolbar.
Global gnButtons As Integer

Sub LoadImage (imgControl As Control, sSuffix As String)

    ' Loads an image into the specified toolbar image control. The file name
    ' prefix is pulled from the tag property and sSuffix is added to it.
    '
    ' The Tag property of the toolbar buttons is set up with the first character holding
    ' the menu item index number, and the rest of the tag holding the picture prefix, ie
    '       1 Open
    '
    ' If this sub is called with a suffix of "dn" then the image loaded is
    '               OPENDN.BMP
    '
    ' Images are always loaded in with this procedure from the applications current path.
    '

    Dim sFileName As String

    ' First the filename of the required toolbar image is built up
    sFileName = Mid$(imgControl.Tag, 3, Len(imgControl.Tag) - 2)
    sFileName = app.Path & "\" & sFileName & sSuffix & ".bmp"

    ' The image is then loaded.
    imgControl.Picture = LoadPicture(sFileName)
    
End Sub

