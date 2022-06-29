DefInt A-Z
Dim HelpFilePath As String

' When either the Editor's colorpalette or the ColorPalette Forms
' ColorPalette need repainting, this routine is called, passing in
' the picture control used for the specific colorpalette.
'
Sub Display_Color_Palette (Pic_ColorPalette As Control)
    
    ' The ColorPalettes consist of 3 rows of 16 colors, so to make
    ' is easy to display and to deterine what color is selected when
    ' the ColorPalette is click, we set the Scale of the ColorPalette
    ' to correspond to the number of color rows and columns.
    '
    Pic_ColorPalette.Scale (0, 0)-(16, 3)

    ' Display ColorPalette column by column
    '
    For I = 0 To 15
        '
        ' Display a column of colors
        '
        Pic_ColorPalette.Line (I, 0)-(I + 1, 1), Colors(I), BF
        Pic_ColorPalette.Line (I, 1)-(I + 1, 2), Colors(I + 16), BF
        Pic_ColorPalette.Line (I, 2)-(I + 1, 3), Colors(I + 32), BF

        ' Display vertical line to left of current columns to visually
        ' divide the columns, but skip first column, since it is not
        ' needed due to the Border of the color palette.
        '
        If I Then Pic_ColorPalette.Line (I, 0)-(I, 3)
    Next I
  
    ' Display 2 horizontal lines to visually divide the color rows.
    '
    Pic_ColorPalette.Line (0, 1)-(16, 1)
    Pic_ColorPalette.Line (0, 2)-(16, 2)

End Sub

' Displays the entire or any portion of the grid, when the Grid option
' is active.  The 4 paramaters passed in, X1, Y1, X2, Y2, define the
' upper left and lower right corners of the area within the maginified
' Icon that needs the grid displayed.
'
Sub Display_Grid (DestHdc, X1, Y1, X2, Y2)
  
    ' The grid is not displayed if the icon is being viewed at normal
    ' size, so check the current value of the scrollbar.
    '
    If editor.Scrl_Zoom.Value > editor.Scrl_Zoom.Min Then
        DestX = X1 * PixelSize
        DestY = Y1 * PixelSize
        DestWidth = (X2 - X1 + 1) * PixelSize
        DestHeight = (Y2 - Y1 + 1) * PixelSize
        R = BitBlt(DestHdc, X1 * PixelSize, Y1 * PixelSize, DestWidth, DestHeight, editor.Pic_Grid.hDC, DestX, DestY, SRCAND)
    End If

End Sub

' Whenever a new color is selected for either the left or right mouse
' button, or the StatusArea needs repainting, this routine is called to
' display the 4 small color squares at the bottom of the StatusArea
' which are filled with the current colors selected for the mouse buttons.
'
Sub Display_Mouse_Colors ()

    ' Calculate the center of the Status bar
    '
    Middle = editor.Pic_StatusArea.ScaleWidth \ 2

    ' Display the 4 color squares
    '
    For I = 0 To 3
        '
        ' The squares are centered within the left and right halfs of the
        ' StatusArea, and the width and height are set equal to the Height
        ' of the Option buttons used to select Left/Right or Screen/Inverse
        ' colors, so we calculate the corners of the the Color squares
        ' based on this information.
        '
        X1 = (I Mod 2) * Middle + (Middle - editor.Opt_Mouse(I \ 2).Height) \ 2
        X2 = X1 + editor.Opt_Mouse(I \ 2).Height
        Y1 = editor.Opt_Mouse(I \ 2).Top
        Y2 = Y1 + editor.Opt_Mouse(I \ 2).Height

        ' Draw the color square
        '
        editor.Pic_StatusArea.Line (X1, Y1)-(X2, Y2), MouseColors(I), BF

        ' Draw a black outline around the square
        '
        editor.Pic_StatusArea.Line (X1, Y1)-(X2, Y2), BLACK, B
    Next I
        
    ' Set the CurrentY value of the StatusArea back to that of the
    ' location where the Mouse Coordinates are displayed, so this
    ' does not have to be done within each MouseMove event of the
    ' Edit area.
    '
    editor.Pic_StatusArea.CurrentY = editor.Pic_Icons(5).Top + editor.Pic_Icons(5).Height + HIGHLIGHT + 1

End Sub

' If a selection has been made, is being made, or a selection is
' being moved, or the Edit area needs repainting while a selection
' is active, this routine is called to display or redisplay a
' rectangle around the current selection.
'
Sub Draw_Selection_Rectangle ()
 
    ' Set drawing mode to INVERSE since this routine also used to erase
    ' the selection rectangle by simply drawing over the currently displayed
    ' rectangle
    '
    editor.Pic_Edit.DrawMode = Inverse

    ' To distinguish between a selection and a selection that is
    ' being moved, a Dotted line is used for a selection and a solid
    ' line is used for a selection being moved.
    '
    If MovingSelection Then editor.Pic_Edit.DrawStyle = SOLID Else editor.Pic_Edit.DrawStyle = DOT

    ' To ensure the entire selection rectangle is visible, the rectangle
    ' is adjusted inward 1 pixel from the right and bottom if the selection
    ' contains either the right most column or bottom most row of pixels.
    '
    If X2Region >= PixelSize * 32 Then XAdjust = 1
    If Y2Region >= PixelSize * 32 Then YAdjust = 1

    ' Draw the selection rectangle.
    '
    editor.Pic_Edit.Line (X1Region, Y1Region)-(X2Region - XAdjust, Y2Region - YAdjust), , B
    editor.Pic_Edit.DrawStyle = SOLID

End Sub

' When the currently selected Icon is changed or a new Icon is
' loaded into the currently selected Icon, the bitmaps that make
' of the Icons Mask and Image must be extracted and placed into
' picture controls where they can easily be edited.
'
Sub Extract_Image_And_Mask (Pic_Ctrl As Control)
Dim Lpicon As Long
    
    ' Get pointer to Icon and prevent Windows form moving it.
    '
    Lpicon = GlobalLock(Pic_Ctrl.Picture)

    ' Copy the Icons Mask to Monochrome Bitmap, then copy the MonoBitmap
    ' the the Picture control.
    '
    editor.Pic_Mask.ForeColor = BLACK
    R = SetBitmapBits(MonoHbm, 128, Lpicon + 12)
    R = BitBlt(editor.Pic_Mask.hDC, 0, 0, 32, 32, MonoHdc, 0, 0, SRCCOPY)

    ' Copy Icons Image bitmap to Picture control
    '
    R = SetBitmapBits(editor.Pic_Image.Image, ImageSize, Lpicon + 12 + 128)

    ' Free icon so Windows is free to move it.
    '
    R = GlobalUnlock(Pic_Ctrl.Picture)

End Sub

' Displays the selected help topic selected from either
' Editors;' or Viewer's help menu.
'
Sub Get_Help (HelpTopic As Integer)
  
    If HelpTopic = MID_USING_HELP Then
        '
        ' "Using Help" was selected so display the Standard Windows Help
        ' Topic for "Using Help".
        '
        R = WinHelp(editor.hWnd, Dummy$, HELP_HELPONHELP, 0)
    Else
        ' A help topic other the "Using help" was selected.
        '
        
        R = WinHelp(editor.hWnd, HelpFilePath, HELP_CONTEXT, CLng(HelpTopic))
    End If

End Sub

Function Help_File_In_Path ()
Dim Path As String, CurrentDir As String

    On Error Resume Next
    CurrentDir = App.Path
    If Right$(CurrentDir, 1) <> "\" Then CurrentDir = CurrentDir + "\"
    Found = Dir$(CurrentDir + "IconWrks.HLP") <> ""
    If Not Found Then
        Path = Environ$("PATH")
        If Path <> "" Then
            If Right$(Path, 1) <> ";" Then Path = Path + ";"
            SemiColon = InStr(Path, ";")
            Do
                CurrentDir = Left$(Path, SemiColon - 1)
                If Right$(CurrentDir, 1) <> "\" Then CurrentDir = CurrentDir + "\"
                Found = Dir$(CurrentDir + "IconWrks.HLP") <> ""
                Path = Right$(Path, Len(Path) - SemiColon)
                SemiColon = InStr(Path, ";")
            Loop While ((SemiColon <> 0) And Not Found)
        End If
    End If
    If Found Then
        HelpFilePath = CurrentDir + "IconWrks.HLP"
        App.HelpFile = CurrentDir + "IconWrks.HLP"
    End If
    Help_File_In_Path = Found
    
    On Error GoTo 0

End Function

' The currently selected icon is distinguished by a solid square
' slightly larger than the icon itself, drawn behind the icon using
' the currently selected screen color.  This routine is called
' whenever this square needs to be displayed or redisplayed.
'
Sub HighLight_Current_Icon ()
  
    ' Erase the current selection square.
    '
    editor.Pic_StatusArea.Line (0, 0)-(editor.Pic_StatusArea.Width, editor.Pic_Icons(4).Top + editor.Pic_Icons(4).Height + 10), editor.Pic_StatusArea.BackColor, BF

    ' Calculate the upper left and lower right corners of the selection square.
    '
    X1 = editor.Pic_Icons(CurrentIcon).Left - HIGHLIGHT
    X2 = editor.Pic_Icons(CurrentIcon).Left + editor.Pic_Icons(CurrentIcon).Width + HIGHLIGHT
    Y1 = editor.Pic_Icons(CurrentIcon).Top - HIGHLIGHT
    Y2 = editor.Pic_Icons(CurrentIcon).Top + editor.Pic_Icons(CurrentIcon).Height + HIGHLIGHT
  
    ' Draw the solid selection square.
    '
    editor.Pic_StatusArea.Line (X1, Y1)-(X2, Y2), MouseColors(2), BF

    ' Draw a Black outline around the square.
    '
    editor.Pic_StatusArea.Line (X1, Y1)-(X2, Y2), BLACK, B

    If editor.Menu_ViewSelection(MID_BORDER).Checked Then
        '
        ' Show edge of selected Icon by outline the icon
        '
        X1 = editor.Pic_Icons(CurrentIcon).Left - 1
        X2 = editor.Pic_Icons(CurrentIcon).Left + editor.Pic_Icons(CurrentIcon).Width
        Y1 = editor.Pic_Icons(CurrentIcon).Top - 1
        Y2 = editor.Pic_Icons(CurrentIcon).Top + editor.Pic_Icons(CurrentIcon).Height
        editor.Pic_StatusArea.Line (X1, Y1)-(X2, Y2), BLACK, B
    End If
    
    ' Set the CurrentY value of the StatusArea back to that of the
    ' location where the Mouse Coordinates are displayed.
    '
    editor.Pic_StatusArea.CurrentY = editor.Pic_Icons(5).Top + editor.Pic_Icons(5).Height + HIGHLIGHT + 1
    
End Sub

' Inverts the specified control when an Icon from the Viewer is being
' dragged over the top of it, signaling that the Icon may be dropped
' on this control.
'
Sub Invert_Control (Ctrl As Control)
Dim rectangle As RECT
  
    ' Calculate the Rectangle to invert
    '
    rectangle.Right = Ctrl.ScaleWidth
    rectangle.bottom = Ctrl.ScaleHeight

    ' Invert the rectangle
    '
    R = InvertRect(Ctrl.hDC, rectangle)

End Sub

' This routine is used to tie the Viewer and the Editor together.  When
' and Icon is selected in one of the various ways from within the Viewer,
' or an Icon is dragged from the Viewer and dropped on a valid location
' of the Editor, this routine is called either from the Viewer or from
' the Editor (depending on how the Icon was selected), to load the
' selected icon into the Editor.
'
Sub Load_An_Icon ()

    ' Check if the new icon would be replacing an existing Icon which
    ' has been changed since the last time it has been saved, and if
    ' so, ask the user if it is ok to discard the changes.
    '
    If Ok_To_Discard_Changes() Then
        '
        ' Get the Filename and Fullpath to the icon, and set its
        ' Changed flag to FALSE.
        '
        IconInfo(CurrentIcon).FileName = Viewer.File_FileList.FileName
        IconInfo(CurrentIcon).FullPath = Viewer.File_FileList.Path
        IconInfo(CurrentIcon).Changed = False

        ' Place the Name and Path of the Icon in the corresponding menu
        ' item in the Editors Icons menu.
        '
        editor.Menu_IconsSelection(CurrentIcon).Caption = "&" + Format$(CurrentIcon + 1) + " - [" + Viewer.File_FileList.Path + "]" + A_TAB + Viewer.File_FileList.FileName

        ' Load the Icon into the selected icon in the StatusArea.
        '
        editor.Pic_Icons(CurrentIcon).Picture = LoadPicture(Viewer.File_FileList.FileName)

        ' If the Menu option is set, bring the Editor to the Foreground
        ' when an Icon is loaded.
        '
        If editor.Menu_ViewSelection(MID_FOCUS).Checked Then editor.Show

        ' Simulate clicking the Icon in the StatusArea to take care of the
        ' visual part of selection.
        '
        Select_New_Icon
        editor.Pic_ToolPalette.Refresh
    Else
        ' Do not discard the changes of the existing icon.
        '
        editor.Pic_Icons(CurrentIcon).Cls
        Magnify_Icon 0, 0, 31, 31
    End If

End Sub

' There are various situations when all or part of the current icon
' needs to be magnified and displayed in the editing area.  this
' routine is called to perform the magnification.  The Windows API
' routine, StretchBlt() is used to perform the magnification.
'
Sub Magnify_Icon (X1, Y1, X2, Y2)
    
    ' Ensure that X1 and Y1 refer to the upper left corner and X2 and Y2
    ' refer to the lower right corner of the area to be magnified.
    '
    If X1 > X2 Then Swap_Values X1, X2
    If Y1 > Y2 Then Swap_Values Y1, Y2

    ' The area to be magnified must not contain any pixels outside
    ' of the Icon itself, so we must check for this situation and
    ' adjust the values if neccessary.
    '
    If X1 < 0 Then X1 = 0
    If X2 > 31 Then X2 = 31
    If Y1 < 0 Then Y1 = 0
    If Y2 > 31 Then Y2 = 31

    ' Calculate the width and height values of the source bitmap
    '
    SrcWidth = X2 - X1 + 1
    SrcHeight = Y2 - Y1 + 1

    ' Calculate the destinations width, height and upper left corner
    ' of the area to be magnified.
    '
    DestX = X1 * PixelSize
    DestY = Y1 * PixelSize
    DestWidth = SrcWidth * PixelSize
    DestHeight = SrcHeight * PixelSize
  
    ' Magnify the icon.  We StretchBlt() from the image of the Icon in
    ' the StatusArea to the Editing area.  Since we always maintain the
    ' size of the Editing area a multiple of 32 (Size of an Icon), the
    ' magnified icon will always be a perfect enlargement of the Icons
    ' image.
    '
    If ImageSize = 1024 Then
        '
        R = StretchBlt(editor.Pic_Edit.hDC, DestX, DestY, DestWidth, DestHeight, editor.Pic_Icons(CurrentIcon).hDC, X1, Y1, SrcWidth, SrcHeight, SRCCOPY)
        '
        ' Redisplay the grid in the area that was magnified if the Grid option
        ' is currently selected.
        '
        If editor.Menu_ViewSelection(MID_GRID).Checked Then Display_Grid (editor.Pic_Edit.hDC), X1, Y1, X2, Y2
    Else
        '
        R = StretchBlt(editor.Pic_EditTemp.hDC, DestX, DestY, DestWidth, DestHeight, editor.Pic_Icons(CurrentIcon).hDC, X1, Y1, SrcWidth, SrcHeight, SRCCOPY)
        '
        ' Redisplay the grid in the area that was magnified if the Grid option
        ' is currently selected.
        '
        If editor.Menu_ViewSelection(MID_GRID).Checked Then Display_Grid (editor.Pic_EditTemp.hDC), X1, Y1, X2, Y2
        R = BitBlt(editor.Pic_Edit.hDC, DestX, DestY, DestWidth, DestHeight, editor.Pic_EditTemp.hDC, DestX, DestY, SRCCOPY)
    End If

    ' Check if there is an active selection in the Editing area.  If so,
    ' we must also redisplay the contents of the selection since the above
    ' StretchBlt() operation may have entirely or partially covered up
    ' the selection.
    '
    If MovingSelection Then
        '
        ' Calculate the width and height values of the source bitmap
        ' containing the selection.  Always maintained in the global values
        ' X1SelectFrom, Y1SelectFrom, X2SelectFrom, and Y2SelectFrom
        '
        SrcWidth = X2SelectFrom - X1SelectFrom
        SrcHeight = Y2SelectFrom - Y1SelectFrom
        
        ' Calculate the destinations width and height of the area to be magnified.
        '
        DestWidth = SrcWidth * PixelSize
        DestHeight = SrcHeight * PixelSize

        ' Determine type of Selection: Opaque, or Not Opaque.
        '
        If Opaque Then
            '
            ' Opaque selection: Magnify the selection bitmap including any Screen
            ' or Inverse Screen attributes
            '
            R = StretchBlt(editor.Pic_Edit.hDC, X1Region, Y1Region, DestWidth, DestHeight, editor.Pic_Work.hDC, X1SelectFrom, Y1SelectFrom, SrcWidth, SrcHeight, SRCCOPY)
        Else
            ' None Opaque Selection: Magnify the selection bitmap but do not include
            ' any Screen or Inverse Screen attributes.
            '
            R = StretchBlt(editor.Pic_Edit.hDC, X1Region, Y1Region, DestWidth, DestHeight, editor.Pic_TempMask.hDC, X1SelectFrom, Y1SelectFrom, SrcWidth, SrcHeight, SRCAND)
            R = StretchBlt(editor.Pic_Edit.hDC, X1Region, Y1Region, DestWidth, DestHeight, editor.Pic_TempImage.hDC, X1SelectFrom, Y1SelectFrom, SrcWidth, SrcHeight, SRCINVERT)
        End If
    End If
  
    ' Redisplay the selection rectangle if currently making a selection
    '
    If Selecting Then Draw_Selection_Rectangle

End Sub

' A Sub Main is used instead of a startup form to allow the user
' to startup either the Editor or Viewer as the main form.  The
' Editor is the Default main form, however starting IconWorks
' with a command line of "v" or "V" will start IconWorks with
' the Viewer as the main form.
'
Sub Main ()
  
    ' Check video mode.  If less than EGA, terminate Iconworks
    '
    If Screen.Height < EGA_HEIGHT Then
        MsgBox "IconWorks requires EGA or Better.", 16, "IconWorks"
        End
    Else
        ' Since you cannot assign values like TAB, CR, and LF to string
        ' constants, the values of TAB and CRLF which are used frequently
        ' thoughout IconWorks when displaying messages, these values are
        ' are assigned to the global string values of A_TAB and CRLF
        '
        A_TAB = Chr$(9)
        CRLF = Chr$(13) + Chr$(10)

        If Not Help_File_In_Path() Then
            Text = "ICONWRKS.HLP not found in your path." + CRLF + CRLF
            Text = Text + "Windows searches your PATH environment variable for help files, "
            Text = Text + "so you need to copy ICONWRKS.HLP to a directory included in your "
            Text = Text + "PATH if you wish to obtain help while running IconWorks."
            MsgBox Text, 48, "IconWorks help not available"
        End If
        
        ' Determine which form to use as main form, Editor) or the Viewer
        '
        If (Command$ = "") Or (UCase$(Left$(Command$, 1)) <> "V") Then
            '
            ' Editor is main form
            '
            MainForm = ICONWORKS_EDITOR
            editor.Show
        Else
            ' Viewer is main form
            '
            MainForm = ICONWORKS_VIEWER
            Viewer.Show
        End If
    End If

End Sub

' Determines if an Icon has been modified since it was saved last, and
' prompts the user if so.
'
Function Ok_To_Discard_Changes ()

    Text = ""
    Ok_To_Discard_Changes = True

    ' Check if Icon has changed since it was last saved.
    '
    If IconInfo(CurrentIcon).Changed Then
        '
        ' Inform user icon has been modifyied.
        '
        Text = Text + "Icon:" + A_TAB + "#" + Format$(CurrentIcon + 1) + CRLF
        Text = Text + "Name:" + A_TAB + IconInfo(CurrentIcon).FileName + CRLF
        Text = Text + "Path:" + A_TAB + IconInfo(CurrentIcon).FullPath + CRLF + CRLF
        Text = Text + "Discard changes?"
        Ok_To_Discard_Changes = MsgBox(Text, 36, "ICON HAS CHANGED") = MBYES
    End If

End Function

' Removes various menu items from the System menu of the specified Form.
'
Sub Remove_Items_From_Sysmenu (A_Form As Form)

    ' Obtain the handle to the forms System menu
    '
    HSysMenu = GetSystemMenu(A_Form.hWnd, 0)
  
    ' Remove all but the MOVE and CLOSE options.  The menu items
    ' must be removed starting with the last menu item.
    '
    R = RemoveMenu(HSysMenu, 8, MF_BYPOSITION) 'Switch to
    R = RemoveMenu(HSysMenu, 7, MF_BYPOSITION) 'Separator
    R = RemoveMenu(HSysMenu, 5, MF_BYPOSITION) 'Separator

End Sub

' The rectanglular Region which is always defined by the global
' variables X1Region, Y1Region, X2Region, and Y2Region, is the
' basis for most of the tools in the toolpalette, and is frequently
' scaled from the scale of the Editing area down to the scale of
' the actual Icon, and in the reverse direction.  This routine
' performs the neccessary scaling, in either direction based on
' the value of *ToIcon*.
'
Sub Scale_Region (ToIcon, X1, Y1, X2, Y2, CheckX1Y1)
  
    ' Determine which direction to scale
    '
    If ToIcon Then
        '
        ' Scale Global variables down to and Icon
        '
        X1 = X1Region \ PixelSize
        Y1 = Y1Region \ PixelSize
        X2 = X2Region \ PixelSize
        Y2 = Y2Region \ PixelSize
    
        ' If requested, ensure X1 and Y1 refer to upper left corner
        ' and X2 and Y2 refer to the lower right corner of the Region.
        '
        If CheckX1Y1 Then
            If X1 > X2 Then Swap_Values X1, X2
            If Y1 > Y2 Then Swap_Values Y1, Y2
        End If
    Else
        ' Scale the values X1, Y1, X2, Y2 upto the Editing
        ' area and assign to global variables
        '
        X1Region = X1 * PixelSize
        Y1Region = Y1 * PixelSize
        X2Region = X2 * PixelSize
        Y2Region = Y2 * PixelSize
    End If
  

End Sub

' When a new Icon from one of the 6 displayed within the StatusArea is selected
' or if a new icon is selected from the viewer to be edited, this routine is
' called to take care of the visual changes within the StatusArea.
'
Sub Select_New_Icon ()
    
    Selecting = False
    MovingSelection = False

    HighLight_Current_Icon

    Extract_Image_And_Mask editor.Pic_Icons(CurrentIcon)
      
    ' Set the Undo Icon to the newly selected Icon.
    '
    Update_Icon editor.Pic_Undo

    ' Display the icon in the editing area
    '
    Magnify_Icon 0, 0, 31, 31

    ' Display the Filename of the selected icon in the Editor's Titlebar
    '
    editor.Caption = "IconWorks Editor: " + Format$(CurrentIcon + 1) + " - " + IconInfo(CurrentIcon).FileName

End Sub

' Since the Swap statement is not supported by Visual Basic, this
' routine is used to perform the task of swapping two integer values.
'
Sub Swap_Values (Param1, Param2)

    temp = Param1
    Param1 = Param2
    Param2 = temp

End Sub

' This routine is used by the SaveFileDlg and the Viewer to update the
' filespec displayed in the FileName TextBox whenever the forms Directory
' ListBox control is Single Clicked.  Since a Single click does not
' actually make a selection, this routine is called in response to a
' single click to display what would be the result if a double click
' is performed or if Enter is pressed.
'
Sub UpDate_FileSpec (A_Form As Form)
Dim SelPath As String, CurPath As String, Slash As String

    CurPath = A_Form.Lbl_CurrentDirectory.Caption
    SelPath = A_Form.Dir_DirectoryList.List(A_Form.Dir_DirectoryList.ListIndex)

    Select Case A_Form.Dir_DirectoryList.ListIndex
        
        Case Is >= 0
            '
            ' A subdirectory from the Current directory was selected
            '
            I = Right$(CurPath, 1) <> "\"
            A_Form.Txt_FileName.Text = Right$(SelPath, Len(SelPath) - Len(CurPath) + I) + "\" + A_Form.File_FileList.Pattern
        
        Case Is = -1
            '
            ' The current directory was selected
            '
            A_Form.Txt_FileName.Text = A_Form.File_FileList.Pattern
        
        Case Is < -1
            '
            ' A parent directory of the Current directory was selected
            '
            SelPath = Right$(SelPath, Len(SelPath) - 2)
            If Len(SelPath) > 1 Then Slash = "\"
            A_Form.Txt_FileName.Text = SelPath + Slash + A_Form.File_FileList.Pattern
    
    End Select

End Sub

' We do not actually modify the Icon directly, but modify the Mask and Image
' bitmaps that make up the Icon. So these bitmaps must be copied over the icons
' Mask and Image bitmaps after each edit to reflect the change in the actual
' icon displayed in the StatusArea.
'
Sub Update_Icon (Pic_Ctrl As Control)
Dim Lpicon As Long
  
    ' Convert the 4-Plane Mask Bitmap contained in the Picture Control to
    ' a 1-Plane Bitmap.
    '
    R = BitBlt(MonoHdc, 0, 0, 32, 32, editor.Pic_Mask.hDC, 0, 0, SRCCOPY)

    ' Obtain a far Pointer to the actual Icons information and Bitmaps
    ' and Lock this information so Windows will not move it.
    '
    Lpicon = GlobalLock(Pic_Ctrl.Picture)

    ' Replace the Icons Mask Bitmap with the new Mask Bitmap.
    '
    R = GetBitmapBits(MonoHbm, 128, Lpicon + 12)

    ' Replace the Icons Image Bitmap with the new Image Bitmap.
    '
    R = GetBitmapBits(editor.Pic_Image.Image, ImageSize, Lpicon + 12 + 128)

    ' Unlock the Icons memory so Windows is free to move it if neccessary
    '
    R = GlobalUnlock(Pic_Ctrl.Picture)

    ' Since VB is unaware of any modifications we make to the Icon using
    ' any API routines, it does not know to redisplay the Icon, so we
    ' must force it to display the new icon.
    '
    Pic_Ctrl.Cls

    ' Set Changed Flag to TRUE since it has been modified.
    '
    If Pic_Ctrl.Tag <> editor.Pic_Undo.Tag Then IconInfo(CurrentIcon).Changed = True

End Sub

' When either the Editors ColorPalette or the ColorPalette Forms
' Color Palette is clicked, this routine is called to set the selected
' color into the Mouse colors, and invoke the ColorPalette Form in
' the case of a Double Click event on the Editors Color Palette.
'
Sub Update_Mouse_Colors (Button, X As Single, Y As Single)
Dim color As Long, SolidColor As Long

    ' The ColorPalettes are a single picture control, so we must calculate
    ' the color selected based on the coordinates of the mouse.
    '
    ColorIndex = Fix(X) + Fix(Y) * 16

    ' Obtain color from color array
    '
    color = Colors(ColorIndex)

    ' VB only supports 16 color mode, so we must obtain the nearest Solid
    ' color to the selected color since the Screen and Inverse colors cannot
    ' be set to dithered colors.
    '
    SolidColor = GetNearestColor(editor.hDC, color)

    If DoubleClicked Then
        '
        ' The Editors ColorPalette was Double Clicked, so reset the Flag
        ' and invoke the ColorPalette Form.
        '
        DoubleClicked = False
        ColorPalette.Show

        ' The ColorPalette Forms initialization is done within the
        ' GotFocus Event for its ColorPalette Picture control, so we
        ' must give that Picture Control the focus.
        '
        ColorPalette.Pic_ColorPalette.SetFocus

    ElseIf editor.Opt_Mouse(SCREEN_COLORS).Value And (color <> SolidColor) Then
        '
        ' An attempt to select a Dithered color into the Screen or Inverse
        ' colors was made, so Prompt the user and do not allow the selection
        '
        MsgBox "Screen and Inverse colors can only be set to solid colors", 16, "Error"
    Else
        ' Obtain the the index of the corresponding mouse Color:
        '   0 - Left Mouse Color
        '   1 - Right Mouse Color
        '   2 - Screen Color
        '   3 - Inverse Screen Color
        '
        Index = editor.Opt_Mouse(SCREEN_COLORS).Value * (-2) + Button - 1

        ' Replace the Mouse color with the new color
        '
        MouseColors(Index) = Colors(ColorIndex)

        ' Changing either the Screen Color or Inverse Screen Color also
        ' changes the other so if either the Screen or Inverse color was
        ' changed, we must change the other to its inverse.
        '
        If Index >= 2 Then
            editor.Pic_Icons(0).PSet (1, 1), MouseColors(Index)
            MouseColors(Abs(Index - 5)) = editor.Pic_Icons(0).Point(1, 1)
            editor.Pic_Icons(0).Cls
        End If
    
        If editor.Opt_Mouse(SCREEN_COLORS).Value Then
            '
            ' The Screen or Inverse Screen color was changed, so we must change
            ' the BackColor of all 6 icons in the StatusArea and the Undo Icon to
            ' the new Screen Color and then redisplay the selected Icon in the
            ' Editing area.
            '
            HighLight_Current_Icon
            For I = 0 To 5
                editor.Pic_Icons(I).BackColor = MouseColors(2)
            Next
            editor.Pic_Undo.BackColor = MouseColors(2)
            Magnify_Icon 0, 0, 31, 31
        End If

    End If

    ' Diplay the New Mouse colors at the Bottom of the StatusArea
    '
    Display_Mouse_Colors

End Sub

' Selecting a new drive from the list of a Drive controls drop
' down list does not generate an error if the drive is not ready,
' so when a new drive is selected, we determine if it is ready
' or not.  This routine validates the selected drive and is use
' by both the SaveFileDlg's and Viewers's Drive control
'
Sub Validate_And_Change_Drives (A_Form As Form)
    
    On Error Resume Next
    Err = False

    ' Invoking the Dir$() function with the selected drive will generate
    ' an error if the drive is not ready.  We don't care about the return
    ' value, we just care if an error is generated or not.
    '
    Dummy$ = Dir$(Left$(A_Form.Drv_DriveList.Drive, 2))

    If Err Then
        '
        ' The drive was not ready, so prompt the user
        '
        Beep
        MsgBox Error$(Err), 16, "IconWorks - ERROR: " + Format$(Err)

        ' Reset the Drive Control back to its previously selected drive
        '
        A_Form.Drv_DriveList.Drive = Left$(A_Form.Dir_DirectoryList.Path, 2)
    Else
        ' The drive is ready, so change to that drive
        '
        ChDrive A_Form.Drv_DriveList.Drive
        A_Form.Dir_DirectoryList.Path = CurDir$
    End If
  
    On Error GoTo 0

End Sub

' When a filespec is entered into either the Viewer's Filename
' TextBox or the SaveFileDlg's Filename TextBox, this routine is
' called to validate the FileSpec.  The name and path, if one is
' given, is validated.  If a valid FileSpec to an actual file is
' entered and the file does not exist, the return value depends
' on which Form called this routine, since a if called from the
' SaveFileDlg a "File Not Found" error is generated but that is
' OK since a file does not have to exist to write to it.  However,
' if called from the Viewer, the same error will be generated but
' in this case the file must exists since the Viewer is wants to
' open the file for editing.
'
Function Validate_FileSpec (AForm As Form, MustExist)
Dim temp As String

    ' Enable error trapping
    '
    On Error GoTo ErrorInSpec

    Validate_FileSpec = False

    ' Check for valid DOS Path and Filenames.
    '
    temp = Dir$(AForm.Txt_FileName.Text)

    ' The following statement does alot.  It the FileSpec contains
    ' a Path, the FileSpec will be parsed and the Path will be assign
    ' to the File ListBox's Path property.  If the FileSpec contains
    ' Wild card characters, it will be assign to the File ListBox's
    ' pattern property.  If the FileSpec contains a valid file name
    ' and the file exists, a Double Click event will automatically be
    ' generated for the File ListBox.  If the File does not exist,
    ' a "File Not Found" error will be generated which we trap.
    '
    AForm.File_FileList.Filename = AForm.Txt_FileName.Text
  
Exit_The_Function:

    ' Turn off error trapping and exit the function
    '
    On Error GoTo 0
    Exit Function

ErrorInSpec:
    If (Err <> FILE_NOT_FOUND) Or ((Err = FILE_NOT_FOUND) And MustExist) Then
        '
        ' An error other than "File Not Found" occured, or the error
        ' "File Not Found" occured and this Function was invoked from
        ' the Viewer which requires the file to exist.
        '
        Beep
        MsgBox Error$(Err), 16, "IconWorks - ERROR: " + Format$(Err)
    Else
        ' The FileSpec entered contain no errors other than maybe
        ' "File Not Found".
        '
        If Err = FILE_NOT_FOUND Then
            ' A Valid filename was entered in the SaveFileDlg which did not exist
            ' so the File Control did not parse the FileSpec for us.  Since the
            ' FileSpec could contain a path specification, force File control
            ' to parse the Filename string for us by changing last character to
            ' an asterisk "*" and assign the modified FileSpec to the File Controls
            ' FileName property.  The asterisk "*" makes the Filename appear as a
            ' FileSpec rather than a Filename to the File ListBox and it will parse
            ' it for us whether there are any matching files or not.  After it has
            ' been parsed, we change the "*" back to its previous value.
            '
            temp = Right$(AForm.Txt_FileName.Text, 1)
            AForm.File_FileList.Filename = Left$(AForm.Txt_FileName.Text, Len(AForm.Txt_FileName.Text) - 1) + "*"
            AForm.Txt_FileName.Text = Left$(AForm.File_FileList.Pattern, Len(AForm.File_FileList.Pattern) - 1) + temp
            
            ' This checks to see that that file name that has been parsed
            ' is a valid DOS file name

             PeriodPos = InStr(1, AForm.Txt_FileName.Text, ".")
             If PeriodPos <> 0 Then
                LeftofPeriod$ = Left$(AForm.Txt_FileName.Text, PeriodPos - 1)
             Else
               LeftofPeriod$ = AForm.Txt_FileName.Text
             End If
             If Len(AForm.Txt_FileName.Text) > 8 Then
                     Resume Exit_The_Function
            End If
            Else
        End If
        Validate_FileSpec = True
    End If
    Resume Exit_The_Function

End Function

' Saves the current icon to disk, and updates the Icon menu and
' Editors title bar with the new Icons filename.
'
Sub Write_Icon_To_File (FullPath As String, FileName As String)
  
    ' Save new Filename and Path information for the Icon
    '
    IconInfo(CurrentIcon).FileName = FileName
    IconInfo(CurrentIcon).FullPath = FullPath
    IconInfo(CurrentIcon).Changed = False

    ' Display the Icons Filename and Path in the Editors Icon menu
    '
    editor.Menu_IconsSelection(CurrentIcon).Caption = "&" + Format$(CurrentIcon + 1) + " - [" + FullPath + "]" + A_TAB + FileName

    ' Display the Icons Filename in the Editors TitleBar
    '
    editor.Caption = "IconWorks Editor: " + Format$(CurrentIcon + 1) + " - " + FileName

    ' Save the Icon to the specified File in the Specified Directory
    '
    If Right$(FullPath, 1) <> "\" Then FullPath = FullPath + "\"
    SavePicture editor.Pic_Icons(CurrentIcon).Picture, FullPath + FileName

End Sub

