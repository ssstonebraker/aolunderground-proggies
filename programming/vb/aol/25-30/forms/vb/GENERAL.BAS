Option Explicit
'---------------------------------------------------------------------------------
'   Name    :   General.bas
'   Author  :   Peter Wright
'   Notice  :   (c)1994 Wrox Press Ltd
'
' This module contains general code which can be used by all the other forms in
' the program. By using a code module like this it is possible to build up a
' of commonly used code which can be included in any other of your projects
'---------------------------------------------------------------------------------

Sub Add_Drop_Shadow (TheForm As Form, TheControl As Control, nColour As Long)

'----------------------------------------------------------------------------------------------
'   Name    :   Add_Drop_Shadow  => sub procedure
'   Author  :   Peter Wright
'   Notice  :   (c)1994 Wrox Press Ltd, All Rights Reserved
'
'   Parameters
'       TheForm     - This "object variable" tells us the form to draw the shadow on
'       TheControl  - Another "object variable", telling us which control to add a shadow to
'       nColour     - This is a number variable holding the colour to use. Colours are usually
'                     very large numbers so the type of variable is a "long" one, rather than
'                     a normal integer.
'
'   This routine adds a shadow to a control on a form. For a better understanding of what this
'   procedure does you may like to read up on "Long variables" in chapter 3, "Object variables"
'   in chapter 13, and chapter 7 which covers drawing graphics, and describes this procedure
'   in full.
'
'----------------------------------------------------------------------------------------------

    TheForm.FillStyle = 0

    TheForm.Line (TheControl.Left + 100, TheControl.Top + 150)-Step(TheControl.Width, TheControl.Height), nColour, BF

End Sub

