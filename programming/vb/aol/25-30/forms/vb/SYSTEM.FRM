VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Using GetProfileString to search Win.INI"
   ClientHeight    =   3465
   ClientLeft      =   1110
   ClientTop       =   1500
   ClientWidth     =   7365
   Height          =   3870
   Left            =   1050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   7365
   Top             =   1155
   Width           =   7485
   Begin PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3075
      Left            =   60
      ScaleHeight     =   3045
      ScaleWidth      =   7155
      TabIndex        =   3
      Top             =   195
      Width           =   7185
      Begin CommandButton cmdFind 
         Caption         =   "&Search Win.Ini for the specified values"
         Default         =   -1  'True
         Height          =   450
         Left            =   180
         TabIndex        =   2
         Top             =   2385
         Width           =   6690
      End
      Begin Frame Frame1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Find what in Win.Ini"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   1590
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   6705
         Begin TextBox txtKeyName 
            Height          =   285
            Left            =   3120
            TabIndex        =   1
            Top             =   900
            Width           =   3150
         End
         Begin TextBox txtAppname 
            Height          =   285
            Left            =   3120
            TabIndex        =   0
            Top             =   495
            Width           =   3150
         End
         Begin Label Label2 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Key Name (Entry Name)"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   75
            TabIndex        =   7
            Top             =   975
            Width           =   2580
         End
         Begin Label Label1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Application (Section) Name"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   105
            TabIndex        =   6
            Top             =   525
            Width           =   2970
         End
      End
      Begin Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reading WIN.INI"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   15
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   450
         Left            =   2265
         TabIndex        =   4
         Top             =   90
         Width           =   2745
      End
   End
End
Option Explicit

' We need to declare the GetProfileString call here to let us read values from the
' Win.Ini file later in the code.
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer


Sub cmdFind_Click ()

    ' First, declare a string variable to hold the returned string from GetProfileString. This
    ' should be a fixed length string to cope for all eventualities, 4000 characters should be
    ' fine.
    Dim sReturnString As String * 4000

    ' Now, declare a varaible to hold the integer return code from GetProfileString.
    Dim nReturnCode As Integer

    ' Now declare two strings, one to hold the message string that we are building, the other
    ' to hold a character count so that we can work through the list of possible keynames
    ' that GetProfileString might return.
    Dim sMessage As String, nCharacter As Integer

    ' First we need to do some error checking. The user must at least enter a section name
    ' to search, if nothing else.
    If txtAppName.Text = "" Then
        MsgBox "You must enter a section name to search.", , "Missing Parameter"
        Exit Sub
    End If

    ' When you call GetProfileString you can either pass a specific keyname to the
    ' call to attempt to find a match, or you can pass a long 0 to the call to list
    ' all keys in the specified section. We need to check whether or not the user has
    ' entered a valid keyname, in order to pass the right value

    If txtKeyName.Text = "" Then
        nReturnCode = GetProfileString(txtAppName.Text, 0&, "NONE", sReturnString, 4000)
    Else
        nReturnCode = GetProfileString(txtAppName.Text, CStr(txtKeyName.Text), "NONE", sReturnString, 4000)
    End If
    
    ' Check to see if the matching section and or keyname were found. The return string will be
    ' equal to "NONE" if either was not.
    If Left$(sReturnString, 4) = "NONE" Then
        MsgBox "Either the application name or the keyname could not be found.", , "Search Failed"
        Exit Sub
    End If

    ' Now we have to interpret the results
    ' If no Keyname was specified then we now have a string with a list of keynames in, each
    ' separated by a character of number 0. We need to get these names and put them into a string
    ' to display using a messagebox.
    If txtKeyName.Text = "" Then

        sMessage = "Valid keynames are : " & Chr$(10)

        For nCharacter = 1 To nReturnCode
            If Mid$(sReturnString, nCharacter, 1) <> Chr$(0) Then
                sMessage = sMessage & Mid$(sReturnString, nCharacter, 1)
            Else
                sMessage = sMessage & ", "
            End If
        Next

        MsgBox sMessage, , "Valid keynames"

    ' ...on the other hand, if a keyname was specified then we need to display the line from
    ' the INI file that represents that keyname here - easy!
    Else

        sMessage = "The entry for " & txtKeyName.Text & " is : " & Chr(10)
        sMessage = sMessage & Left$(sReturnString, nReturnCode)
        MsgBox sMessage, , "Key found"

    End If

End Sub

Sub txtAppname_GotFocus ()
    
    ' When the text box receives focus, the text in it should highlight
    txtAppName.SelStart = 0
    txtAppName.SelLength = Len(txtAppName.Text)

End Sub

Sub txtKeyName_GotFocus ()
    
    ' When the text box receives focus, the text in it should highlight
    txtKeyName.SelStart = 0
    txtKeyName.SelLength = Len(txtKeyName.Text)

End Sub

