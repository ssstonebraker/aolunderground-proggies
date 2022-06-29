VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Splatter painting with PSet"
   ClientHeight    =   5820
   ClientLeft      =   1110
   ClientTop       =   2115
   ClientWidth     =   7365
   Height          =   6225
   Left            =   1050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   Top             =   1770
   Width           =   7485
End
Option Explicit

Sub Form_Load ()

    Dim nXCoord As Integer
    Dim nYCoord As Integer

    Dim nRed As Integer
    Dim nGreen As Integer
    Dim nBlue As Integer

    Dim nIndex As Integer

    ' First, we show the form to make it appear on the screen, before the
    ' load event is finished.
    frmMain.Show

    ' The Doevents method is described in full in chapter 14 - it lets windows get on with
    ' updating the screen and responding to our events, before our code kicks in.
    DoEvents

    ' The randomize method tells VB to get ready to produce some random numbers.
    ' See chapter 4, variables and data types for more information.
    Randomize

    For nIndex = 1 To 2000

        ' First we need to work out some random coordinates to plot the dot at
        nXCoord = Int(Rnd(1) * frmMain.ScaleWidth)
        nYCoord = Int(Rnd(1) * frmMain.ScaleHeight)
    
        ' Then, random values for the red, green and blue settings of the dot's colour are
        ' decided upon.
        nRed = Int(Rnd(1) * 255)
        nGreen = Int(Rnd(1) * 255)
        nBlue = Int(Rnd(1) * 255)

        ' Finally, the point is actually drawn, using the Pset routine
        PSet (nXCoord, nYCoord), RGB(nRed, nGreen, nBlue)

    Next

    MsgBox "All done!"

End Sub

