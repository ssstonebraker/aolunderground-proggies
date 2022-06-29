Attribute VB_Name = "modWebLaunch"
   '*******************************************************
   '* Sup all, thanks for trying one of my examples.This  *
   '* will show you how to make an fake hyperlink in your *
   '* program.                                            *
   '*                                                     *
   '* First make a label and set the properties for       *
   '* the font to underlined and the forecolor to be      *
   '* blue. This will give the appearance of a hyperlink  *
   '* that people are used to.                            *
   '*                                                     *
   '* To change the mouse pointer while over your label   *
   '* goto the labels property-mouseicon and direct the   *
   '* dialog box to your cursor that you want. Then goto  *
   '* the mousepointer property and choose (99-custom).   *
   '*                                                     *
   '* Then in the objects click event call to the         *
   '* following sub 'WebLaunch'.                          *
   '* An example:                                         *
   '*                                                     *
   '*  Private Sub Label1_Click()                         *
   '*      Call WebLaunch("http://www.knk2000.com")       *
   '*  End Sub                                            *
   '*                                                     *
   '* Thats all there is to it. Just add this sub into    *
   '* your current bas or form. Or just add this whole    *
   '* bas into your project.                              *
   '*                                                     *
   '* But if your really interested in knowing how the    *
   '* sub works. Don't cut and paste, spend the time to   *
   '* type out every word. Ya'll get a better             *
   '* understanding of why something was coded that way.  *
   '*                                                     *
   '* On parts of the code that ya don't understand, try  *
   '* changing it and then test it. Then ya should        *
   '* atleast see why that particular code was needed.    *
   '*                                                     *
   '* Big thanks go out to KNK, Dos, & Tko. I learned     *
   '* alot from these people and the time they spent      *
   '* making web sites, examples, help files and          *
   '* answering questions is appreciated.                 *
   '*                                                     *
   '* Any questions can be sent to me at:                 *
   '*           NightShadeXX@hotmail.com                  *
   '*                                                     *
   '* Tip of the day: Try printing out a bas file and     *
   '* reading it when your not busy.                      *
   '*                                                     *
   '* Make an example of your own and submitt it to KnK   *
   '* or another web site. Small or large, someone will   *
   '* benefit by it.                                      *
   '*                                                     *
   '*                   -NightShade                       *
   '*                    08/09/1999                       *
   '*******************************************************


Option Explicit

Public Sub WebLaunch(YourURL As String)

    Dim FileNumber As Integer
    
    'This will write a temp file containing the
    'shortcut information. Then launch your default
    'web browser and have it read and assign the
    'url according to the temp file
    
    'gets the next free number of files being used
    FileNumber% = FreeFile

    'open a new file
    Open App.Path & "\Temp.URL" For Output As #FileNumber%
    'write to the new file
    Print #FileNumber%, "[InternetShortcut]"
    Print #FileNumber%, "URL=" & YourURL$
    'close up the file
    Close #FileNumber%
    
    'Launch the web browser and set the url by the file
    'that you just made
    Shell "rundll32.exe shdocvw.dll,OpenURL " _
          & App.Path & "\temp.url", vbNormalFocus
          
    'delete the temp file
    Kill App.Path & "\Temp.URL"
    
End Sub
