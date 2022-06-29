Attribute VB_Name = "sfdCommonDialog1"
Option Explicit
'thanks for downloading my bas file.
'the coding below is the basics behind using common dialog
'it includes: printing, color change, font change,
'save, and opening files.
'they aren't the best methods of doing it but shows you
'some of the things that you need to know to write your
'own subs for common dialog.
'take the following code and paste it into a common dialog
'ocx. this ocx is made by microsoft and can be obtained
'by downloading the service packs for vb5 or 6.
'a few things:
'where ever you see CommonDialog1, is the name of your
'common dialog control. i advise leaving it the
'defalt CommonDialog1. it's just easier to work with.
'also the rtfbox1 is a rich text box.
'you need to use them if your dealing with saving fonts,
'colors, size changes, etc.
'feel free to use this code in whatever your making.
'the next version will inculde updated code and
'forms showing you how to use the code.
'|----------------------------------------------------------|
'|created by: SFD                                           |
'|sfd@sfdteam2g.zzn.com                                     |
'|http://www.sfd.com-us.com                                 |
'|http://www.sfd.com-us.com/journal.html                    |
'|http://www.hellmouth.com-us.com                           |
'|----------------------------------------------------------|

Sub ColorChange_Text()
    CommonDialog1.Flags = cdlCCFullOpen
    CommonDialog1.ShowColor
    rtfBox1.SelColor = CommonDialog1.Color
End Sub

Sub FontChange_Text()
    CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
    CommonDialog1.ShowFont
    
    With rtfBox1
        .SelFontName = CommonDialog1.FontName
        .SelFontSize = CommonDialog1.FontSize
        .SelBold = CommonDialog1.FontBold
        .SelItalic = CommonDialog1.FontItalic
        .SelStrikeThru = CommonDialog1.FontStrikethru
        .SelUnderline = CommonDialog1.FontUnderline
        .SelColor = CommonDialog1.Color
    End With
End Sub

Sub FilePrint()
    On Error Resume Next
    If rtfBox1 Is Nothing Then Exit Sub
    With CommonDialog1
        .DialogTitle = "Print - Print Now"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If rtfBox1.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            rtfBox1.SelPrint .hDC
        End If
    End With
End Sub

Sub OpenFile(Filename As String)
    On Error GoTo 10
    With CommonDialog1
    .Filter = "Text Files (*.txt)|*.txt|Rich Text Files (*.rtf)|*.rtf|Active Docs (*.wri)|*.wri|All Files (*.*)|*.*|"
    .ShowOpen
    If UCase(Right$(.Filename, 3)) = "RTF" Or UCase(Right$(.Filename, 3)) = "WRI" Then
        tmode = rtfRTF
    Else
        tmode = rtfText
    End If
    rtfBox1.LoadFile CommonDialog1.Filename, tmode
    End With
10  Exit Sub
End Sub

Sub SaveFile()
    On Error GoTo 10
    With CommonDialog1
    .Filter = "Active Docs (*.wri)|*.wri|Rich Text File (*.rtf)|*.rtf|All Files (*.*)|*.*"
    .ShowSave
    rtfBox1.SaveFile .Filename
    End With
10  Exit Sub
End Sub
