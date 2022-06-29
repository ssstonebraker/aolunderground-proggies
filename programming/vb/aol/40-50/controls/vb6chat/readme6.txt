***********************************************************************
*                                                                     *
*            ;;;;;;;;;;;;¸¸      ¸;;;;¸      ¸¸;;;;;;¸                *
*            ;;;;;;;;;;;;;;;    ;;´  ´;;    ;;;;´´´;;;;¸              *
*            ;;;;       ´;;;;   ;;¸  ¸;;    ;;;;¸   ´´´´              *
*            ;;;;        ´;;;;   ´;;;;´      ;;;;;¸¸                  *
*          ;;;;;;;;;;;    ;;;;              ;;;´;;;;;¸¸               *
*          ´´;;;;´´´´´    ;;;;             ;;;;  ´´;;;;;¸             *
*            ;;;;         ;;;;             ;;;;;¸   ´;;;;             *
*            ;;;;       ¸;;;;               ;;;;;;¸  ;;;;             *
*            ;;;;;;;;;;;;;;;                  ´;;;;;;;;´              *
*            ;;;;;;;;;;;;´´                      ´;;;;;¸              *
*                                           ¸¸¸¸    ;;;;              *
*                                           ´;;;;¸¸¸;;;´              *
*                                             ´;;;;;;´´               *
*                                                                     *
***********************************************************************
*                   *** DoS's AoL4 Chat OCX 2.0 ***                   *
***********************************************************************
*                        *** DISCLAIMER ***                           *
*  This Active X control is designed to read text from the AOL4 chat  *
*  window. It can send chat to the window as well. However, this tool *
*  is simply a demonstration only. Its purpose is to show that the    * 
*  chat window can be monitored without the use of subclassing. Do not*
*  use this control to cause problems on AOL as it may be a violation *
*  of AOL's Terms Of Service.                                         *
***********************************************************************
*              *** What you can do with the control ***               *
*  As of 9/25/95, this control has only four properties, methods, or  *
*  events. The following is a breif discription of those;             *
*	1. you can turn it on	-	Chat1.ScanOn                  *
*	2. you can turn it off	-	Chat1.ScanOff                 *
*	3. you can send to chat	-	Chat1.SendChat (Text1.Text)   *
*	4. it fires an event with the screen name called "Screen_Name"*
*  and what they said called "What_Said". The event is below.         *
*                                                                     *
*Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)*
*                                                                     *
*End Sub                                                              *
*                                                                     *
*  That event is fired each and every time the chat changes. It will  *
*  give you two strings, "Screen_Name" and "What_Said". Their value   *
*  will be the newest chat text.                                      *
***********************************************************************
*                          *** EXAMPLE ***                            *
*  If you wanted to make an echo bot to test this, you will only      *
*  need a form with a textbox and 2 command buttons. We'll call the   *
*  text box txtName and the command buttons cmdOn and cmdOff.         *
*                                                                     *
*  Private Sub cmdON_Click()                                          *
*      Chat1.ScanOn 'tell the ocx to start reading the chat           *
*  End Sub                                                            *
*                                                                     *
*  Private Sub cmdOFF_Click()                                         *
*      Chat1.scanoff 'tell the ocx to stop reading the chat           *
*  End Sub                                                            *
*                                                                     *
*Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)*
*      If LCase(Screen_Name) = LCase(txtName.Text) Then               *
*      'check to see if the screen name is the same                   *
*      'as the name in our textbox                                    *
*          Chat1.ChatSend What_Said                                   *
*          'if it is, we'll chat send what they said                  *
*      End If                                                         *
*End Sub                                                              *
***********************************************************************
*                      *** About the control ***                      *
*  This control was designed with Visual Basic 6, and is intended for *
*  use with Visual Basic 6. Also, the control is designed to read the *
*  chat from AOL's software, version 4 for Windows 95.                *
*                                                                     *
*  After using this control, you may notice that it "skips" chat lines*
*  from time to time. This is intentional. The control is designed to *
*  ignore chat lines which have a length greater than 250 characters. *
*  Also, the control may ignore scrolls of the same text from the same*
*  person.                                                            *
***********************************************************************
*                          *** CONTACT ***                            *
*  Included in this zip is the version6.txt. This text contains the   *
*  version number for this control.                                   *
*  If you have any questions, comments, or bug reports, feel free to  *
*  contact me at xdosx@hotmail.com.                                   *
***********************************************************************
