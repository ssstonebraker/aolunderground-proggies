??, o  g    
SearchMail"& ' ? -< Search Mail >-??? ?  ?  o  /   B #????$Form25?  6?  7o  8/  ?>   Panel3D3?SSPanel?      Xo?
 Panel3D3  ?  ?}   List1?SSList !" #$?&?' (,Y?D- 6?x x ?	    )?/5 	                                                   ? ?>   Panel3D2?SSPanel?      _
g Panel3D2  ?  ?>   
Command3D2?	SSCommand?    (x ?wCancel      ?>   
Command3D1?	SSCommand?    x x ?wSearch      ?>   Panel3D1?SSPanel?        g  Panel3D1  ?  ?E   Text1???? x x /hString to searchMS Sans Serif  A  ?,  ?  __	 
?P?P?   ?   ? ?  r  Z   F? ?? ?? ?   ?g    ? ?n ? ? ??:                  ?  |  Command3D2_Click
 
SearchMail   Command3D1_Click? aol55 
FindWindow? aolI mdi   FindChildByClass^ welc   FindChildByTitle?  	option3D1? Value? Option1 	option3D2Z Option2* 	Option3D3? Option3 X   SetFocusAPI5 RunMenuString? 	Form_Load   success? SetWindowPos* Main? hWnd? HWND_TOPMOST   FLAGS   
Command3D1   Font3D? 
Command3D2   nmail   ao? treebxR TreeSize   SendMessageByNumu LB_GETCOUNT   p   getmailz   MailLen? LB_GETTEXTLEN   mailname? TreeName? SendMessagebyString   
LB_GETTEXT   List1   Selected   Panel3d2   Caption? 	ListCount   LSTTXT   List   TXTSTR   text1   Text    	  ????????     Command3D1_Click 4?      X  |       5?  ? ??? > ???   5 ?  ? ? ? ? ? ?? ? ?? ? I ?   , 5 ?  ? 8    5N ??   9 	  ????????	     Command3D2_Click 4       X  Z       n    9 	  ????????    	 Form_Load 4?     X  Z     ?  g   n  ? ?? ? ? ?  ?  rg  ? ? ?  ? ? ?  ?  l  AOL Frame25 ?       ? ?   ? ?  ? 	 MDIClient   ? ?   ? ?  ? 	 Welcome,    ? ?   ? ? ? I f  ?  $ You must sign on first. ?  ?  H Not signed on &  = ?  2 ?  ?  p AOL Frame25 ?       ? ?   ? ?  ?	 MDIClient   ? ?   ? ?  ? New Mail    ? ?  ?? ? I l  ?  ( Load your mail box first. ?  ?  N Box not found &  = ?  2 ?  ?  v AOL Frame25 ?       ? ?  ??  ?	 _AOL_TREE   ? ?  ? 
? ?   ? ?  ??    ?  ??? > ???   ??    ? 5 ?   ? *   *?? ? ? F  ? u F  ^ R   ?e F  x   N ??    ? ? ? ?  8   8   9 	  ????????"   ?