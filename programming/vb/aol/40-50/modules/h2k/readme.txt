*******************************
**      Hound2000.Bas        **
**           +               **
**      Help Program         **
*******************************

H2K_HF.HF is a data file for the help file
it must be in the same directory as the .Exe
to work properly.

Bug warnings/notes:
these are just things i have encountered
that i feel should be mentioned.

MemberSearch Freeze:
	Depending on your connection to AOL
	you might encounter an error from AOL
	when you attempt to use the MemberSearch
	subs. When clicking the "More" button, AOL
	might hang for awhile then get an internal
	error. This error will not close AOL, and i
	wrote the sub so that it will intercept the
	error and exit the sub. This problem is from
	AOL and not my sub. If your connection is
	lagging this might happen.

PasswordCrack Delay option:
	This is just a reminder, it's inside the help also.
	if you want to use the passwordcrack sub on a
	TCP/IP connection to AOL. I've been told that
	it goes too fast for this connection and will
	lag your computer. I wrote this sub using a 
	dial-up connection and didn't notice. In the
	tcp/ip case, use the optional parameter for a 
	delay, an optional parameter doesn't have to be
	entered, but can be.  Ex:
	Call PasswordCrack(List1,List2,List3,20)
	for a 20 second delay.
	Call PasswordCrack(List1,List2,List3)
	for no delay.

that's it really
thats a good thing :o)
***The version of Hound2000.Bas that was included in this zip file has been
updated, the Im_On, Im_Off, ImIgnore, and ImUnignore subs have been
revised due to a bug.***