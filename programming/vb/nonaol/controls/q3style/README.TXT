Thanks for dowloading my control!
Some things you should know...
I creadted this with VB 6.0 SP4 so if you don't have SP4 it may not work correctly!  You can dowload it for free at microsoft.com or you can order the CD for about 10 bucks...
Use the INF file to register this, I think so anyway, that should register it to use with design, if it won't work then you can get ahold of me at
MikeMrrymn@juno.com
my web page is at
www.geocities.com/mikemrrymn
you can also find so me information there, if you notice any bugs or glitches email me and I will make a bug list and try to fix them all and send out an update
The version you are downling is version 3.0 first released 2-18-01.


Q3sytle.ocx
	q3buttons
	use the property page to set options or the property window
		button style 0 - 16
		picture change on mouse over or mouse click
		this is all pretty self explainitory
	
	numberdisplay
		use the property page...
		digit count set automatically but you can use it if you want to.
		caption used at design time for sizing purposes only
		use caption at run time like so
		numberdisplay1.caption = "123"
		the caption must be a string if you want to use a variable that is a number do this
		Dim Number as long
		numberdisplay1.caption = Str$(Number&)
		
	Glass
		glass1.glassify me,true
			that will make the form transparent
		glass1.glassify me, false
			that will make the form visible 	
	*Note the glass control isn't perfect it messes up quite often, I am going to post the code for it on a message board that I will be placing on my site soon, so If anyone has any bright ideas or knows what to do to fix it they can let me know.
Also I have included a sample form.. if it doesn't work then open a new exe and add the ocx then add the form...
