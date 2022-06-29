

	Programmer's Toolkit for Internet Mail


	  using Visual Basic to explore SMTP, POP, IMAP, LDAP and DNS



These programs have been written to illustrate the Internet Mail protocols.
They are provided free of charge and unconditionally.  However, they are not
intended for production use, and therefore without warranty or any
implication of support.

You can find an explanation of the concepts behind this code in
the book:  Programmer's Guide to Internet Mail by John Rhoton,
Digital Press 1999.  ISBN: 1-55558-212-5.

For ordering information please see http://www.amazon.com or
you can order directly with http://www.bh.com/digitalpress.


This software was tested on Windows NT, primarily using Microsoft Exchange
and Microsoft Outlook Express.  It has not been full tested for interoperability
so you may need to make modifications if running against different products
or different configurations than the test environment.


The products listed below are merely suggestions.  The list is not exhaustive nor
have these configurations necessarily been tested.


Pre-requisites:

Visual Basic 5.0 - the code has been written with Visual Basic 5.0.  It has
not been tested on other versions.  It should work on higher versions but 
may require some sacrifice in functionality in order to work with earlier
version.

In order to exploit all the functionality you need to have access to the
following servers:

 - SMTP
 - LDAP
 - POP
 - IMAP
 - DNS

The servers should not enforce any restrictions in terms of encryption (TLS,
SSL, NTLM...) and should not restrict the set of IP addresses from which they
accept connections.  Currently this is the most common configuration of most
enterprises although this may change in the future as security becomes more
of an issue.

If you have a BackOffice test environment you can use Microsoft Exchange 5.5 to 
cover the first four protocols or alternatively Lotus Notes Domino Server 5.0.
For DNS you can use Microsoft DNS or a Unix BIND server.

You need a client which can use the following protocols:

 - SMTP
 - LDAP
 - POP
 - IMAP

You can use Microsoft Outlook Express, Microsoft Outlook 2000 or Netscape
Communicator for all four protocols.  If you are mainly interested in SMTP and
POP you can also consider Eudora or Pegasus Mail or Microsoft Outlook 98.



