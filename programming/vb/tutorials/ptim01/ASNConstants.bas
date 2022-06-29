Attribute VB_Name = "ASNConstants"
'========================================================================
' Copyright 1999 - Digital Press, John Rhoton
'
' This program has been written to illustrate the Internet Mail protocols.
' It is provided free of charge and unconditionally.  However, it is not
' intended for production use, and therefore without warranty or any
' implication of support.
'
' You can find an explanation of the concepts behind this code in
' the book:  Programmer's Guide to Internet Mail by John Rhoton,
' Digital Press 1999.  ISBN: 1-55558-212-5.
'
' For ordering information please see http://www.amazon.com or
' you can order directly with http://www.bh.com/digitalpress.
'
'========================================================================


Public Const ASN_SCOPE_MASK = 192

Public Const ASN_UNIVERSAL = 0
Public Const ASN_APPLICATION = 64
Public Const ASN_CONTEXT_SPECIFIC = 128
Public Const ASN_PRIVATE = 192

Public Const ASN_COMPOSITE = 32

Public Const ASN_BOOLEAN_TAG = 1
Public Const ASN_INTEGER_TAG = 2
Public Const ASN_OCTETSTRING_TAG = 4
Public Const ASN_ENUMERATED_TAG = 10
Public Const ASN_SEQUENCE_TAG = 16 + ASN_COMPOSITE  ' 48
Public Const ASN_SET_TAG = 17 + ASN_COMPOSITE  ' 49

