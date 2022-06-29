VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Welcome to my stuff for my final encrypter.  This is my 3rd
'example of an encrypter.  Version 3.o


'Cya

'SaBrE
Option Explicit

   Private LCW As Integer                 'Length of the CodeWord
   Private LS2E As Integer                 'Length of String to be Encrypted
   Private LAM As Integer                 'Length of Array Matrix
   Private MP As Integer                    'Matrix Position
   Private Matrix As String                  'Starting Matrix
   Private mov1 As String                    'First Part of Replacement String
   Private mov2 As String                    'Second Part of Replacement String
   Private CodeWord As String            'the CodeWord
   Private CWL As String                    'the CodeWord Letter
   Private EncryptedString As String     'String to Return for Encrypt or String to UnEncrypt for UnEncrypt
   Private EncryptedLetter As String     'Storage Variable for Character just Encrypted
   Private strCryptMatrix(97) As String 'Matrix Array
Public Property Let KeyString(sKeyString As String)
    CodeWord = sKeyString
End Property
Public Function Encrypt(mstext As String) As String
    Dim X As Integer                    ' Loop Counter
    Dim Y As Integer                    'Loop Counter
    Dim Z As Integer                     'Loop Counter
    Dim C2E As String                   'Character to Encrypt
    Dim Str2Encrypt As String        'Text from TextBox

    Str2Encrypt = mstext
    LS2E = Len(mstext)
    LCW = Len(CodeWord)
    EncryptedLetter = ""
    EncryptedString = ""

    Y = 1
    For X = 1 To LS2E
        C2E = Mid(Str2Encrypt, X, 1)
        MP = InStr(1, Matrix, C2E, 0)
        CWL = Mid(CodeWord, Y, 1)
        For Z = 1 To LAM
            If Mid(strCryptMatrix(Z), MP, 1) = CWL Then
                EncryptedLetter = Left(strCryptMatrix(Z), 1)
                EncryptedString = EncryptedString + EncryptedLetter
                Exit For
            End If
        Next Z
        Y = Y + 1
        If Y > LCW Then Y = 1
    Next X
    Encrypt = EncryptedString
'this changes the letters, *reminder* this is a poly-encrypter
End Function
Private Sub Class_Initialize()

    Dim W As Integer 'Loop Counter to set up Matrix
    Dim X As Integer     'Loop through Matrix
    
    Matrix = "8x3p5BeabcdfghijklmnoqrstuvwyzACDEFGHIJKLMNOPQRSTUVWXYZ 1246790-.#/\!@$<>&*()[]{}';:,?=+~`^|%_"
    Matrix = Matrix + Chr(13)  'Add Carriage Return to Matrix
    Matrix = Matrix + Chr(10)  'Add Line Feed to Matrix
    Matrix = Matrix + Chr(34)  'Add "
    ' My String used to make Matrix - 8x3p5Be
    ' My String can be any combination that has a character only ONCE.
    ' EACH Letter in the Matrix is Input ONLY once.
    W = 1
    LAM = Len(Matrix)
    strCryptMatrix(1) = Matrix
  
    For X = 2 To LAM ' LAM = Length of Array Matrix
        mov1 = Left(strCryptMatrix(W), 1)   'First Character of strCryptMatrix
        mov2 = Right(strCryptMatrix(W), (LAM - 1))   'All but First Character of strCryptMatrix
        strCryptMatrix(X) = mov2 + mov1  'Makes up each row of the Array
        W = W + 1
    Next X
End Sub

