Attribute VB_Name = "fStream"

' Modulo per la lettura dei solidi in Input.
' Il modello è rappresentato nei files
' *.DAT


Type FCoord
  i As Integer              ' Numero Vertice
  x As Double
  Y As Double
  Z As Double
End Type

Type FVertex                ' Superfice
     Count As Integer       ' Numero di Vertici
     Vert(100) As Integer   ' Puntatori a FCoord
End Type

Public FileCoord() As FCoord
Public FileVertex() As FVertex

Public MaxVertNr As Integer
Public MinVertNr As Integer

Sub GetVertexFromLine(St As String, FV As FVertex)
         
 ' Preleva i numeri di Vertice da un File di Win3D (L. Ammeraal)

 Dim j As Integer
 Dim VaS As String
 Dim VaN As Integer
 Dim b As Integer
 Dim Ch As String * 1
   
   For j = 1 To Len(St)
       Ch = Mid$(St, j, 1)
       If Ch <> " " Then
          VaS = VaS + Ch
       Else
          If Len(VaS) > 0 Then
               VaN = Val(VaS)
               b = b + 1
               FV.Vert(b) = Val(VaS)
               VaS = ""
          End If
      End If
  Next j
   
   
' Completa l'ultimo Vertice

    If Len(VaS) > 0 Then
       b = b + 1
       FV.Vert(b) = Val(VaS)
    End If
    
    FV.Count = b

End Sub

Function LoadFile(File As String) As Integer
  
  Dim St As String
  Dim nn As Integer
  Dim Facce As Integer
  Dim i As Integer
  Dim x As Double
  Dim Y As Double
  Dim Z As Double
  Dim Pl As Integer
  Dim m As Integer
  Dim Ps As Integer
  Dim Vrt
  
  Erase FileCoord
  Erase FileVertex
  
  LoadFile = True
  
  On Error Resume Next
  nn = FreeFile
  Open File For Input As nn
  
  If Err <> 0 Then
     LoadFile = False
     Exit Function
  End If
  
  OpenFile = nn
  
 On Error GoTo 0

 ReDim FileCoord(1)

 Do Until EOF(nn)
   
   Line Input #nn, St
   If Mid$(St, 1, 6) = "Faces:" Then
      Facce = True
      Line Input #nn, St
   End If
      
   If Not Facce Then
      Vrt = Vrt + 1
      Call GetCoordFromLine(St, i, x, Y, Z)
      If St = "FILE NON VALIDO" Then
         LoadFile = False
         Exit Function
      End If
      If Vrt > UBound(FileCoord) Then ReDim Preserve FileCoord(Vrt)
      FileCoord(Vrt).i = i
      FileCoord(Vrt).x = x
      FileCoord(Vrt).Y = Y
      FileCoord(Vrt).Z = Z
   Else
      
         Pl = Pl + 1
         ReDim Preserve FileVertex(Pl)
         GetVertexFromLine St, FileVertex(Pl)
         
   End If


Loop
   
Close nn%


SetLimits

End Function

Sub GetCoordFromLine(St As String, i As Integer, x As Double, Y As Double, Z As Double)
   On Error Resume Next
 ' Preleva le coordinate da un File di Win3D (L. Ammeraal)

 Dim j As Integer
 Dim VaS As String
 Dim VaN As Double
 Dim b As Integer
   
   For j = 1 To Len(St)
       Ch = Mid$(St, j, 1)
       If Ch <> " " Then
          VaS = VaS + Ch
       Else
          If Len(VaS) > 0 Then
               VaN = Val(VaS)
               b = b + 1
               Select Case b
                 Case 1
                    i = VaN
                 Case 2
                    x = VaN
                 Case 3
                    Y = VaN
                 Case 4
                    Z = VaN
              End Select
              VaS = ""
          End If
      End If
  Next j
   
   
   
' Completa la Z

    If Len(VaS) > 0 Then Z = Val(VaS)

    If i + x + Y + Z = 0 Then St = "FILE NON VALIDO"
   
End Sub


Sub SetLimits()
 
' Ritorna il numero massimo di Vertici

Dim i As Integer
Dim k As Integer
Dim x As Double
Dim Y As Double
Dim Z As Double

' assegna il numero di vertici totale
' e le dimensioni min,max dell'oggetto

xmin = BIG
xmax = -BIG
ymin = BIG
ymax = -BIG
zmax = -BIG
zmin = BIG

For k = 1 To UBound(FileCoord)

    i = FileCoord(k).i
    x = FileCoord(k).x
    Y = FileCoord(k).Y
    Z = FileCoord(k).Z
    
    If (i > MaxVertNr) Then MaxVertNr = i
    If (i < MinVertNr) Then MinVertNr = i
    If (x < xmin) Then xmin = x
    If (x > xmax) Then xmax = x
    If (Y < ymin) Then ymin = Y
    If (Y > ymax) Then ymax = Y
    If (Z < zmin) Then zmin = Z
    If (Z > zmax) Then zmax = Z

Next

End Sub


