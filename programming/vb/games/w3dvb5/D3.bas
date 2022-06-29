Attribute VB_Name = "D3"
' Modulo per la rappresentazione e il
' calcolo dei vettori tridimensionali

'NB.: Non tutte le Sub e Function definite in questo
'     modulo sono effettivamente utilizzate dal programma.
'     Sono comunque presenti per facilitare eventuali
'     implementazioni


Type Vec3
    x As Single
    Y As Single
    Z As Single
End Type

' Variabili in Coeff

Public v11 As Double
Public v12 As Double
Public v13 As Double
Public v21 As Double
Public v22 As Double
Public v23 As Double
Public v32 As Double
Public v33 As Double
Public v43 As Double


' Variabili Globali Tridimensionali

Public ObjPoint As Vec3

Sub Coeff(rho As Double, Theta As Double, Phi As Double)
 
 Dim costh As Double
 Dim sinth As Double
 Dim cosph As Double
 Dim sinph As Double
   
'   Angoli in radianti:
   
 costh = Cos(Theta)
 sinth = Sin(Theta)
 cosph = Cos(Phi)
 sinph = Sin(Phi)
 v11 = -sinth
 v12 = -cosph * costh
 v13 = -sinph * costh
 v21 = costh
 v22 = -cosph * sinth
 v23 = -sinph * sinth
 v32 = sinph
 v33 = -cosph
 v43 = rho
 
End Sub

Function CopyVec3(v As Vec3) As Vec3
         CopyVec3 = v
End Function

Function AssignVec3(x As Double, Y As Double, Z As Double) As Vec3
    AssignVec3.x = x
    AssignVec3.Y = Y
    AssignVec3.Z = Z
End Function

Function DotProduct(a As Vec3, b As Vec3) As Double

   DotProduct = a.x * b.x + a.Y * b.Y + a.Z * b.Z

End Function

Sub Eyecoord(pw As Vec3, pe As Vec3)


  pe.x = v11 * pw.x + v21 * pw.Y
  pe.Y = v12 * pw.x + v22 * pw.Y + v32 * pw.Z
  pe.Z = v13 * pw.x + v23 * pw.Y + v33 * pw.Z + v43


End Sub

Function SommaVec3(u As Vec3, v As Vec3) As Vec3
   SommaVec3.x = u.x + v.x
   SommaVec3.Y = u.Y + v.Y
   SommaVec3.Z = u.Z + v.Z
End Function

Function IncrVec3(u As Vec3, v As Vec3) As Vec3
   u.x = u.x + v.x
   u.Y = u.Y + v.Y
   u.Z = u.Z + v.Z
   IncrVec3 = u
End Function

Function DecrVec3(u As Vec3, v As Vec3) As Vec3
   u.x = u.x - v.x
   u.Y = u.Y - v.Y
   u.Z = u.Z - v.Z
   DecrVec3 = u
End Function

Function MultIncVec3(v As Vec3, C As Double) As Vec3
   v.x = C * v.x
   v.Y = C * v.Y
   v.Z = C * v.Z
   MultIncVec3 = v
End Function

Function MoltiplicaVec3(C As Double, v As Vec3) As Vec3
   MoltiplicaVec3.x = C * v.x
   MoltiplicaVec3.Y = C * v.Y
   MoltiplicaVec3.Z = C * v.Z
End Function

Function SottraiVec3(u As Vec3, v As Vec3) As Vec3
   SottraiVec3.x = u.x - v.x
   SottraiVec3.Y = u.Y - v.Y
   SottraiVec3.Z = u.Z - v.Z
End Function

