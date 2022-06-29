Attribute VB_Name = "SUPERNAS1"
' Modulo per l'algoritmo del Pittore.
' Routine di entrata: Pittore


Type Vec_Int
   x As Integer
   Y As Integer
End Type

Type Vertex
   x As Integer
   Y As Integer
   Z As Integer
   Used As Integer
End Type
Public v() As Vertex


Type triadata
   Normal As Vec3
   h As Single
   Color As Integer
End Type


Type Tria
    Anr As Integer
    Bnr As Integer
    Cnr As Integer
    Z As Integer
    PTria As triadata ' Valore Triadata
End Type

Public Triangles() As Tria
Public Ptriangles As Integer

Type TriaNode
    trnr As Integer
    NextNodo As Integer
End Type


Public Const LARGE = 32000 ' // Non maggiore di sqrt(LONG_MAX)
Public Const LARGE1 = 32001 ' Usato solo da LineaNas
Public Const PIdiv180 = 0.0174532
Public Const BIG = 1E+30

Public X__max As Integer ' Coordinate reali della finestra
Public Y__max As Integer ' Coordinate reali della finestra

Public f As Double
Public zfactor As Double
Public rcolormin As Double
Public rcolormax As Double
Public delta As Double
Public zemin As Double
Public zemax As Double
Public xsC As Double
Public ysC As Double
Public XLCreal As Double
Public YLCreal As Double

Public kEye As Integer
Public hK As Integer
Public XLC As Integer
Public YLC As Integer


Public Vt() As Vec3
Public lightvector As Vec3



Public TriaNodeNext() As TriaNode

Public PStart As Integer
Public PEnd As Integer


Sub Complete_Triangles(n As Integer, offset As Integer, nrs_tr() As Trianrs)

'// Completa triangles[offset],..., triangles[offset+n-1].
'// Numeri di vertice: nrs_tr[0],..., nrs_tr[n-1].
'// Questi triangoli appartengono allo stesso poligono. L'equazione
'// del loro piano è nx . x + ny . y + nz . z = h.
  
  
  Dim i As Integer
  Dim Anr As Integer
  Dim Bnr As Integer
  Dim Cnr As Integer
  Dim ZA As Integer
  Dim ZB As Integer
  Dim ZC As Integer
  Dim zmin As Single
  Dim zmax As Single

  Dim nx As Double
  Dim ny As Double
  Dim nz As Double
  Dim ux As Double
  Dim uy As Double
  Dim uz As Double
  Dim vx As Double
  Dim vy As Double
  Dim vz As Double
  Dim factor As Double
  Dim h As Double
  Dim Ax As Double
  Dim Ay As Double
  Dim Az As Double
  Dim Bx As Double
  Dim By As Double
  Dim Bz As Double
  Dim Cx As Double
  Dim Cy As Double
  Dim Cz As Double
  
  Dim p As Integer
  Dim q As triadata
  
'   // Se il poligono è un'approssimazione di una circonferenza, i
'   // primi tre vertici possono giacere quasi sulla stessa linea,
'   // da cui n/2 invece di 0 nell'istruzione for che segue:
  
  For i = n \ 2 To n
     Anr = nrs_tr(i).a
     Bnr = nrs_tr(i).b
     Cnr = nrs_tr(i).C
     If (orienta(Anr, Bnr, Cnr) > 0) Then Exit For
   Next

   ZA = v(Anr).Z
   ZB = v(Bnr).Z
   ZC = v(Cnr).Z

   Az = zEye(ZA)
   Bz = zEye(ZB)
   Cz = zEye(ZC)
   Ax = xScreen(v(Anr).x) * Az
   Ay = yScreen(v(Anr).Y) * Az
   Bx = xScreen(v(Bnr).x) * Bz
   By = yScreen(v(Bnr).Y) * Bz
   Cx = xScreen(v(Cnr).x) * Cz
   Cy = yScreen(v(Cnr).Y) * Cz
   ux = Bx - Ax
   uy = By - Ay
   uz = Bz - Az
   vx = Cx - Ax
   vy = Cy - Ay
   vz = Cz - Az
   nx = uy * vz - uz * vy
   ny = uz * vx - ux * vz
   nz = ux * vy - uy * vx
   h = nx * Ax + ny * Ay + nz * Az
   factor = 1 / Sqr(nx * nx + ny * ny + nz * nz)
   q.Normal.x = nx * factor
   q.Normal.Y = ny * factor
   q.Normal.Z = nz * factor
   q.h = h * factor
   For i = 0 To n - 1
      p = offset + i
      Triangles(p).Anr = nrs_tr(i).a
      Triangles(p).Bnr = nrs_tr(i).b
      Triangles(p).Cnr = nrs_tr(i).C
      Triangles(p).PTria = q
   '   // Sceglie il lato del triangolo per cui Z varia maggiormente;
   '   // Triangles(p).Z si baserà sul punto medio di questo lato:
      zmin = v(Triangles(p).Anr).Z
      zmax = zmin
      ZB = v(Triangles(p).Bnr).Z
      ZC = v(Triangles(p).Cnr).Z
      If (ZB < zmin) Then
         zmin = ZB
      ElseIf (ZB > zmax) Then
         zmax = ZB
      End If
      If (ZC < zmin) Then
         zmin = ZC
      ElseIf (ZC > zmax) Then
         zmax = ZC
      End If
      Triangles(p).Z = (zmin + zmax) / 2
  
   Next

End Sub

Sub DeleteList(Start() As TriaNode)

   Dim p As TriaNode
'   Do While (Start <> Null)
'     p = start;
'     start = start->next;

End Sub

Function Distance(ITria As Integer, x As Integer, Y As Integer) As Double


'// Si considera la linea passante per il punto di osservazione E
'// e il punto (X, Y) del video. Interessa il punto in cui questa
'// linea interseca il triangolo itria. Sarà restituita la
'// coordinata ze di questo punto.

  Static Dist0 As Double
  Static X0 As Integer
  Static Y0 As Integer
  Static PTria0 As Integer
   
  Dim a As Double
  Dim b As Double
  Dim C As Double
  Dim h As Double
  Dim xs As Double
  Dim ys As Double
  
'  Dist0 = 0: X0 = 0: Y0 = 0
   
 ' // Variabili statiche, qualora lo stesso punto (X, Y) sia usato
 ' // per due triangoli consecutivi appartenenti allo stesso
 ' // poligono (e quindi con lo stesso puntatore ptria).
  
  Dim TriaPtr As Integer
  Dim PTria As Integer
  
  TriaPtr = ITria
'  PTria = Triangles(TriaPtr).PTria
  
 '  tria huge* triaptr = triangles + itria; // = &triangles[itria]
 '  triadata huge*ptria=triaptr->ptria;
   
   If (PTria0 <> TriaPtr Or x <> X0 Or Y <> Y0) Then
      a = Triangles(TriaPtr).PTria.Normal.x
      b = Triangles(TriaPtr).PTria.Normal.Y
      C = Triangles(TriaPtr).PTria.Normal.Z
      h = Triangles(TriaPtr).PTria.h
      xs = xScreen(x)
      ys = yScreen(Y)
      Dist0 = h * Sqr(xs * xs + ys * ys + 1) / (a * xs + b * ys + C)
      X0 = x
      Y0 = Y
      PTria0 = TriaPtr
   End If
   
   Distance = Dist0

End Function



Sub DrawWireFrame(Pic As PictureBox)

Dim i As Integer, k As Integer, j As Integer
  
For k = 1 To UBound(FileVertex)
 i = Abs(FileVertex(k).Vert(1))
 If i > 0 Then
    Xl = to_pix(v(i).x)
    Yl = to_pix(v(i).Y)
    x1 = Xl
    y1 = Yl
    For j = 1 To FileVertex(k).Count
        i = FileVertex(k).Vert(j)
        If i > 0 Then
           x = to_pix(v(i).x)
           Y = to_pix(v(i).Y)
           Pic.Line (x1, y1)-(x, Y)
           x1 = x: y1 = Y
        End If
    Next j
 End If
          
Next

   
End Sub

Sub Fill_Triangle(Pic As PictureBox, i As Integer)

' // Riempie il triangolo i
 
 Dim Triangle(2) As CornerRec
 Dim Anr As Integer
 Dim Bnr As Integer
 Dim Cnr As Integer
 
 Anr = Triangles(i).Anr
 Bnr = Triangles(i).Bnr
 Cnr = Triangles(i).Cnr
 
 Triangle(0).x = to_pix(v(Anr).x)
 Triangle(0).Y = to_pix(v(Anr).Y)
 
 Triangle(1).x = to_pix(v(Bnr).x)
 Triangle(1).Y = to_pix(v(Bnr).Y)
 
 Triangle(2).x = to_pix(v(Cnr).x)
 Triangle(2).Y = to_pix(v(Cnr).Y)
 
 Shade% = Triangles(i).PTria.Color

 Call DrawTriangle(Pic, Triangle(), Shade%)

End Sub

Sub FindRange(i As Integer)
   Dim Normal As Vec3
   Normal.x = Triangles(i).PTria.Normal.x
   Normal.Y = Triangles(i).PTria.Normal.Y
   Normal.Z = Triangles(i).PTria.Normal.Z
   Dim rcolor As Single
   rcolor = DotProduct(Normal, lightvector)
   If (rcolor < rcolormin) Then rcolormin = rcolor
   If (rcolor > rcolormax) Then rcolormax = rcolor
End Sub


Function Inside_Triangle(x As Integer, Y As Integer, XA As Integer, YA As Integer, XB As Integer, YB As Integer, xC As Integer, yC As Integer) As Integer


'  // (X, Y) giace sopra o dentro il triangolo ABC?
   
 Inside_Triangle = Orientation(XB - XA, YB - YA, x - XA, Y - YA) >= 0 And _
                  Orientation(xC - XB, yC - YB, x - XB, Y - YB) >= 0 And _
                  Orientation(XA - xC, YA - yC, x - xC, Y - yC) >= 0
   
End Function


Function Int_To_Pix(x As Double)
  Int_To_Pix = (x + hK) / k
End Function

Function IntersectOrizontal(a As Vec_Int, b As Vec_Int, Y As Integer, xxMin As Integer, xxmax As Integer) As Integer


' // Il segmento AB ha dei punti in comune con il
' // segmento orizzontale {(Xmin, Y), (Xmax, Y)}?


Dim XA As Integer
Dim YA As Integer
Dim XB As Integer
Dim YB As Integer
Dim dx As Long
Dim dy As Long
Dim yDx As Long

   XA = a.x
   YA = a.Y
   XB = b.x
   YB = b.Y


   If (YA < Y And YB < Y Or YA > Y And YB > Y) Then
      IntersectOrizontal = 0
      Exit Function
   End If
   
   If (YA = Y And XA >= xxMin And XA <= xxmax Or _
       YB = Y And XB >= xxMin And XB <= xxmax) Then
       IntersectOrizontal = 1
      Exit Function
   End If
   
   If (YA = YB) Then
      IntersectOrizontal = YA = Y And (CLng(XA - xxmax) * (XB - xxmax) < 0 Or CLng(XA - xxMin) * (XB - xxMin) < 0)
      Exit Function
   End If
      
   If (YA > YB) Then
      Swap XA, XB
      Swap YA, YB
   End If
   
   dx = XB - XA
   dy = YB - YA
   XdY = XA * dy + (Y - YA) * dx
   
   IntersectOrizontal = XdY >= xmin * dy And XdY <= xmax * dy

End Function





Function IntersectVertical(a As Vec_Int, b As Vec_Int, x As Integer, yyMin As Integer, yymax As Integer) As Integer

' // Il segmento AB ha dei punti in comune con il
' // segmento verticale {(X, Ymin), (X, Ymax)}?

Dim XA As Integer
Dim YA As Integer
Dim XB As Integer
Dim YB As Integer
Dim dx As Long
Dim dy As Long
Dim yDx As Long

   XA = a.x
   YA = a.Y
   XB = b.x
   YB = b.Y
   
   If (XA < x And XB < x Or XA > x And XB > x) Then
      IntersectVertical = 0
      Exit Function
   End If
 
   If (XA = x And YA >= yyMin And YA <= yymax Or _
       XB = x And YB >= yyMin And YB <= yymax) Then
      IntersectVertical = 1
      Exit Function
   End If
   
   If (XA = XB) Then
      IntersectVertical = XA = x And (CLng(YA - yymax) * (YB - yymax) < 0 Or CLng(YA - yyMin) * (YB - yyMin) < 0)
      Exit Function
   End If
       
   If (XA > XB) Then
      Swap XA, XB
      Swap YA, YB
   End If
   
   dx = XB - XA
   dy = YB - YA
   yDx = YA * dx + (x - XA) * dy
   IntersectVertical = yDx >= yyMin * dx And yDx <= yymax * dx

End Function
Sub LoadVec3(Vettore() As Vec3, i As Integer, x As Double, Y As Double, Z As Double)
    Vettore(i).x = x
    Vettore(i).Y = Y
    Vettore(i).Z = Z
End Sub

Sub Pittore(Pic As PictureBox, Come As Integer)

Dim i As Integer
Dim k As Integer
Dim vertexnr As Integer
Dim maxnpoly As Integer
Dim totntria As Integer
Dim code As Integer
Dim ntr As Integer
Dim Poly() As Integer
Dim nPoly As Integer
Dim k1 As Integer
Dim k2 As Integer
Dim XLMax As Integer
Dim YLMax As Integer
Dim nvertex As Integer
Dim Pnr As Integer
Dim Qnr As Integer
Dim Orient As Integer
Dim testtria(3) As Integer

Dim nrs_tr() As Trianrs ' // Il tipo trianrs è dichiarato in triangol.h
Dim t1 As Single
Dim t2 As Single

Dim fx As Double
Dim fy As Double
Dim rho As Double
Dim Theta As Double
Dim Phi As Double
Dim xs As Double
Dim ys As Double
Dim xe As Double
Dim ye As Double
Dim ze As Double
Dim xsRange As Double
Dim ysRange As Double
Dim xsmin As Double
Dim xsmax As Double
Dim ysmin As Double
Dim ysmax As Double


Dim x As Double
Dim Y As Double
Dim Z As Double

Dim Ch As String
Dim method As String
Dim St As String


rcolormin = BIG
rcolormax = -BIG

AlgoritmoAttivo = 0 ' per Orienta
   
   
   nvertex = MaxVertNr + 1
   ReDim Vt(nvertex)
   
   SetVista rho, Theta, Phi
   lightvector = AssignVec3(-1, 1, 0)
   SetLimitiVista xsmin, xsmax, ysmin, ysmax, nvertex, Vt()
   
'   // Calcola le costanti del video:
  
  xsRange = xsmax - xsmin
  ysRange = ysmax - ysmin

  xsC = 0.5 * (xsmin + xsmax)
  ysC = 0.5 * (ysmin + ysmax)
  k1 = LARGE / (X__max + 1)
  k2 = LARGE / (Y__max + 1)
  kEye = Min2(k1, k2)
  hK = kEye / 2             '  // k = 50, hk = 25 con VGA
  XLMax = kEye * (X__max + 1)
  YLMax = kEye * (Y__max + 1)
  
  ' // Coordinate dei pixel: Xpix = to_pix(X) and Ypix = to_pix(Y)
  
  XLC = XLMax / 2
  YLC = YLMax / 2
  XLCreal = XLC + 0.5
  YLCreal = YLC + 0.5
  fx = XLMax / xsRange
  fy = YLMax / ysRange
  If fx < fy Then
     f = 0.95 * fx
  Else
     f = 0.95 * fy
  End If
  zfactor = LARGE / (zemax - zemin)
   
  ReDim v(nvertex)
  
 '  // Inizializza l'array dei vertici:
   For i = 0 To nvertex - 1
      If (Vt(i).Z < -100000#) Then
         v(i).Used = False        ' V[i] non in uso
      Else
         v(i).Used = True
            
         xs = Vt(i).x / Vt(i).Z
         ys = Vt(i).Y / Vt(i).Z
         
         v(i).x = XLarge(xs)
         v(i).Y = YLarge(ys)
         v(i).Z = ZLarge(Vt(i).Z)
      End If
   Next
   
   If Come = 0 Then
      DrawWireFrame Pic
      Exit Sub
   End If
   
   Erase Vt
   
'   // Trova il numero massimo di vertici in un solo poligono
'   // e il numero totale dei triangoli che non sono
'   // retrosuperfici:
   
maxnpoly = 0
totntria = 0

nPoly = 0
For k = 1 To UBound(FileVertex)
 nPoly = 0
 i = Abs(FileVertex(k).Vert(1))
 If i > 0 Then
    For j = 1 To FileVertex(k).Count
        i = Abs(FileVertex(k).Vert(j))
              
        If i >= nvertex Or Not v(i).Used Then
           MsgBox "Vertice nr." & CStr(i) & " indefinito"
           End
        End If
        If nPoly < 3 Then testtria(nPoly) = i
    
        nPoly = nPoly + 1
    Next j
         
         If (nPoly > maxnpoly) Then maxnpoly = nPoly
         If Not (nPoly < 3) Then  '  // Ignora il segmento 'libero'
            If (orienta(testtria(0), testtria(1), testtria(2)) >= 0) Then totntria = totntria + nPoly - 2
         End If
         
 End If
          
Next
         
         
' =========

  ReDim Triangles(totntria)
  ReDim Poly(maxnpoly)
  ReDim nrs_tr(maxnpoly - 2)


'   // Lettura delle facce dell'oggetto e memorizzazione dei
'   // triangoli:
   
   
For k = 1 To UBound(FileVertex)
         
    nPoly = 0
    For j = 1 To FileVertex(k).Count
        
        i = Abs(FileVertex(k).Vert(j))
        If nPoly = maxnpoly Then
           MsgBox "Errore di programmazione maxnpoly"
           End
        End If
        Poly(nPoly) = i
        nPoly = nPoly + 1
    
    Next j
    
   
    If (nPoly >= 3) Then
    
    Pnr = Abs(Poly(0))
    Qnr = Abs(Poly(1))
    For i = 2 To nPoly - 1
      Orient = orienta(Pnr, Qnr, Abs(Poly(i)))
      If (Orient <> 0) Then Exit For ' // Normalmente, i = 2
    Next
   
   End If
   
    If (Orient >= 0) Then   ' ; // Non Retrosuperficie
   
'      // Suddivisione di un poligono in triangoli:
      
      code = Triangul(Poly(), nPoly, nrs_tr(), Orient)
      If (code > 0) Then
        If (ntr + code > totntria) Then
             MsgBox "Errore di programmazione: totntria"
             End
        End If
        Call Complete_Triangles(code, ntr, nrs_tr())
        ntr = ntr + code
      End If
      
   End If
      
Next k
   
Erase Poly
Erase nrs_tr

   
   For i = ntr - 1 To 0 Step -1
       FindRange i
   Next
   
   ncolors = 12
   delta = 0.999 * (ncolors - 1) / (rcolormax - rcolormin + 0.001)
   
   For i = ntr - 1 To 0 Step -1
     Call Set_Tr_Color(i)
   Next
   
  
     ntr_b% = ntr
     Call Q_Sort(Triangles(), 0, ntr_b%)   '  triangles[0] è il triangolo più vicino
   
     For i = ntr - 1 To 0 Step -1
       Fill_Triangle Pic, i
     Next

End Sub

Function Max2(i As Integer, j As Integer) As Integer
     If i > j Then Max2 = i Else Max2 = j
End Function

Function Max3(i As Integer, j As Integer, k As Integer) As Integer
    Max3 = Max2(i, Max2(j, k))
End Function


Function Min2(i As Integer, j As Integer) As Integer
     If i < j Then Min2 = i Else Min2 = j
End Function

Function Min3(i As Integer, j As Integer, k As Integer) As Integer
        Min3 = Min2(i, Min2(j, k))
End Function

Function orienta(Pnr As Integer, Qnr As Integer, Rnr As Integer) As Integer
  If AlgoritmoAttivo = 0 Then ' Pittore
     orienta = Orientation(v(Qnr).x - v(Pnr).x, v(Qnr).Y - v(Pnr).Y, v(Rnr).x - v(Pnr).x, v(Rnr).Y - v(Pnr).Y)
  Else
     orienta = LOrienta(Pnr, Qnr, Rnr)
  End If
End Function

Function Orientation(u1 As Integer, U2 As Integer, v1 As Integer, v2 As Integer) As Long
   Dim Det As Long

   Det = CLng(u1) * v2 - CLng(U2) * v1
   If Det < -250 Then
      Det = -1
   ElseIf Det > 250 Then
      Det = 1
   End If

   Orientation = Det

End Function

Sub Q_Sort(a() As Tria, Ptr As Integer, n As Integer)
    
 ' Quick Sort
 ' a = Triangles()
 ' Ptr = Puntatore ad a()
 ' n = Num. elemento corrente per il sort


    Dim i As Integer, j As Integer
    Dim x As Tria
    Dim w As Tria

   Do
      i = Ptr
      j = n - 1
      x = a(j / 2)
      Do
         Do While (a(i).Z < x.Z): i = i + 1: Loop
         Do While (a(j).Z > x.Z): j = j - 1: Loop

         If (i < j) Then
              w = a(i)
              a(i) = a(j)
              a(j) = w
          End If
          i = i + 1
          j = j - 1
      Loop While i <= j
          
      If i = j + 3 Then
         i = i - 1
         j = j + 1
      End If

      If j + 1 < n - i Then
         If j > 0 Then Q_Sort a(), 0, j + 1
        ' Ptr = Ptr + i
         n = n - i
       Else
         Pt% = i
         If i < n - 1 Then Q_Sort a(), Pt%, n - i
         n = j + 1
       End If
  
  Loop While n > 1

End Sub

Sub Set_Tr_Color(i As Integer)
   Dim Color As Integer
   Dim rcolor As Double
   Dim Normal As Vec3
   
   Normal.x = Triangles(i).PTria.Normal.x
   Normal.Y = Triangles(i).PTria.Normal.Y
   Normal.Z = Triangles(i).PTria.Normal.Z
   rcolor = DotProduct(Normal, lightvector)
   Color = 1 + (rcolor - rcolormin) * delta
   If (Color < 0) Then MsgBox ("Codice colore negativo")
   If (Color >= 16) Then MsgBox ("Codice colore troppo grande")
'  // (in caso di un errore di programma)
   Triangles(i).PTria.Color = Color
  ' MsgBox Color
End Sub

Sub SetLimitiVista(xsmin As Double, xsmax As Double, ysmin As Double, ysmax As Double, nvertex As Integer, Vt() As Vec3)
 
 Dim PNew As Vec3
 Dim Ve As Vec3
 Dim Vi As Vec3
 Dim Va As Vec3
 
 Dim i As Integer
 Dim k As Integer
   
   For i = 0 To nvertex
      Vt(i).Z = -1000000# ' Non usato
   Next
   
   
   xsmin = BIG
   ysmin = BIG
   zemin = BIG
   xsmax = -BIG
   ysmax = -BIG
   zemax = -BIG
      

For k = 1 To UBound(FileCoord)
      
      i = FileCoord(k).i
      Vi.x = FileCoord(k).x
      Vi.Y = FileCoord(k).Y
      Vi.Z = FileCoord(k).Z
      
  If i > 0 Then
      
      If (i >= nvertex) Then
         MsgBox "Troppi vertici o numero di vertice non legale"
         End
      End If
      
      PNew.x = Vi.x - ObjPoint.x
      PNew.Y = Vi.Y - ObjPoint.Y
      PNew.Z = Vi.Z - ObjPoint.Z
      
      Call Eyecoord(PNew, Ve)
      Va.x = Ve.x
      Va.Y = Ve.Y
      Va.Z = Ve.Z

      If (Va.Z < 0) Then
         MsgBox "Il punto 0 dell'oggetto e un vertice " & Chr(10) & "su lati diversi del punto di osservazione E." & Chr(10) & "Provare con un valore maggiore per rho."
         Exit Sub
      End If

      xs = Va.x / Va.Z
      ys = Va.Y / Va.Z

      If (xs < xsmin) Then xsmin = xs
      If (xs > xsmax) Then xsmax = xs
      If (ys < ysmin) Then ysmin = ys
      If (ys > ysmax) Then ysmax = ys
      If (Va.Z < zemin) Then zemin = Va.Z
      If (Va.Z > zemax) Then zemax = Va.Z
      Vt(i) = Ve
  
  End If

Next k
      
If (xsmin = BIG) Then
 MsgBox "File di input non corretto"
 End
End If


End Sub

Sub SetVista(rho As Double, Theta As Double, Phi As Double)
   
   ObjPoint = AssignVec3(0.5 * (xmin + xmax), 0.5 * (ymin + ymax), 0.5 * (zmin + zmax))
   rho = xmax - xmin
   If (ymax - ymin > rho) Then rho = ymax - ymin
   If (zmax - zmin > rho) Then rho = zmax - zmin
   rho = rho * 3
   Theta = 20
   Phi = -65

   Call Coeff(rho, Theta * PIdiv180, Phi * PIdiv180)

End Sub

Sub Swap(x As Integer, Y As Integer)

   Dim t As Integer
   t = x
   x = Y
   Y = t
   
End Sub

Function to_pix(x As Integer) As Integer

'  // Arrotondamento e conversione
 to_pix = (x + hK) / kEye


End Function


Function XLarge(xs As Double) As Integer
    XLarge = Int(XLCreal + f * (xs - xsC))
End Function

Function xScreen(x As Integer) As Double
   xScreen = xsC + (x - XLC) / f
End Function

Function YLarge(ys As Double) As Integer
     YLarge = Int(YLCreal + f * (ys - ysC))
End Function

Function yScreen(Y As Integer) As Double
     yScreen = ysC + (Y - YLC) / f
End Function

Function zEye(Z As Integer) As Double
       zEye = Z / zfactor + zemin
End Function

Function ZLarge(ze As Single) As Integer
   ZLarge = Int((ze - zemin) * zfactor + 0.5)
End Function

