Attribute VB_Name = "Lineanas"

' Modulo per l'eliminazione delle linee nascoste.
'
' ATTENZIONE - La Sub LineNas è la routine principale
'              di questo modulo e (sigh!) non funziona.
'              Comunque altre routine e variabili presenti
'              sono indispensabili per l'algoritmo del pittore
'              che è funzionante. Ergo, non eliminate il modulo
'              dal progetto! Non funzionerà più niente!


Public AlgoritmoAttivo As Integer ' 0 - Pittore, 1 - Linenas (serve per Orienta)

Public Const Nscreen = 10  '  // Ci saranno Nscreen x Nscreen quadrati
Public density As Double

Public d As Double
Public c1 As Double
Public c2 As Double
Public xfactor As Double
Public yfactor As Double

Public Xrange As Double
Public Yrange As Double
Public Xvp_range As Double
Public Yvp_range As Double

Public xmin As Double
Public xmax As Double
Public ymin As Double
Public ymax As Double
Public zmin As Double
Public zmax As Double

Public deltax As Double
Public deltay As Double
Public denom As Double
'Public zemin As Double
'Public zemax As Double

Public eps1 As Double
Public trset() As Integer
Public dummy As Integer
Public vertexcount As Integer

Public x_center As Double
Public y_center As Double
Public r_max As Double
Public x_max As Double
Public y_max As Double
Public x_min As Double
Public y_min As Double

Type Vertexes
     Vt As Vec_Int
     Z As Double
     Connect(5) As Integer
End Type

Public VV() As Vertexes
Public pVertex As Integer

Type Nodo
  idx As Integer
  jtr As Integer
  nextn As Integer
End Type

Public VScreen(Nscreen, Nscreen) As Nodo

Type Point
  Pntscr As Vec_Int
  zPnt As Double
  nrPnt As Integer
End Type

Type linked_stack
    p As Point
    q As Point
    k0 As Integer
    nextn As Integer
End Type
    
Public stptr(1) As linked_stack


Sub add_linesegment(Pr As Integer, Qr As Integer)
Dim iaux As Integer
Dim p As Integer
Dim i As Integer
Dim n As Integer
Dim Pt(3) As Integer
Dim p_old(3) As Integer
Dim Pnr As Integer
Dim Qnr As Integer

 Pnr = Pr
 Qnr = Qr
  

   If (Pnr > Qnr) Then
      iaux = Pnr
      Pnr = Qnr
      Qnr = iaux
   End If
   
 ' Ora: Pnr < Qnr
   p = VV(Pnr).Connect(0)
   If (p = 0) Then
       VV(Pnr).Connect(0) = 1
       VV(Pnr).Connect(1) = Qnr
       Exit Sub
   End If
   
   n = VV(Pnr).Connect(0)
   For i = 1 To n
      If VV(Pnr).Connect(i) = Qnr Then Exit Sub ' Già nella lista
   Next i
   
   n = n + 1 ' Ora Q deve essere posto in p[n]
   If (n Mod 3 = 0) Then
      p_old(0) = VV(Pnr).Connect(0)
      p_old(1) = VV(Pnr).Connect(1)
      p_old(2) = VV(Pnr).Connect(2)
    
    ' Blocchi di tre interi
      For i = 1 To n - 1
          VV(Pnr).Connect(i) = p_old(i)
      Next
      VV(Pnr).Connect(0) = n
      VV(Pnr).Connect(n) = Qnr '  // n è un multiplo di 3
                               '  // *p=n, p[1],..., p[n] usati
                               '  // (p[n+1], p[n+2] liberi)
   Else
      VV(Pnr).Connect(0) = n
      VV(Pnr).Connect(n) = Qnr ' // n non è un multiplo di 3 (e n > 1)
   End If


End Sub

Function ColNr(x As Integer) As Integer
         ColNr = (CLng(x) * Nscreen) / LARGE1
End Function

Sub dealwithlinkedstack()

Dim Pt As linked_stack
Dim p As Point
Dim q As Point
Dim k0 As Integer
Dim Ptr As Integer

Ptr = 1
Do While Ptr <> 0
    Pt = stptr(Ptr)
    p = Pt.p
    q = Pt.q
    k0 = Pt.k0
    Ptr = Pt.nextn
    linesegment Form1.Pict, p, q, k0
Loop


End Sub


Sub LineNas(Pic As PictureBox)

Dim i As Integer
Dim Pnr As Integer
Dim Qnr As Integer
Dim ii As Integer
Dim vertexnr As Integer
Dim Ptr As Integer
Dim iconnect As Integer
Dim code As Integer
Dim ntr As Integer
Dim i_i As Integer
Dim j_j As Integer
Dim jtop As Integer
Dim jbot As Integer
Dim jI As Integer
Dim trnr As Integer
Dim jtr As Integer
Dim Poly() As Integer
Dim nPoly As Integer
Dim iLeft As Integer
Dim iRight As Integer
Dim nvertex As Integer
Dim ntrset As Integer
Dim maxntrset As Integer
Dim VLOWER(Nscreen) As Integer
Dim VUPPER(Nscreen) As Integer
Dim Orient As Integer
Dim maxnpoly As Integer
Dim totntria As Integer
Dim testtria(3) As Integer

Dim xsmin As Double
Dim xsmax As Double
Dim ysmin As Double
Dim ysmax As Double


Dim nrs_tr() As Trianrs

Dim deltax As Long
Dim deltay As Long

Dim rho As Double
Dim Theta As Double
Dim Phi As Double
Dim x As Double
Dim Y As Double
Dim Z As Double
Dim xe As Double
Dim ye As Double
Dim ze As Double
Dim xx As Double
Dim yy As Double
Dim fx As Double
Dim fy As Double
Dim Xcenter As Double
Dim Ycenter As Double

Dim Ps As Vec_Int
Dim Qs As Vec_Int
Dim vLeft As Vec_Int
Dim vRight As Vec_Int

Dim p As Vec3

Dim pNode As Integer

minvertex = 32000
maxntrset = 400
AlgoritmoAttivo = 1 ' Per Funct. Orienta

Erase stptr


   nvertex = MaxVertNr + 1
   ReDim Vt(nvertex)
   
   SetVista rho, Theta, Phi
   SetLimitiVista xsmin, xsmax, ysmin, ysmax, nvertex, Vt()

' Da InitGr
   
   x_max = 10
   density = X__max / (x_max - x_min)
   y_max = y_min + Y__max / density
   x_center = 0.5 * (x_min + x_max)
   y_center = 0.5 * (y_min + y_max)

   zfactor = LARGE / (zemax - zemin)
   eps1 = 0.001 * (zemax - zemin)
   
'   // Calcola le costanti del video:
   
   Xrange = xsmax - xsmin
   Yrange = ysmax - ysmin
   
   Xvp_range = x_max - x_min
   Yvp_range = y_max - y_min
   fx = Xvp_range / Xrange
   fy = Yvp_range / Yrange
   If fx < fy Then
      d = 0.95 * fx
   Else
      d = 0.95 * fy
   End If
   
   Xcenter = 0.5 * (xsmin + xsmax)
   Ycenter = 0.5 * (ysmin + ysmax)
   c1 = x_center - d * Xcenter
   c2 = y_center - d * Ycenter
   deltax = Xrange / Nscreen
   deltay = Yrange / Nscreen
   
   xfactor = LARGE / Xrange
   yfactor = LARGE / Yrange
   
   
   
   ReDim VV(nvertex)
   
' Inizializza l'array dei vertici:
   
   For i = 0 To nvertex
      If Vt(i).Z < -100000# Then
         Erase VV(i).Connect
      Else
         Erase VV(i).Connect
         VV(i).Vt.x = xIntScr(Vt(i).x / Vt(i).Z, xsmin)
         VV(i).Vt.Y = yIntScr(Vt(i).Y / Vt(i).Z, ysmin)
         VV(i).Z = Vt(i).Z
  '       MsgBox "x= " & VV(i).Vt.X & "y= " & VV(i).Vt.Y & "z= " & VV(i).Z
       End If
  Next i
  
  Erase Vt

' Trova il numero massimo di vertici in un solo poligono
' e il numero totale dei triangoli che non sono
' retrosuperfici:
   
maxnpoly = 0
totntria = 0
         
nPoly = 0
For k = 1 To UBound(FileVertex)
 nPoly = 0
 i = Abs(FileVertex(k).Vert(1))
 If i > 0 Then
    For j = 1 To FileVertex(k).Count
        i = Abs(FileVertex(k).Vert(j))
              
        If i >= nvertex Then
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
          
Next k
         
         
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
    
   
  '  If (nPoly = 1) Then
      '  MsgBox "Solo un vertice del poligono?"
      '  End
  '  End If
    
    If nPoly = 2 Then
      Call add_linesegment(Poly(0), Poly(1))
    Else
    
       Pnr = Abs(Poly(0))
       Qnr = Abs(Poly(1))
       For s = 2 To nPoly - 1
          Orient = LOrienta(Pnr, Qnr, Abs(Poly(s)))
          If (Orient <> 0) Then Exit For ' // Normalmente, s = 2
       Next
   
    
       If (Orient >= 0) Then   ' ; // Non Retrosuperficie
   
          For s = 1 To nPoly
             i_i = s Mod nPoly
             code = Poly(i_i)
             vertexnr = Abs(code)
             If code < 0 Then
                Poly(i_i) = vertexnr
              Else
                Call add_linesegment(Poly(s - 1), vertexnr)
             End If
          Next s
   
   
'      // Suddivisione di un poligono in triangoli:
      
      
          code = Triangul(Poly(), nPoly, nrs_tr(), Orient)
          If (code > 0) Then
             If (ntr + code > totntria) Then
                  MsgBox "Errore di programmazione: totntria"
                  End
             End If
             Call LComplete_Triangles(code, ntr, nrs_tr())
             ntr = ntr + code
          End If
      
       End If
    End If
    
Next k
   
Erase Poly
Erase nrs_tr

 Call setupscreenlist(Triangles, ntr)
 ReDim trset(maxntrset)
   

 '  // Traccia tutti i segmenti finché sono visibili:
   
  For Pnr = MinVertNr To MaxVertNr
      Ptr = VV(Pnr).Connect(0)
      
'    0: Pnr non in uso; NULL: nessun segmento memorizzato
'    Pnr non in uso oppure nessun segmento memorizzato
      
      If Ptr > 0 Then
         Ps = VV(Pnr).Vt
         For iconnect = 1 To Ptr
             Qnr = VV(Pnr).Connect(iconnect)
             Qs = VV(Qnr).Vt
              
      '   Usando le liste video, si costruirà l'insieme
      '   dei triangoli che possono nascondere i punti di PQ:
          
             If (Ps.x < Qs.x Or (Ps.x = Qs.x And Ps.Y < Qs.Y)) Then
                vLeft = Ps: vRight = Qs
             Else
                vLeft = Qs: vRight = Ps
             End If
             iLeft = ColNr(vLeft.x)
             iRight = ColNr(vRight.x)
         
             If (iLeft <> iRight) Then
                 deltay = vRight.Y - vLeft.Y
                 deltax = vRight.x - vLeft.x
             End If
         
             jbot = RowNr(vLeft.Y)
             jtop = jbot
             
             For ii = iLeft To iRight
                If ii = iRight Then
                   jI = RowNr(vRight.Y)
                Else
                   hh& = vLeft.Y + (xCoord(ii + 1) - vLeft.x) * deltay / deltax
                   If hh& > 32000 Then
                      jI = Nscreen
                   Else
                      jI = RowNr(CInt(hh&))
                   End If
                End If
             
                VLOWER(ii) = Min2(jbot, jI)
                jbot = jI
                VUPPER(ii) = Max2(jtop, jI)
                jtop = jI
             
             Next ii
         
         Next iconnect
         
         ntrset = 0
         For i = iLeft To iRight
             For j = VLOWER(i) To VUPPER(i)
                 pNode = VScreen(i, j).idx
              '   Do While pNode <> 0
                    trnr = VScreen(i, j).jtr
                    '  /* Il triangolo trnr sarà memorizzato solo se
                    '     non è già presente nell'array trset (insieme dei
                    '     triangoli)
                    '  */
                     trset(ntrset) = trnr   '  /* sentinella */
                     jtr = 0
                     Do While trset(jtr) <> trnr: jtr = jtr + 1: Loop
                     If (jtr = ntrset) Then
                         ntrset = ntrset + 1 ' // Significa che trnr è memorizzato
                         If (ntrset = maxntrset) Then
                          '  P = jtr
                            maxntrset = maxntrset + 200
                            ReDim Preserve trset(maxntrset)
                     
                            For s = 0 To ntrset - 1
                                trset(s) = trset(jtr + s)
                            Next s
                         End If
                     End If
                    
                     pNode = VScreen(i, j).nextn
                    
               '  Loop ' Pnode <> 0
             
             Next j
'         // Ora trset[0],..., trset[ntrset-1] è l'insieme dei
'         // triangoli che possono nascondere i punti di PQ.
         Call linesegment(Pic, SetPoint(Ps, VV(Pnr).Z, Pnr), SetPoint(Qs, VV(Qnr).Z, Qnr), ntrset)
         
         dealwithlinkedstack
         
         Next i
      
      End If ' Ptr = 0
      
   Next Pnr

End Sub

Sub LComplete_Triangles(n As Integer, offset As Integer, nrs_tr() As Trianrs)

' Completa triangles[offset],..., triangles[offset+n-1].
' Numeri di vertice: nrs_tr[0],..., nrs_tr[n-1].
' Questi triangoli appartengono allo stesso poligono. L'equazione
' del loro piano è nx . x + ny . y + nz . z = h.
  
  
  Dim i As Integer
  Dim Anr As Integer
  Dim Bnr As Integer
  Dim Cnr As Integer
  Dim ZA As Integer
  Dim ZB As Integer
  Dim ZC As Integer
 ' Dim zmin As Single
 ' Dim zmax As Single

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

   ZA = VV(Anr).Z
   ZB = VV(Bnr).Z
   ZC = VV(Cnr).Z

   Az = zFloat(ZA)
   Bz = zFloat(ZB)
   Cz = zFloat(ZC)
   Ax = xFloat(VV(Anr).Vt.x) * Az
   Ay = yFloat(VV(Anr).Vt.Y) * Az
   Bx = xFloat(VV(Bnr).Vt.x) * Bz
   By = yFloat(VV(Bnr).Vt.Y) * Bz
   Cx = xFloat(VV(Cnr).Vt.x) * Cz
   Cy = yFloat(VV(Cnr).Vt.Y) * Cz
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
   Next

End Sub


Sub linesegment(Pic As PictureBox, p As Point, q As Point, k0 As Integer)


'   Si deve tracciare il segmento PQ, finché non viene nascosto
'   dai triangoli trset[0],..., trset[k0-1].
   Dim Ps As Vec_Int
   Dim Qs As Vec_Int
   Dim Ass As Vec_Int
   Dim Bs As Vec_Int
   Dim Cs As Vec_Int
   Dim Temp As Vec_Int
   Dim Iss As Vec_Int
   Dim Js As Vec_Int
   
   
   Dim x1 As Single
   Dim x2 As Single
   Dim y1 As Single
   Dim y2 As Single
   Dim xP As Double
   Dim yP As Double
   Dim xQ As Double
   Dim yQ As Double
   Dim zP As Double
   Dim zQ As Double
   Dim xI As Double
   Dim yI As Double
   Dim hP As Double
   Dim hQ As Double
   Dim xJ As Double
   Dim yJ As Double
   Dim lam_min As Double
   Dim lam_max As Double
   Dim lambda As Double
   Dim mu As Double
   Dim hh As Double
   Dim h1 As Double
   Dim h2 As Double
   Dim zI As Double
   Dim zJ As Double
   Dim ZA As Double
   Dim ZB As Double
   Dim ZC As Double
   Dim zmaxPQ As Double
   
   
   Dim Pnr As Integer
   Dim Qnr As Integer
   Dim kk As Integer
   Dim j As Integer
   Dim Anr As Integer
   Dim Bnr As Integer
   Dim Cnr As Integer
   Dim i As Integer
   Dim Poutside As Integer
   Dim Qoutside As Integer
   Dim Pnear As Integer
   Dim Qnear As Integer
   
   Dim APB As Integer
   Dim AQB As Integer
   Dim BPC As Integer
   Dim BQC As Integer
   Dim CPA As Integer
   Dim CQA As Integer
   Dim xminPQ As Integer
   Dim xmaxPQ As Integer
   Dim yminPQ As Integer
   Dim ymaxPQ As Integer
   Dim X_P As Integer
   Dim Y_P As Integer
   Dim X_Q As Integer
   Dim Y_Q As Integer
   Dim u1 As Integer
   Dim U2 As Integer
   
   Dim denom As Long
   Dim v1 As Long
   Dim v2 As Long
   Dim w1 As Long
   Dim w2 As Long
   
   Dim Normal As Vec3
   
   Ps = p.Pntscr
   Qs = q.Pntscr
   
   zP = p.zPnt
   zQ = q.zPnt
   
   Pnr = p.nrPnt
   Qnr = q.nrPnt
   
   X_P = Ps.x
   Y_P = Ps.Y
   X_Q = Qs.x
   Y_Q = Qs.Y
   u1 = X_Q - X_P
   U2 = Y_Q - Y_P
   kk = k0
   
   
   If (X_P < X_Q) Then
      xminPQ = X_P
      xmaxPQ = X_Q
   Else
      xminPQ = X_Q
      xmaxPQ = X_P
   End If
   If (Y_P < Y_Q) Then
      yminPQ = Y_P
      ymaxPQ = Y_Q
   Else
      yminPQ = Y_Q
      ymaxPQ = Y_P
   End If
   
   Do While (kk > 0)
      kk = kk - 1
      j = trset(kk)
      Anr = Triangles(j).Anr
      Bnr = Triangles(j).Bnr
      Cnr = Triangles(j).Cnr

 '  Test 1 (3D): PQ è uno dei lati del triangolo?
      If ((Pnr = Anr Or Pnr = Bnr Or Pnr = Cnr) And _
          (Qnr = Anr Or Qnr = Bnr Or Qnr = Cnr)) Then GoTo Continua
      
      Ass = VV(Anr).Vt
      Bs = VV(Bnr).Vt
      Cs = VV(Cnr).Vt

    '   Test 2 (2D): I test minimax:
      If (xmaxPQ <= Ass.x And xmaxPQ <= Bs.x And xmaxPQ <= Cs.x Or _
          xminPQ >= Ass.x And xminPQ >= Bs.x And xminPQ >= Cs.x Or _
          ymaxPQ <= Ass.Y And ymaxPQ <= Bs.Y And ymaxPQ <= Cs.Y Or _
          yminPQ >= Ass.Y And yminPQ >= Bs.Y And yminPQ >= Cs.Y) Then GoTo Continua
             '  continue; // continue significa: 'visibile'

  ' Test 3 (2D): P e Q giacciono in un semipiano definito da
  ' un lato del triangolo ABC (e sono esterni a questo
  ' triangolo)?
      
      APB = orientv(Ass, Ps, Bs)
      AQB = orientv(Ass, Qs, Bs)
      If (APB + AQB > 0) Then GoTo Continua
      BPC = orientv(Bs, Ps, Cs)
      BQC = orientv(Bs, Qs, Cs)
      If (BPC + BQC > 0) Then GoTo Continua
      CPA = orientv(Cs, Ps, Ass)
      CQA = orientv(Cs, Qs, Ass)
      If (CPA + CQA > 0) Then GoTo Continua

  ' Test 4 (2D): A, B e C giacciono sullo stesso semipiano
  '               definito da PQ?:
      If (Abs(orientv(Ps, Qs, Ass) + orientv(Ps, Qs, Bs) + orientv(Ps, Qs, Cs)) > 1) Then GoTo Continua

  ' Test 5 (3D): Sono sia zP che zQ minori di zA, zB e zC?
      
      ZA = VV(Anr).Z
      ZB = VV(Bnr).Z
      ZC = VV(Cnr).Z
      If zP > zQ Then zmaxPQ = zP Else zmaxPQ = zQ
      
      If (zmaxPQ <= ZA And zmaxPQ <= ZB And zmaxPQ <= ZC) Then GoTo Continua

  ' Test 6 (3D): E' vero che né P né Q giacciono dietro il
  '               piano ABC?
      
      Normal = Triangles(j).PTria.Normal
      hh = Triangles(j).PTria.h
      If (hh = 0) Then GoTo Continua    ' Il piano passa per il punto di
                                        ' osservazione
      xP = zP * xFloat(X_P)
      yP = zP * yFloat(Y_P)
      xQ = zQ * xFloat(X_Q)
      yQ = zQ * yFloat(Y_Q)
      hP = Normal.x * xP + Normal.Y * yP + Normal.Z * zP
      hQ = Normal.x * xQ + Normal.Y * yQ + Normal.Z * zQ
      h2 = hh + eps1
      If (hP <= h2 And hQ <= h2) Then GoTo Continua

  ' Test 7 (2D) Il triangolo ABC oscura completamente PQ?
  
      Poutside = APB = 1 Or BPC = 1 Or CPA = 1
      Qoutside = AQB = 1 Or BQC = 1 Or CQA = 1
      If (Not Poutside And Not Qoutside) Then Exit Sub

  ' Nessuna delle precedenti istruzioni continue è stata
  ' eseguita, per cui il segmento PsQs ha dei punti in comune
  ' con il triangolo AsBsCs.
      
      h1 = hh - eps1
      Pnear = hP < h1
      Qnear = hQ < h1
      If (Pnear And Not Poutside Or Qnear And Not Qoutside) Then GoTo Continua
  ' Ora P giace fuori dalla piramide EABC oppure dietro
  ' il triangolo ABC, e lo stesso vale per Q.

  ' Ora sono calcolati i punti di intersezione:
      lam_min = 1#
      lam_max = 0#
      For i = 0 To 2
      
         v1 = Bs.x - Ass.x
         v2 = Bs.Y - Ass.Y
         w1 = Ass.x - xP
         w2 = Ass.Y - yP
         denom = u1 * v2 - U2 * v1
         If (denom <> 0) Then       ' PsQs non parallelo ad AsBs
            mu = (U2 * w1 - u1 * w2) / CDbl(denom)
            ' mu = 0 dà A, e mu = 1 dà B.
            If (mu > -0.0001 And mu < 1.0001) Then
               lambda = (v2 * w1 - v1 * w2) / CDbl(denom)
               ' lambda = PI/PQ (I è il punto di intersezione)
               If (lambda > -0.0001 And lambda < 1.0001) Then
                  If (Poutside <> Qoutside And _
                  lambda > 0.0001 And lambda < 0.9999) Then
                     lam_min = lam_max = lambda
                     Exit For     '  Un solo punto di intersezione
                  End If
                  If (lambda < lam_min) Then lam_min = lambda
                  If (lambda > lam_max) Then lam_max = lambda
               End If ' lambda ...
            End If ' mu > -...
         End If ' Denom <> 0
         Temp = Ass
         Ass = Bs
         Bs = Cs
         Cs = Temp
      Next i
      
   '  Test 8: I e J sono punti di intersezione.
   '  Verifica se questi punti giacciono di fronte al
   '  triangolo ABC:
      
      If (Poutside And lam_min > 0.01) Then
         Iss.x = Int(xP + lam_min * u1 + 0.5)
         Iss.Y = Int(yP + lam_min * U2 + 0.5)
         zI = 1 / (lam_min / zQ + (1 - lam_min) / zP)
         xI = zI * xFloat(Iss.x)
         yI = zI * yFloat(Iss.Y)
         If (Normal.x * xI + Normal.Y * yI + Normal.Z * zI) < h1 Then GoTo Continua
         Call stack_linesegment(SetPoint(Ps, zP, Pnr), SetPoint(Iss, zI, -1), kk)
      End If ' POutside..
      If (Qoutside And lam_max < 0.99) Then
         Js.x = Int(xP + lam_max * u1 + 0.5)
         Js.Y = Int(yP + lam_max * U2 + 0.5)
         zJ = 1 / (lam_max / zQ + (1 - lam_max) / zP)
         xJ = zJ * xFloat(Js.x)
         yJ = zJ * yFloat(Js.Y)
         If (Normal.x * xJ + Normal.Y * yJ + Normal.Z * zJ) < h1 Then GoTo Continua
         
         Call stack_linesegment(SetPoint(Qs, zQ, Qnr), SetPoint(Js, zJ, -1), kk)
      End If ' Qoutside..
      
      Exit Sub  '  Solo se non è stata eseguita nessuna istruzione
                '  Goto Continua
Continua:

   Loop ' While (kk > 0)
   
   x1 = SetX(d * xFloat(X_P) + c1)
   y1 = SetY(d * yFloat(Y_P) + c2)
   x2 = SetX(d * xFloat(X_Q) + c1)
   y2 = SetY(d * yFloat(Y_Q) + c2)
   
   Pic.Line (x1, y1)-(x2, y2)
   
'   Ms = " x1= " & x1 & Chr(10)
'   Ms = Ms & " y1= " & y1 & Chr(10)
'   Ms = Ms & " x2= " & x2 & Chr(10)
'   Ms = Ms & " y2= " & y2
'   MsgBox Ms
   
  ' move(d * xfloat(XP) + c1, d * yfloat(YP) + c2);
  ' draw(d * xfloat(XQ) + c1, d * yfloat(YQ) + c2);


End Sub


Function LOrienta(Pnr As Integer, Qnr As Integer, Rnr As Integer) As Integer
 Dim Ps As Vec_Int
 Dim Qs As Vec_Int
 Dim Rs As Vec_Int
 
 Dim u1 As Integer
 Dim U2 As Integer
 Dim v1 As Integer
 Dim v2 As Integer
 
 Dim Det As Long
 
 Ps = VV(Pnr).Vt
 Qs = VV(Qnr).Vt
 Rs = VV(Rnr).Vt
 
 u1 = Qs.x - Ps.x
 U2 = Qs.Y - Ps.Y
 v1 = Rs.x - Ps.x
 v2 = Rs.Y - Ps.Y
 
 Det = CLng(u1) * v2 - CLng(U2) * v1
 
 If Det < -300 Then
    LOrienta = -1
 Else
    LOrienta = Abs(Det > 300)
 End If
 
End Function

Function orientv(Ps As Vec_Int, Qs As Vec_Int, Rs As Vec_Int) As Integer
 Dim u1 As Integer
 Dim U2 As Integer
 Dim v1 As Integer
 Dim v2 As Integer

 Dim Det As Long
  
 u1 = Qs.x - Ps.x
 U2 = Qs.Y - Ps.Y
 v1 = Rs.x - Ps.x
 v2 = Rs.Y - Ps.Y

 Det = CLng(u1) * v2 - CLng(U2) * v1
   
 If Det < -10 Then
    orientv = -1
 Else
    orientv = Det > 10
 End If

End Function

Function RowNr(Y As Integer) As Integer
         RowNr = (CLng(Y) * Nscreen) / LARGE1
End Function


Function SetPoint(p As Vec_Int, Z As Double, nr As Integer) As Point
   SetPoint.Pntscr = p
   SetPoint.zPnt = Z
   SetPoint.nrPnt = nr
End Function

Sub setupscreenlist(Tr() As Tria, n As Integer)

'   Predispone le liste dei triangoli indicati con TR[0],...,
'   TR[n-1].
 Dim i As Integer
 Dim l As Integer
 Dim j As Integer
 Dim iMin As Integer
 Dim iMax As Integer
 Dim j_old As Integer
 Dim jI As Integer
 Dim topcode(2) As Integer
 Dim iLeft As Integer
 Dim iRight As Integer
 Dim LLOWER(Nscreen) As Integer
 Dim LUPPER(Nscreen) As Integer
 
 Dim deltax As Long
 Dim deltay As Long

 Dim Ass As Vec_Int
 Dim Bs As Vec_Int
 Dim Cs As Vec_Int
 Dim vLeft(2) As Vec_Int
 Dim vRight(2) As Vec_Int
 Dim Aux As Vec_Int

 Dim p As Integer
 Dim p_New As Integer
 Dim p_old As Integer
 
 '  tria huge*p;
 '  node huge*p_new, huge*p_old;
   
 For i = 0 To n - 1
     p = i
     Ass = VV(Tr(p).Anr).Vt
     Bs = VV(Tr(p).Bnr).Vt
     Cs = VV(Tr(p).Cnr).Vt
     
   '  MsgBox "Triangolo: " & i
   '  Ms = "Ass.x= " & Ass.X & Chr(10)
   '  Ms = Ms & "Ass.y= " & Ass.Y & Chr(10)
   '  MsgBox Ms
   '  Ms = "Bs.x= " & Bs.X & Chr(10)
   '  Ms = Ms & "Bs.y= " & Bs.Y & Chr(10)
   '  MsgBox Ms
   '  Ms = "Cs.x= " & Cs.X & Chr(10)
   '  Ms = Ms & "Cs.y= " & Cs.Y & Chr(10)
   '  MsgBox Ms
     
     topcode(0) = Ass.x > Bs.x  ' // Per l'orientamento positivo
     topcode(1) = Cs.x > Ass.x
     topcode(2) = Bs.x > Cs.x
     vLeft(0) = Ass
     vRight(0) = Bs
     vLeft(1) = Ass
     vRight(1) = Cs
     vLeft(2) = Bs
     vRight(2) = Cs
     For l = 0 To 2 '  // l = numero di lati del triangolo
        If (vLeft(l).x > vRight(l).x Or _
           (vLeft(l).x = vRight(l).x And vLeft(l).Y > vRight(l).Y)) Then
             Aux = vLeft(l)
             vLeft(l) = vRight(l)
             vRight(l) = Aux
        End If
     Next l
        
     iMin = ColNr(Min3(Ass.x, Bs.x, Cs.x))
     iMax = ColNr(Max3(Ass.x, Bs.x, Cs.x))
      
     'iMin = Max2(iMin, 0)
      
     For ii = iMin To iMax
         LLOWER(ii) = 32000
         LUPPER(ii) = -32000
     Next ii
       
     For l = 0 To 2
         iLeft = ColNr(vLeft(l).x)
         iRight = ColNr(vRight(l).x)
         If (iLeft <> iRight) Then
           deltay = vRight(l).Y - vLeft(l).Y
           deltax = vRight(l).x - vLeft(l).x
         End If
         j_old = RowNr(vLeft(l).Y)
         For ii = iLeft To iRight
             If ii = iRight Then
                jI = RowNr(vRight(l).Y)
             Else
                g& = vLeft(l).Y + CLng(xCoord(ii + 1) - vLeft(l).x * deltay / deltax)
                If g& > 32000 Then g& = 32000
                If g& < 0 Then g& = 0
                jI = RowNr(CInt(g&))
             End If
            If topcode(l) Then
               LUPPER(ii) = Max3(j_old, jI, LUPPER(ii))
            Else
               LLOWER(ii) = Min3(j_old, jI, LLOWER(ii))
            End If
            j_old = jI
         Next ii
      
      Next l
      
'      // Per la colonna I del video, il triangolo è associato solo
'      // con i rettangoli delle righe LOWER[I],...,UPPER[I].
      
       For ii = iMin To iMax
          For j = LLOWER(ii) To LUPPER(ii)
  
             p_New = p_New + 1
  
             p_old = VScreen(ii, j).idx
  
             VScreen(ii, j).idx = p_New
             VScreen(ii, j).jtr = i
             VScreen(ii, j).nextn = p_old
  
          Next j
       Next ii
 Next i


'{  // Predispone le liste dei triangoli indicati con TR[0],...,
'   // TR[n-1].
'   int i, l, I, J, Imin, Imax, j_old, jI, topcode[3],
'      ileft, iright, LOWER[Nscreen], UPPER[Nscreen];
'   long deltax, deltay;
'   vec_int As, Bs, Cs, Left[3], Right[3], Aux;
'   tria huge*p;
'   node huge*p_new, huge*p_old;
'   for (i=0; i<n; i++)
'   {  p = TR + i;
'      As = V[p->Anr].VT; Bs = V[p->Bnr].VT; Cs = V[p->Cnr].VT;
'      topcode[0] = As.X > Bs.X; // Per l'orientamento positivo
'      topcode[1] = Cs.X > As.X;
'      topcode[2] = Bs.X > Cs.X;

'      Left[0] = As; Right[0] = Bs;
'      Left[1] = As; Right[1] = Cs;
'      Left[2] = Bs; Right[2] = Cs;
'      for (l=0; l<3; l++)  // l = numero di lati del triangolo
'         if (Left[l].X > Right[l].X ||
'         (Left[l].X == Right[l].X && Left[l].Y > Right[l].Y))
'         {  Aux = Left[l]; Left[l] = Right[l]; Right[l] = Aux;
'         }
'      Imin = colnr(min3(As.X, Bs.X, Cs.X));
'      Imax = colnr(max3(As.X, Bs.X, Cs.X));
'      for (I = Imin; I<=Imax; I++)
'      {  LOWER[I] = INT_MAX; UPPER[I] = INT_MIN;
'      }
'      for (l=0; l<3; l++)
'      {  ileft = colnr(Left[l].X); iright = colnr(Right[l].X);
'         if (ileft != iright)
'         { deltay = Right[l].Y - Left[l].Y;
'           deltax = Right[l].X - Left[l].X;
'         }
'         j_old = rownr(Left[l].Y);
'         for (I=ileft; I<=iright; I++)
'         {  jI = (I == iright ? rownr(Right[l].Y) : rownr(Left[l].Y
'                  + (Xcoord(I+1) - Left[l].X) * deltay / deltax));
'            if (topcode[l])
'                UPPER[I] = max3(j_old, jI, UPPER[I]);
'            else LOWER[I] = min3(j_old, jI, LOWER[I]);
'            j_old = jI;
'         }
'      }
'      // Per la colonna I del video, il triangolo Š associato solo
'      // con i rettangoli delle righe LOWER[I],...,UPPER[I].
'      for (I=Imin; I<=Imax; I++)
'      for (J=LOWER[I]; J<=UPPER[I]; J++)
'      {  p_old = SCREEN[I][J];
'         SCREEN[I][J] = p_new = AllocMem1(node);
'         if (p_new == NULL) memproblem('G');
'         p_new->jtr = i; p_new->next = p_old;
'      }
'   }
'}

End Sub


Function SetX(x As Double) As Integer
    x = Int(density * (x - x_min))
    If (x < 0) Then
       x = 0
       outside = 1
    End If
    If (x > X__max) Then
      x = X__max
      outside = 1
    End If
   SetX = x
End Function

Function SetY(Y As Double) As Integer

   Y = Y__max - Int(density * (Y - y_min))
   If (Y < 0) Then
      Y = 0
      outside = 1
   End If
   
   If (Y > Y__max) Then
      Y = Y__max
      outside = 1
   End If
   
   SetY = Y

End Function


Sub stack_linesegment(p As Point, q As Point, k0 As Integer)

  Dim Pt As Integer
  Dim xP As Integer
  Dim yP As Integer
  Dim xQ As Integer
  Dim yQ As Integer
  xP = p.Pntscr.x
  yP = p.Pntscr.Y
  xQ = q.Pntscr.x
  yQ = q.Pntscr.Y
  If (Abs(xP - xQ) + Abs(yP - yQ) < 50) Then Exit Sub ' Non conviene
  
 ' Pt = UBound(stptr) + 1
 ' ReDim Preserve stptr(Pt)
  
  stptr(Pt).p = p
  stptr(Pt).q = q
  stptr(Pt).k0 = k0
  
End Sub


Function xCoord(nr As Integer) As Integer
  xCoord = (CLng(nr) * LARGE) / Nscreen
End Function

Function xFloat(x As Integer) As Double
      xFloat = (x / xfactor) + xmin
End Function

Function zFloat(Z As Integer) As Double
      zFloat = Z / zfactor + zmin
End Function


Function xIntScr(x As Double, xxMin As Double) As Integer
   xIntScr = (x - xxMin) * xfactor + 1 '0.5
End Function

Function yFloat(Y As Integer) As Double
   yFloat = (Y / yfactor) + ymin
End Function

Function yIntScr(Y As Double, yyMin As Double) As Integer
   yIntScr = (Y - yyMin) * yfactor + 1 '  0.5
End Function

