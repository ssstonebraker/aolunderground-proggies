Attribute VB_Name = "TRIANGL"
' Modulo per la triangolazione di poligoni.
' Qui viene scomposta una matrice di poligoni "convessi"
' in un'altra fatta di soli triangoli.

Type Trianrs
  a As Integer
  b As Integer
  C As Integer
End Type

Function Triangul(pol() As Integer, n As Integer, nrs() As Trianrs, Ori As Integer) As Integer
'/* TRIANGUL: Triangolazione di poligoni
'      Triangolazione di un poligono con numeri di vertice consecutivi
'      pol[0],..., pol[n-1], in senso antiorario.
'      Dati tre numeri di vertice P, Q, R, la funzione orienta deve
'      determinare il loro orientamento:
'      Negativo = in senso orario
'      Zero     = sulla stessa linea
'      Positivo = in senso antiorario
'      Se la triangolazione è possibile, i triangoli risultanti
'      sono memorizzati in successione nell'array 'nrs'. Il triangolo j ha
'      numeri di vertice nrs[j].A, nrs[j].B, nrs[j].C.
'      Lo spazio in memoria per l'array 'nrs' deve essere fornito dalla
'      funzione chiamante.
'      Valore restituito:
'            il numero dei triangoli trovati, oppure
'            -1 se il poligono non è adatto o i vertici sono in senso orario.
'            -2: memoria insufficiente.
'*/

'int triangul(int *pol, int n, trianrs *nrs,
'             int orienta(int P, int Q, int R))
   Dim Ptr() As Integer
   Dim ort() As Integer
   Dim q As Integer, qA As Integer, qB As Integer, qC As Integer, r As Integer '  -1 usato come 'NULL'
   Dim i As Integer, i1 As Integer, i2 As Integer, j As Integer, k As Integer, m As Integer, ok As Integer, ortB As Integer, Polconvex  As Integer
   Dim a As Integer, b As Integer, C As Integer, p As Integer, collinear As Integer
    
   r = -1
   Polconvex = True
    
    If n < 3 Then
       Triangul = -1 ' // Nessun poligono
       Exit Function
    End If

    If n = 3 Then
      nrs(0).a = pol(0)
      nrs(0).b = pol(1)
      nrs(0).C = pol(2)
      Triangul = 1   ' // Solo un triangolo
      Exit Function
    End If
    
    ReDim ort(n) ' // ort[i] = 1 se il vertice i è convesso
    
    Do
      collinear = False
       For i = 0 To n - 1
          If i < n - 1 Then i1 = i + 1 Else i1 = 0
          If i1 < n - 1 Then i2 = i1 + 1 Else i2 = 0
          ort(i1) = orienta(pol(i), pol(i1), pol(i2))
          If ort(i1) = 0 Then
            collinear = True
            For j = i1 To n - 1
                pol(j) = pol(j + 1)
            Next j
            n = n - 1
            Exit For
          End If
          If ort(i1) < 1 Then Polconvex = False
        Next i
    Loop While collinear
    
    If n < 3 Then
       Triangul = -1
       Exit Function
    End If


    If Polconvex Then          '  // Usa le diagonali passanti per il vertice 0:
       For j = 0 To n - 2
          nrs(j).a = pol(0)
          nrs(j).b = pol(j + 1)
          nrs(j).C = pol(j + 2)
       Next
       
       Erase ort
       Triangul = n - 2
       Exit Function
    End If

    ReDim Ptr(n)

   ' // Crea una lista concatenata circolare con i numeri di vertice:
    
    For i = 1 To n - 1: Ptr(i - 1) = i: Next i
    
    Ptr(n - 1) = 0
    q = 0
    qA = Ptr(q)
    qB = Ptr(qA)
    qC = Ptr(qB)
    j = 0            '  // j triangoli memorizzati fino a questo punto
    
    For m = n To 3 Step -1 ' // m nodi restanti nella lista circolare.
      For k = 0 To m
       '  // Prova con il triangolo ABC:
          ortB = ort(qB)
          ok = False
       '   // B è un candidato, se è convesso:
          If (ortB > 0) Then
             a = pol(qA)
             b = pol(qB)
             C = pol(qC)
             ok = True
             r = Ptr(qC)
             Do While r <> qA And ok
                p = pol(r)     ' // ABC in senso antiorario:
                ok = p = a Or p = b Or p = C Or orienta(a, b, p) < 0 Or orienta(b, C, p) < 0 Or orienta(C, a, p) < 0
                r = Ptr(r)
             Loop
          '   // ok significa: P coincidente con A, B o C
          '   //           oppure esterno ad ABC
             If ok Then
               nrs(j).a = pol(qA)
               nrs(j).b = pol(qB)
               nrs(j).C = pol(qC)
               j = j + 1
             End If
          End If
          
          If (ok Or ortB = 0) Then
         ' {  // Elimina il trianglolo ABC dal poligono:
             Ptr(qA) = qC
             qB = qC
             qC = Ptr(qC)
             If ort(qA) < 1 Then ort(qA) = orienta(pol(q), pol(qA), pol(qB))
             If ort(qB) < 1 Then ort(qB) = orienta(pol(qA), pol(qB), pol(qC))
             Do While ort(qA) = 0 And m > 2
                Ptr(q) = qB
                qA = qB
                qB = qC
                qC = Ptr(qC)
                m = m - 1
             Loop
             Do While ort(qB) = 0 And m > 2
               Ptr(qA) = qC
               qB = qC
               qC = Ptr(qC)
               m = m - 1
             Loop

             Exit For
          
          End If
          
          q = qA
          qA = qB
          qB = qC
          qC = Ptr(qC)
       Next
   Next
   Triangul = j ' // j N° triangoli

End Function

