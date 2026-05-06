Attribute VB_Name = "ModuleFonctionsMath"
Function Min(a As Variant, b As Variant) As Variant

    If a > b Then
        Min = b
        Else: Min = a
    End If
    
End Function

Function Max(a As Variant, b As Variant) As Variant

    If a < b Then
        Max = b
        Else: Max = a
    End If
    
End Function

Function SommeCol(CelOrigine As Range, i As Integer, j As Integer) As Double
    Dim k As Integer
    
    SommeCol = 0
    For k = i To j
        SommeCol = SommeCol + CelOrigine.Offset(k, 0)
    Next k
    
End Function

Public Sub Polint3(X() As Double, y() As Double, poly() As Double)

'Trouve les coefficients du polynôme d'ordre 3 passant par 4 points
'X() et Y() : tableau de dimension 4 contenant les coordonnées des points
'poly() contient les coefficients a b c et d du polynôme : a + bx + cx2 + dx3
' La formule de Lagrange est utilisée

Dim denom1, denom2, denom3, denom4 As Double 'dénominateurs de la formule de Lagrange


denom1 = (X(1) - X(2)) * (X(1) - X(3)) * (X(1) - X(4))
denom2 = (X(2) - X(1)) * (X(2) - X(3)) * (X(2) - X(4))
denom3 = (X(3) - X(1)) * (X(3) - X(2)) * (X(3) - X(4))
denom4 = (X(4) - X(1)) * (X(4) - X(2)) * (X(4) - X(3))

poly(4) = y(1) / denom1 + y(2) / denom2 + y(3) / denom3 + y(4) / denom4

poly(3) = y(1) / denom1 * (-X(4) - X(2) - X(3))
poly(3) = poly(3) + y(2) / denom2 * (-X(4) - X(1) - X(3))
poly(3) = poly(3) + y(3) / denom3 * (-X(4) - X(1) - X(2))
poly(3) = poly(3) + y(4) / denom4 * (-X(3) - X(1) - X(2))

poly(2) = y(1) / denom1 * (X(4) * (X(2) + X(3)) + X(2) * X(3))
poly(2) = poly(2) + y(2) / denom2 * (X(4) * (X(1) + X(3)) + X(1) * X(3))
poly(2) = poly(2) + y(3) / denom3 * (X(4) * (X(1) + X(2)) + X(1) * X(2))
poly(2) = poly(2) + y(4) / denom4 * (X(3) * (X(1) + X(2)) + X(1) * X(2))

poly(1) = y(1) / denom1 * (-X(2) * X(3) * X(4))
poly(1) = poly(1) + y(2) / denom2 * (-X(1) * X(3) * X(4))
poly(1) = poly(1) + y(3) / denom3 * (-X(1) * X(2) * X(4))
poly(1) = poly(1) + y(4) / denom4 * (-X(1) * X(2) * X(3))

End Sub

Public Function MaxCol(CelRef As Range, NbRows As Integer)
'Calcul la valeur maxi contenue dans une colonne

Dim i As Integer

MaxCol = 0

For i = 0 To NbRows
    If (CelRef.Offset(i).Value > MaxCol) Then
        MaxCol = CelRef.Offset(i).Value
    End If
Next i
End Function


Function Interpolate(x1 As Double, x2 As Double, Y1 As Double, Y2 As Double, X As Double)
    'Interpolation linéaire entre 2 points pour trouver la valeur Y à l'abscisse T
    
    Interpolate = (Y2 - Y1) / (x2 - x1) * X - (Y2 - Y1) / (x2 - x1) * x1 + Y1

End Function

Public Function Interpolation_lineaire(X As Double, abscisses() As Double, ordonnees() As Double, n As Integer) As Double

Dim i As Integer
    
If X <= abscisses(1) Then
    Interpolation_lineaire = ordonnees(1)
ElseIf X >= abscisses(n) Then
    Interpolation_lineaire = ordonnees(n)
Else
    i = 1
    While X > abscisses(i + 1) And i <= n
        i = i + 1
    Wend
        
    Interpolation_lineaire = (X - abscisses(i)) * (ordonnees(i + 1) - ordonnees(i)) / (abscisses(i + 1) - abscisses(i)) + ordonnees(i)
End If

End Function

Public Sub Encadrer(y As Double, donnees() As Double, n As Integer, a As Integer, b As Integer)
    '******************************************************************/
    '*  Objet :                                                       */
    '*  donne les indices a et b tels que donnees[a] < y < donnees[b] */
    '******************************************************************/
    '*  Auteur : L. FERRAND                                           */
    '*  Date : 8 février 2006                                        */
    '******************************************************************/
    '*  Description des paramètres :                                  */
    '*      y : variable à encadrer                                   */
    '*      donnee() : tableau des données                           */
    '*      n : taille du tableau                                     */
    '*      a, b : indices tels que donnees(a) < y < donnees(b)       */
    '******************************************************************/
    '*  Remarques :                                                   */
    '*  les données contenues dans le tableau sont triées dans        */
    '*  l'ordre croissant                                             */
    '******************************************************************/

Dim i As Integer
    
i = 1
    
If y > donnees(n) Then
    a = n - 1
    b = n
    MsgBox "Attention : emissivité hors du domaine des tables de Hottel. La valeur par saturation sera retenue", vbInformation
Else
    While y > donnees(i) And i <= n
        i = i + 1
    Wend
            
    If i = 0 Then
        a = 1
        b = 2
    Else
        If (i = n) Then
            a = n - 1
            b = n
        Else
            a = i - 1
            b = i
        End If
    End If
End If

End Sub

Public Function Integrale(n As Integer, poly() As Double, x1 As Double, x2 As Double) As Double
'---------------------------------------------------------------------
'  Objet : Calcul l'integrale d'un polynôme poly() d'ordre n-1
'               y = poly(1)+poly(2)*T+poly(3)*T^2+poly(n)*T^(n-1)
'          entre les bornes T1 et T2
'---------------------------------------------------------------------
'  L.Ferrand, 14/02/2006
'---------------------------------------------------------------------
Dim i As Integer

Integrale = 0

For i = 1 To nval
    Integrale = Integrale + poly(i) / i * x2 ^ i - poly(i) / i * x1 ^ i
Next i

End Function

