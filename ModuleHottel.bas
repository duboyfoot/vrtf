Attribute VB_Name = "ModuleHottel"
Public Sub ChargerTablesHottel(temperature() As Double, pCO2H2O_L() As Double, emiss_CO2() As Double, emiss_H2O() As Double, T_correcH2O() As Double, pH2O_L() As Double, correc_H2O() As Double, pCO2H2O_L_correc() As Double, pH2O_pH2OCO2_correc() As Double, correc126() As Double, correc535() As Double, correc940() As Double)
'Charge les tables à partir de la feuille Excel formattée

ReDim temperature(25)
ReDim pCO2H2O_L(18)
ReDim emiss_CO2(18, 25)
ReDim emiss_H2O(18, 25)
ReDim T_correcH2O(5)
ReDim pH2O_L(8)
ReDim correc_H2O(8, 5)
ReDim pCO2H2O_L_correc(10)
ReDim pH2O_pH2OCO2_correc(11)
ReDim correc126(10, 11)
ReDim correc535(10, 11)
ReDim correc940(10, 11)

Dim i As Integer
Dim j As Integer
Dim Cel As Range

' Remplissage de temperature()
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "Températures[25]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 25
    temperature(i) = Cel.Offset(0, i - 1)
Next i

' Remplissage de pCO2H2O_L()
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "pCO2H2O_L[18]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 18
    pCO2H2O_L(i) = Cel.Offset(0, i - 1)
Next i

' Remplissage de emiss_CO2(18,25)
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "emiss_CO2[18][25]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 18
    For j = 1 To 25
        emiss_CO2(i, j) = Cel.Offset(j - 1, i - 1)
    Next j
Next i

' Remplissage de emiss_H2O(18,25)
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "emiss_H20[18][25]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 18
    For j = 1 To 25
        emiss_H2O(i, j) = Cel.Offset(j - 1, i - 1)
    Next j
Next i

' Remplissage de T_correcH2O()
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "T_correcH2O[5]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 5
    T_correcH2O(i) = Cel.Offset(0, i - 1)
Next i


' Remplissage de pH2O_L()
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "pH2O_L[8]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 8
    pH2O_L(i) = Cel.Offset(0, i - 1)
Next i

' Remplissage de correc_H2O(8,5)
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "correc_H2O[8][5]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 8
    For j = 1 To 5
        correc_H2O(i, j) = Cel.Offset(j - 1, i - 1)
    Next j
Next i


' Remplissage de pCO2H2O_L_correc()
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "pCO2H2O_L_correc[10]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 10
    pCO2H2O_L_correc(i) = Cel.Offset(0, i - 1)
Next i

' Remplissage de pH20_pH2OCO2_correc[11]
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "pH20_pH2OCO2_correc[11]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 11
    pH2O_pH2OCO2_correc(i) = Cel.Offset(0, i - 1)
Next i

' Remplissage de correc126[10][11]
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "correc126[10][11]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 10
    For j = 1 To 11
        correc126(i, j) = Cel.Offset(j - 1, i - 1)
    Next j
Next i

' Remplissage de correc535[10][11]
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "correc535[10][11]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 10
    For j = 1 To 11
        correc535(i, j) = Cel.Offset(j - 1, i - 1)
    Next j
Next i

' Remplissage de correc940[10][11]
Set Cel = Worksheets(SHottel).Range("A1")
i = 0
While Cel.Offset(i).Value <> "correc940[10][11]"
    i = i + 1
Wend
Set Cel = Cel.Offset(i + 1) 'Origine du tableau
For i = 1 To 10
    For j = 1 To 11
        correc940(i, j) = Cel.Offset(j - 1, i - 1)
    Next j
Next i

End Sub

Public Function EmissiviteHottel(pCO2 As Double, pH2O As Double, Tzone As Double, epaisseurgaz As Double)

'---------------------------------------------------------------'
'   Calcul de l'émissivité des fumées d'une zone                '
'   par les tables de Hottel                                    '
'   L. Ferrand, 8/02/2006                                       '
'---------------------------------------------------------------'
' pCO2 : pression partielle de CO2                              '
' pH2O : pression partielle de H2O                              '
' Tzone : Température de la zone [°C]                           '
' epaisseurgaz : longeur moyenne de rayon de la zone [m]        '
'---------------------------------------------------------------'

Dim i As Integer
Dim j As Integer

Dim temperature() As Double
Dim pCO2H2O_L() As Double
Dim emiss_CO2() As Double
Dim emiss_H2O() As Double
Dim T_correcH2O() As Double
Dim pH2O_L() As Double
Dim correc_H2O() As Double
Dim pCO2H2O_L_correc() As Double
Dim pH2O_pH2OCO2_correc() As Double
Dim correc126() As Double
Dim correc535() As Double
Dim correc940() As Double


Dim correc(10, 11) As Double
Dim pL(10) As Double
Dim errH2O(10) As Double
Dim errCO2(10) As Double
Dim R_CO2 As Double ' Rayon corrigé de la masse gazeuse
Dim R_H2O As Double ' Rayon corrigé de la masse gazeuse
Dim J1 As Integer
Dim J2 As Integer
Dim curve(25) As Double
Dim epsilonCO2 As Double
Dim epsilonH2O As Double
Dim epsilonH2Op As Double
Dim T(3) As Double
Dim Delta_eps As Double

pL(1) = 0.01
pL(2) = 0.013
pL(3) = 0.02
pL(4) = 0.05
pL(5) = 0.1
pL(6) = 0.2
pL(7) = 0.3
pL(8) = 0.5
pL(9) = 0.7
pL(10) = 1

errH2O(1) = 0.97
errH2O(2) = 0.964
errH2O(3) = 0.958
errH2O(4) = 0.942
errH2O(5) = 0.93
errH2O(6) = 0.906
errH2O(7) = 0.892
errH2O(8) = 0.874
errH2O(9) = 0.862
errH2O(10) = 0.85

errCO2(1) = 0.85
errCO2(2) = 0.844
errCO2(3) = 0.835
errCO2(4) = 0.815
errCO2(5) = 0.8
errCO2(6) = 0.791
errCO2(7) = 0.786
errCO2(8) = 0.779
errCO2(9) = 0.775
errCO2(10) = 0.77

T(1) = 126
T(2) = 535
T(3) = 940

Call ChargerTablesHottel(temperature, pCO2H2O_L, emiss_CO2, emiss_H2O, T_correcH2O, pH2O_L, correc_H2O, pCO2H2O_L_correc, pH2O_pH2OCO2_correc, correc126, correc535, correc940)
    
' Correction des rayons équivalents de la masse gazeuze
' (cf. W.H. McADAMS "Transmission de la Chaleur", éd. Dunod 1961, tableau 4.3 page 100

R_CO2 = epaisseurgaz * Interpolation_lineaire(pCO2 * epaisseurgaz, pL, errCO2, 10)
R_H2O = epaisseurgaz * Interpolation_lineaire(pH2O * epaisseurgaz, pL, errH2O, 10)
    
'--------------- Calcul de l'émissivité du composant CO2------------------'
    
Call Encadrer(pCO2 * R_CO2, pCO2H2O_L, 18, J1, J2)
    
For i = 1 To 25
    If pCO2H2O_L(J2) - pCO2H2O_L(J1) = 0 Then
        curve(i) = 0
    Else
        curve(i) = emiss_CO2(J1, i) + (pCO2 * R_CO2 - pCO2H2O_L(J1)) * (emiss_CO2(J2, i) - emiss_CO2(J1, i)) / (pCO2H2O_L(J2) - pCO2H2O_L(J1))
    End If
Next i
        
epsilonCO2 = Interpolation_lineaire(Tzone, temperature, curve, 25)
    
    
'----- Calcul de l'émissivité du composant H2O, sans correction par sa pression partielle ----'
   
Call Encadrer(pH2O * R_H2O, pCO2H2O_L, 18, J1, J2)
    
For i = 1 To 25
    If pCO2H2O_L(J2) - pCO2H2O_L(J1) = 0 Then
        curve(i) = 0
    Else
        curve(i) = emiss_H2O(J1, i) + (pH2O * R_H2O - pCO2H2O_L(J1)) * (emiss_H2O(J2, i) - emiss_H2O(J1, i)) / (pCO2H2O_L(J2) - pCO2H2O_L(J1))
    End If
Next i

epsilonH2Op = Interpolation_lineaire(Tzone, temperature, curve, 25)
    
'------------- Correction de epsilon H20 par la pression partielle--------------------'
    
Call Encadrer(pH2O * R_H2O, pH2O_L, 8, J1, J2)
    
For i = 1 To 5
    If (pH2O_L(J2) - pH2O_L(J1)) = 0 Then
        curve(i) = 0
    Else
        curve(i) = correc_H2O(J1, i) + (pH2O * R_H2O - pH2O_L(J1)) * (correc_H2O(J2, i) - correc_H2O(J1, i)) / (pH2O_L(J2) - pH2O_L(J1))
    End If
Next i

epsilonH2O = epsilonH2Op * Interpolation_lineaire(pH2O, T_correcH2O, curve, 5)
    
'-------- Calcul de facteur de correction du au recouvrement de bandes Delta_eps ---'
    
If (Tzone <= 126) Then
    For j = 1 To 11
        For i = 1 To 10
            correc(i, j) = correc126(i, j)
        Next i
    Next j
ElseIf Tzone >= 940 Then
    For j = 1 To 11
        For i = 1 To 10
             correc(i, j) = correc940(i, j)
        Next i
    Next j
Else
    For j = 1 To 11
        For i = 1 To 10
            curve(1) = correc126(i, j)
            curve(2) = correc535(i, j)
            curve(3) = correc940(i, j)
            correc(i, j) = Interpolation_lineaire(Tzone, T, curve, 3)
        Next i
    Next j
End If
    
Call Encadrer(pH2O * R_H2O + pCO2 * R_CO2, pCO2H2O_L_correc, 10, J1, J2)

For i = 1 To 11
    If (pCO2H2O_L_correc(J2) - pCO2H2O_L_correc(J1)) = 0 Then
        curve(i) = 0
    Else
        curve(i) = correc(J1, i) + (pH2O * R_H2O + pCO2 * R_CO2 - pCO2H2O_L_correc(J1)) * (correc(J2, i) - correc(J1, i)) / (pCO2H2O_L_correc(J2) - pCO2H2O_L_correc(J1))
    End If
Next i
    
Delta_eps = Interpolation_lineaire(pH2O / (pH2O + pCO2), pH2O_pH2OCO2_correc, curve, 11)
    
EmissiviteHottel = epsilonCO2 + epsilonH2O - Delta_eps
    
End Function

Public Function CoeffAbsorptionHottel(pCO2 As Double, pH2O As Double, Tzone As Double, l As Double)
'-----------------------------------------------------------------'
'   Calcul du coeffcient d'absorption des fumées                  '
'   d'une zone par les tables de Hottel et la loi de Béer-Lambert '
'   L. Ferrand, 8/02/2006                                         '
'-----------------------------------------------------------------'
' pCO2 : pression partielle de CO2                                '
' pH2O : pression partielle de H2O                                '
' Tzone : Température de la zone [°C]                             '
' L : longeur moyenne de rayon de la zone [m]                     '
'-----------------------------------------------------------------'

Dim Epsilon As Double
Dim p As Double ' Hypothèse : pression atmosphérique

p = 1.013
    
Epsilon = EmissiviteHottel(pCO2, pH2O, Tzone, l)

If Epsilon > 1 Then Epsilon = 0.99

CoeffAbsorptionHottel = -1 / (p * l) * Math.Log(1 - Epsilon)
    
End Function
