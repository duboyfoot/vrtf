Attribute VB_Name = "Enthalpyfluide"

Function EnthalpAir(Tair As Double) As Double
    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim Name As String
    Name = "air"
    Call Gaz_GetCompo(SComburant, Name, nbAir, AirName, AirCompo)
    EnthalpAir = CalculEnthalp(nbAir, AirName, AirCompo, Tair, 273, "[J/Kg]")
    EnthalpAir = EnthalpAir * AirDensity(273)
End Function
Function EnthalpAirKG(Tair As Double) As Double
    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim Name As String
    Name = "air"
    Call Gaz_GetCompo(SComburant, Name, nbAir, AirName, AirCompo)
    EnthalpAirKG = CalculEnthalp(nbAir, AirName, AirCompo, Tair, 273, "[J/Kg]")
End Function
Function EnthalpO2(Tair As Double) As Double
    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim Name As String
    Name = "o2"
    Call Gaz_GetCompo(SComburant, Name, nbAir, AirName, AirCompo)
    EnthalpO2 = CalculEnthalp(nbAir, AirName, AirCompo, Tair, 273, "[J/Kg]")
    EnthalpO2 = EnthalpO2 * AirDensity(Tair)
End Function
Function EnthalpN2H2(PourcentageH2 As Double, TGaz As Double) As Double

    Dim Ngaz As Integer
    Dim Tref As Double
    Dim GazName(2) As String
    Dim GazCompo(2) As Double
    Dim Name As String
    Ngaz = 2
    Tref = 293
    GazName(1) = "H2"
    GazName(2) = "N2"
    GazCompo(1) = PourcentageH2 / 100
    GazCompo(2) = 1 - PourcentageH2 / 100
    
    EnthalpN2H2 = CalculEnthalp(Ngaz, GazName(), GazCompo(), TGaz, Tref, "[J/Nm3]")
End Function
Function EnthalpN2H2Kg(PourcentageH2 As Double, TGaz As Double) As Double

    Dim Ngaz As Integer
    Dim Tref As Double
    Dim GazName(2) As String
    Dim GazCompo(2) As Double
    Dim Name As String
    Ngaz = 2
    Tref = 293
    GazName(1) = "H2"
    GazName(2) = "N2"
    GazCompo(1) = PourcentageH2 / 100
    GazCompo(2) = 1 - PourcentageH2 / 100
    EnthalpN2H2Kg = CalculEnthalp(Ngaz, GazName(), GazCompo(), TGaz, Tref, "[J/Kg]")
End Function
Function EnthalpGas(Name As String, Tgas As Double) As Double
    
    Dim Nbgas As Integer
    Dim Gasname() As String
    Dim gasCompo() As Double
   
    Call Gaz_GetCompo(SFuel, Name, Nbgas, Gasname, gasCompo)

EnthalpGas = CalculEnthalp(Nbgas, Gasname, gasCompo, Tgas, 273, "[J/kg]")
EnthalpGas = EnthalpGas * GasDensity(Name, 273)
End Function
Function TWasteGasEnthalp(Name As String, Egas As Double) As Double


' Calcul de la température de sortie du gaz initialement froid

Dim Tfum As Double
olddelta_p = 350
delta_p = 1
TempWG = 350
Delta_t = 350


i = 0
     
While (Abs(P2 - Egas) > 2 And i < 100)
    i = i + 1
    If (olddelta_p * delta_p < 0) Then
        Delta_t = Delta_t / 2
    End If
    olddelta_p = delta_p
    If (P2 <= Egas) Then
        Tfum = Tfum + Delta_t
    Else
        Tfum = Tfum - Delta_t
        If (Tfum < 273) Then Tfum = 273
    End If
    P2 = EnthalpWasteGas(Name, Tfum, Egas)
    delta_p = P2 - Egas

Wend
    If i < 1000 Then TWasteGasEnthalp = Tfum
End Function
Function EnthalpGasKG(Name As String, Tgas As Double) As Double
    
    Dim Nbgas As Integer
    Dim Gasname() As String
    Dim gasCompo() As Double
   
    Call Gaz_GetCompo(SFuel, Name, Nbgas, Gasname, gasCompo)

EnthalpGasKG = CalculEnthalp(Nbgas, Gasname, gasCompo, Tgas, 273, "[J/kg]")
End Function
Function EnthalpWasteGas(Gasname As String, TFume As Double, ratio As Double) As Double
    Dim Nbgas As Integer
    Dim FuelCompo() As Double
    Dim FuelName() As String
    Dim nbFuel As Integer
    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim NameAir As String
    NameAir = "air"
    Dim FumName(7) As String
    Dim FumCompo(7) As Double
    FumName(1) = "CO2"
    FumName(2) = "H2O"
    FumName(3) = "O2"
    FumName(4) = "N2"
    FumName(5) = "SO2"
    FumName(6) = "CO"
    FumName(7) = "H2"
    Call Gaz_GetCompo(SFuel, Gasname, nbFuel, FuelName, FuelCompo)
    If (nbFuel = 0) Then
        EnthalpWasteGas = -1
        Exit Function
    End If
    Call Gaz_GetCompo(SComburant, NameAir, nbAir, AirName, AirCompo)
    Call WasteGasCompo(FumName, FumCompo, 20, nbFuel, FuelName, FuelCompo, nbAir, AirName, AirCompo, ratio)
    WasteGasDensity1 = CalculMasseVolumique(7, FumName, FumCompo, 273, 101325, "[kg/m3]")
    EnthalpWasteGas = CalculEnthalp(7, FumName, FumCompo, TFume, 273, "[J/kg]") * WasteGasDensity1
    
End Function
Function EnthalpWasteGasKG(Gasname As String, TFume As Double, ratio As Double) As Double
    Dim Nbgas As Integer
    Dim FuelCompo() As Double
    Dim FuelName() As String
    Dim nbFuel As Integer
    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim NameAir As String
    NameAir = "air"
    Dim FumName(7) As String
    Dim FumCompo(7) As Double
    FumName(1) = "CO2"
    FumName(2) = "H2O"
    FumName(3) = "O2"
    FumName(4) = "N2"
    FumName(5) = "SO2"
    FumName(6) = "CO"
    FumName(7) = "H2"
    Call Gaz_GetCompo(SFuel, Gasname, nbFuel, FuelName, FuelCompo)
    If (nbFuel = 0) Then
        EnthalpWasteGasKG = -1
        Exit Function
    End If
    Call Gaz_GetCompo(SComburant, NameAir, nbAir, AirName, AirCompo)
    Call WasteGasCompo(FumName, FumCompo, 20, nbFuel, FuelName, FuelCompo, nbAir, AirName, AirCompo, ratio)

    EnthalpWasteGasKG = CalculEnthalp(7, FumName, FumCompo, TFume, 273, "[J/kg]")
    
End Function
Function Gethgaz(NomGaz As String, Tref As Double, TKelvin As Double, PropSheet As String)

' Recherche du gaz dans la liste et calcul de son enthalpie à la température Tcelsius, Tref étant la température de référence
' Enthalpie en [J/kg]

    For j = 1 To 7
        Gethgaz = Gethgaz + ThisWorkbook.Sheets(PropSheet).Range("B3").Offset(GetIndexGaz(NomGaz), 10 + j) / j * ((TKelvin) ^ j - (Tref) ^ j)
    Next j


End Function


Public Function AirDensity(Tair As Double) As Double

    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim Gasname As String
    Gasname = "air"
    Call Gaz_GetCompo(SComburant, Gasname, nbAir, AirName, AirCompo)
    
    AirDensity = CalculMasseVolumique(nbAir, AirName, AirCompo, Tair, 101325, "[kg/m3]")
    
    

End Function
Function N2H2Density(PourcentageH2 As Double, TGaz As Double) As Double

    Dim Ngaz As Integer
    Dim Tref As Double
    Dim GazName(2) As String
    Dim GazCompo(2) As Double
    Dim Name As String
    Ngaz = 2
    Tref = 293
    GazName(1) = "H2"
    GazName(2) = "N2"
    GazCompo(1) = PourcentageH2
    GazCompo(2) = 100 - PourcentageH2
    
    N2H2Density = CalculMasseVolumique(Ngaz, GazName(), GazCompo(), TGaz, 101325, "[kg/m3]")
End Function
Public Function GasDensity(Name As String, temp As Double) As Double

    Dim Nbgas As Integer
    Dim Gasname() As String
    Dim gasCompo() As Double


    Call Gaz_GetCompo(SFuel, Name, Nbgas, Gasname, gasCompo)
    
    GasDensity = CalculMasseVolumique(Nbgas, Gasname, gasCompo, temp, 101325, "[kg/m3]")
    
    

End Function
Public Function WasteGasDensity(Name As String, temp As Double) As Double


   Dim Nbgas As Integer
    Dim FuelCompo() As Double
    Dim FuelName() As String
    Dim nbFuel As Integer
    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim NameAir As String
    NameAir = "air"
    Dim FumName(7) As String
    Dim FumCompo(7) As Double
    FumName(1) = "CO2"
    FumName(2) = "H2O"
    FumName(3) = "O2"
    FumName(4) = "N2"
    FumName(5) = "SO2"
    FumName(6) = "CO"
    FumName(7) = "H2"
    Call Gaz_GetCompo(SFuel, Name, nbFuel, FuelName, FuelCompo)
    If (nbFuel = 0) Then
        WasteGasDensity = -1
        Exit Function
    End If
    Call Gaz_GetCompo(SComburant, NameAir, nbAir, AirName, AirCompo)
    Call WasteGasCompo(FumName, FumCompo, 20, nbFuel, FuelName, FuelCompo, nbAir, AirName, AirCompo, 1.1)
    WasteGasDensity = CalculMasseVolumique(7, FumName, FumCompo, temp, 101325, "[kg/m3]")
  
End Function

Public Sub Gaz_GetCompo(SBasdo As String, Name As String, ByRef NbPureGaz As Integer, ByRef PureGazNames() As String, PureGazCompo() As Double)


    '----------------------------------------------------------------
    ' Look-up for Gaz compostion in Database sheet SBasdo
    '----------------------------------------------------------------
    ' L. Ferrand, 01/2009
    '----------------------------------------------------------------
    Dim i As Integer
    Dim j As Integer

    With ThisWorkbook.Sheets(SBasdo)
    
        i = 1
        While LCase(Name) <> LCase(.Cells(2, i + 1)) And i <= .Cells(2, 1) + 1
            i = i + 1
        Wend
        If i > .Cells(2, 1) + 1 Then GoTo Suite
'            Call MsgBox("The Gaz " & Name & " was not found in the database.", vbExclamation, Fox)
'
'            Exit Sub
'        End If
        
        j = 1
        NbPureGaz = 0
        While j <= .Cells(1, 1)
                      
            If .Cells(j + 2, i + 1) <> 0 And .Cells(j + 2, i + 1) <> " " Then
                NbPureGaz = NbPureGaz + 1
                ReDim Preserve PureGazNames(NbPureGaz)
                ReDim Preserve PureGazCompo(NbPureGaz)
                PureGazNames(NbPureGaz) = .Cells(j + 2, 1)
                PureGazCompo(NbPureGaz) = .Cells(j + 2, i + 1)
            End If
            j = j + 1
        Wend
    
    End With
    Exit Sub
' le gaz est peut-etre en locale
Suite:
    With Sheets("Fuels")
           i = 1
        While LCase(Name) <> LCase(.Cells(2, i + 1)) And i <= .Cells(2, 1) + 1
            i = i + 1
        Wend
        If i > .Cells(2, 1) + 1 Then
             Exit Sub
        End If
        
        j = 1
        NbPureGaz = 0
        While j <= .Cells(1, 1)
                      
            If .Cells(j + 2, i + 1) <> 0 And .Cells(j + 2, i + 1) <> " " Then
                NbPureGaz = NbPureGaz + 1
                ReDim Preserve PureGazNames(NbPureGaz)
                ReDim Preserve PureGazCompo(NbPureGaz)
                PureGazNames(NbPureGaz) = .Cells(j + 2, 1)
                PureGazCompo(NbPureGaz) = .Cells(j + 2, i + 1)
            End If
            j = j + 1
        Wend
    
    End With
    
End Sub
Function CalculMasseVolumique(Ngaz As Integer, CombNames() As String, CombCompo() As Double, TKelvin As Double, Pabsolue As Double, Unite As String)
'Calcul de la masse volumique d'un composé constitué de Ngaz composants à la température Tcelsius
' et à la pression Pabsolue en Pa

'Ngaz : nombre de gaz contenus dans le composé
'CombNames() : nom de gaz élémentaires contenus dans le composé - l'index zéro n'est pas utilisé
'CombCompo() : pourcentage volumique [%] de chacun des gaz - l'index zéro n'est pas utilisé
    Dim i As Integer

    CalculMasseVolumique = 0

    For i = 1 To Ngaz
        CalculMasseVolumique = CalculMasseVolumique + GetMasseMolgaz(CombNames(i)) * CombCompo(i) / 100
    Next i

    CalculMasseVolumique = CalculMasseVolumique / 22.4136 * 273 / (TKelvin) * Pabsolue / 101325

    If Unite = "[lb/cf]" Then
        CalculMasseVolumique = CalculMasseVolumique * 2.20462 / 35.31467
    End If
End Function

Public Sub WasteGasCompo(ByRef FumNames() As String, ByRef FumCompo() As Double, _
                Tequilibrium As Double, _
                NgazFuel As Integer, FuelNames() As String, FuelCompo() As Double, _
                NgazAir As Integer, AirNames() As String, AirCompo() As Double, _
                AirFuelRatio As Double)

    '------------------------------------------------
    ' Calculation of Waste gases composition in
    ' sub-stoechiometric conditions
    ' at the Temperature Tequilibrium
    '------------------------------------------------
    ' Q.K.MAI, 12/2008
    '------------------------------------------------
    ' Modifications :
    ' L. FERRAND, 03/2009 - lifting for HMI
    '------------------------------------------------
    ' Temperature [°C]
    ' Compositions [%vol.]
    ' Tableau FumNames retourné :
    '    FumNames(1) = "CO2"
    '    FumNames(2) = "H2O"
    '    FumNames(3) = "O2"
    '    FumNames(4) = "N2"
    '    FumNames(5) = "SO2"
    '    FumNames(6) = "CO"
    '    FumNames(7) = "H2"
    '------------------------------------------------
    
    Dim i As Integer
    Dim j As Integer
    
    Dim n As Double '-- Equivalent CnHm of the hydrocarbon
    Dim m As Double
    
    Dim fuel_O2 As Double
    Dim fuel_CO2 As Double
    Dim fuel_H2O As Double
    Dim fuel_N2 As Double
    Dim fuel_SO2 As Double
    Dim fuel_H2 As Double
    Dim fuel_CO As Double
    
    Dim fuel_H2S As Double
    Dim fuel_CnHm As Double
    
    Dim air_O2 As Double
    Dim air_CO2 As Double
    Dim air_H2O As Double
    Dim air_N2 As Double
    Dim air_SO2 As Double
    
    Dim PC_AirO2 As Double ' Oxygen content in the comburant
    
    Dim combus_CO2 As Double
    Dim combus_H2O As Double
    Dim combus_CO As Double
    Dim combus_H2 As Double
    Dim combus_SO2 As Double
    
    Dim Fum_CO2 As Double
    Dim Fum_H2O As Double
    Dim Fum_O2 As Double
    Dim Fum_N2 As Double
    Dim Fum_SO2 As Double
    Dim Fum_CO As Double
    Dim Fum_H2 As Double
    
    Dim molTotFum As Double
        
    Dim molSt_O2 As Double
    Dim molreact_O2 As Double
    
    Dim GasFound As Boolean
    
    Dim KTfum As Double
    Dim KTadiab As Double
    
    Dim BB As Double
    Dim DD As Double
    
    '--- Initialisations --------
    n = 0
    m = 0
    fuel_O2 = 0
    fuel_CO2 = 0
    fuel_H2O = 0
    fuel_N2 = 0
    fuel_SO2 = 0
    fuel_H2 = 0
    fuel_CO = 0
    fuel_H2S = 0
    fuel_CnHm = 0
    
    
    '---- Calcul nombre de moles des réactifs dans le combustible ------
    For i = 1 To NgazFuel
    
        Select Case FuelNames(i)
            Case "O2"
                fuel_O2 = fuel_O2 + FuelCompo(i) / 100
            Case "CO2"
                fuel_CO2 = fuel_CO2 + FuelCompo(i) / 100
            Case "H2O"
                fuel_H2O = fuel_H2O + FuelCompo(i) / 100
            Case "N2"
                fuel_N2 = fuel_N2 + FuelCompo(i) / 100
            Case "SO2"
                fuel_SO2 = fuel_SO2 + FuelCompo(i) / 100
            Case "H2"
                fuel_H2 = fuel_H2 + FuelCompo(i) / 100
            Case "CO"
                fuel_CO = fuel_CO + FuelCompo(i) / 100
            Case "H2S"
                fuel_H2S = fuel_H2S + FuelCompo(i) / 100
            Case Else
                GasFound = False
                j = 0
                While (Not GasFound) And j < NbGazElementaires + 1
                    j = j + 1
                    If ThisWorkbook.Worksheets(Sprop).Range(CelDataOrig).Offset(j, 1) = FuelNames(i) Then
                        GasFound = True
                        If ThisWorkbook.Worksheets(Sprop).Range(CelDataOrig).Offset(j, 2) > 0 And ThisWorkbook.Worksheets(Sprop).Range(CelDataOrig).Offset(j, 3) > 0 Then ' Hydrocarbure
                            ' Hydrocarbure équivalent
                            fuel_CnHm = fuel_CnHm + FuelCompo(i) / 100
        
                            n = n + ThisWorkbook.Worksheets(Sprop).Range(CelDataOrig).Offset(j, 2) * FuelCompo(i) / 100
                            m = m + ThisWorkbook.Worksheets(Sprop).Range(CelDataOrig).Offset(j, 3) * FuelCompo(i) / 100
                        End If
                    End If
                    
                Wend
        End Select
        
    Next i
        
    ' Normalisation de n et m
    If fuel_CnHm <> 0 Then
        n = n / fuel_CnHm
        m = m / fuel_CnHm
    End If
    
    
    'Calcul nombre de moles des réactifs dans le comburant
    
    molSt_O2 = (n + m / 4) * fuel_CnHm + 1 / 2 * fuel_CO + 1 / 2 * fuel_H2 + 3 / 2 * fuel_H2S
    
    air_O2 = (molSt_O2 - fuel_O2) * AirFuelRatio
    PC_AirO2 = GetO2(NgazAir, AirNames(), AirCompo())
    air_CO2 = GetCO2(NgazAir, AirNames(), AirCompo()) / PC_AirO2 * air_O2
    air_H2O = GetH2O(NgazAir, AirNames(), AirCompo()) / PC_AirO2 * air_O2
    air_N2 = GetN2(NgazAir, AirNames(), AirCompo()) / PC_AirO2 * air_O2
    air_SO2 = GetSO2(NgazAir, AirNames(), AirCompo()) / PC_AirO2 * air_O2
    
    molreact_O2 = air_O2 + fuel_O2
        
    'Calcul nombre mole de produits de combustion
    
    
    KTfum = Exp(3.75 - 4075 / (Tequilibrium + 273))
    
    '--- KT=(combus_CO*combus_H2O)/(combus_H2*combus_CO2)
    '--- bilan de matiere
        'combus_CO2+combus_CO=n*gaz_CnHm  Carbone
        'combus_H2+combus_H2O=m/2*gaz_CnHm+gaz_H2  Hydrogene
        'combus_CO+2*combus_CO2+combus_H2O=gaz_CO+2*molreact_O2 Oxygene
    '==> 4 equations à 4 inconnues
    
    
    If AirFuelRatio < 1 Then
        
        '---- Source : Manuel de Thermique -----------------'
        ' Partie 2 Génération de la chaleur                 '
        ' § 1.5.3 Combustion avec défaut d'air              '
        ' Equilibre du "gaz à 'eau" CO2+H2 <-> CO + H2O     '
        ' Page 43-47                                        '
        '---------------------------------------------------'
        
        BB = -1 * (2 * molreact_O2 + fuel_CO - 2 * fuel_H2S + 2 * (fuel_CO2 + air_CO2) + _
                (fuel_H2O + air_H2O) + KTfum * (m / 2 * fuel_CnHm + n * fuel_CnHm + fuel_H2 - (fuel_CO2 + air_CO2) _
                + 3 * fuel_H2S - 2 * molreact_O2))
        
        DD = BB ^ 2 - 4 * (1 - KTfum) * (n * fuel_CnHm + fuel_CO + fuel_CO2 + air_CO2) * (2 * molreact_O2 - 2 * fuel_H2S - n * fuel_CnHm + fuel_CO2 + air_CO2 + fuel_H2O + air_H2O)
        
        combus_CO2 = (-BB - Sqr(DD)) / (2 * (1 - KTfum))
        combus_CO = n * fuel_CnHm + fuel_CO + fuel_CO2 + air_CO2 - combus_CO2
        combus_H2O = 2 * molreact_O2 - 2 * fuel_H2S + fuel_CO2 + air_CO2 + fuel_H2O + air_H2O - n * fuel_CnHm - combus_CO2
        combus_H2 = m / 2 * fuel_CnHm + fuel_H2 + fuel_H2S + fuel_H2O + air_H2O - combus_H2O
        
    Else ' Combustion stoechiométrique ou en excès d'air
    
        combus_CO2 = n * fuel_CnHm + fuel_CO + fuel_CO2 + air_CO2
        combus_CO = 0
        combus_H2O = m / 2 * fuel_CnHm + fuel_H2S + fuel_H2 + fuel_H2O + air_H2O
        combus_H2 = 0
        
    End If
    
    combus_SO2 = fuel_H2S
    
    If AirFuelRatio <= 1 Then
        Fum_O2 = 0 ' Pas d'oxygène résiduel en combustion sub-stoechiométrique
    Else
        Fum_O2 = fuel_O2 + air_O2 - molSt_O2
    End If
    
    'Calcul nombre de moles dans les fumées
    Fum_CO2 = combus_CO2
    Fum_H2O = combus_H2O
    Fum_N2 = fuel_N2 + air_N2
    Fum_SO2 = combus_SO2 + Fum_SO2 + air_SO2
    Fum_CO = combus_CO
    Fum_H2 = combus_H2
    
    molTotFum = Fum_O2 + Fum_CO2 + Fum_H2O + Fum_N2 + Fum_SO2 + Fum_CO + Fum_H2
    
    'Calcul des %vol. dans les fumees
    
    
    FumNames(1) = "CO2"
    FumNames(2) = "H2O"
    FumNames(3) = "O2"
    FumNames(4) = "N2"
    FumNames(5) = "SO2"
    FumNames(6) = "CO"
    FumNames(7) = "H2"
    
    FumCompo(1) = Fum_CO2 / molTotFum * 100
    FumCompo(2) = Fum_H2O / molTotFum * 100
    FumCompo(3) = Fum_O2 / molTotFum * 100
    FumCompo(4) = Fum_N2 / molTotFum * 100
    FumCompo(5) = Fum_SO2 / molTotFum * 100
    FumCompo(6) = Fum_CO / molTotFum * 100
    FumCompo(7) = Fum_H2 / molTotFum * 100


End Sub
Function CalculEnthalp(Ngaz As Integer, CombNames() As String, CombCompo() As Double, TKelvin As Double, Tref As Double, Unite As String)
'Calcul de l'enthalpie d'un composé constitué de Ngaz composants à la température Tcelsius, Tref étant la température de référence
' Enthalpie en [J/kg]
'Ngaz : nombre de gaz contenus dans le composé
'CombNames() : nom de gaz élémentaires contenus dans le composé - l'index zéro n'est pas utilisé
'CombCompo() : pourcentage volumique [%] de chacun des gaz - l'index zéro n'est pas utilisé
Dim i As Integer
CalculEnthalp = 0

For i = 1 To Ngaz
    CalculEnthalp = CalculEnthalp + Gethgaz(CombNames(i), Tref, TKelvin, Sprop) * GetFracMass(Ngaz, CombNames(), CombCompo(), i)
Next i

If Unite = "[kJ/kg]" Then CalculEnthalp = CalculEnthalp / 1000
If Unite = "[kcal/kg]" Then CalculEnthalp = CalculEnthalp / 1000 * 0.2388459
If Unite = "[BTU/lb]" Then CalculEnthalp = CalculEnthalp * 0.0009478171 / 2.20462


End Function

Function GetO2(Ngaz As Integer, CombNames() As String, CombCompo() As Double)

GetO2 = 0

' Recherche du % volumique d'O2 dans le comburant
Dim j As Integer

j = 0
For i = 1 To Ngaz
    If CombNames(i) = "O2" Then
        j = j + 1
        GetO2 = GetO2 + CombCompo(i)
    End If
Next i


End Function
Function GetCO2(Ngaz As Integer, CombNames() As String, CombCompo() As Double)

GetCO2 = 0

' Recherche du % volumique de CO2 dans le comburant
Dim j As Integer

j = 0
For i = 1 To Ngaz
    If CombNames(i) = "CO2" Then
        j = j + 1
        GetCO2 = GetCO2 + CombCompo(i)
    End If
Next i


End Function
Function GetH2O(Ngaz As Integer, CombNames() As String, CombCompo() As Double)

GetH2O = 0

' Recherche du % volumique d'H2O dans le comburant
Dim j As Integer

j = 0
For i = 1 To Ngaz
    If CombNames(i) = "H2O" Then
        j = j + 1
        GetH2O = GetH2O + CombCompo(i)
    End If
Next i

End Function

Function GetN2(Ngaz As Integer, CombNames() As String, CombCompo() As Double)

GetN2 = 0

' Recherche du % volumique de N2 dans le comburant
Dim j As Integer

j = 0
For i = 1 To Ngaz
    If CombNames(i) = "N2" Then
        j = j + 1
        GetN2 = GetN2 + CombCompo(i)
    End If
Next i


End Function

Function GetSO2(Ngaz As Integer, CombNames() As String, CombCompo() As Double)

GetSO2 = 0

' Recherche du % volumique de SO2 dans le comburant
Dim j As Integer

j = 0
For i = 1 To Ngaz
    If CombNames(i) = "SO2" Then
        j = j + 1
        GetSO2 = GetSO2 + CombCompo(i)
    End If
Next i

End Function
Function GetMasseMolgaz(NomGaz As String) As Double


    GetMasseMolgaz = ThisWorkbook.Sheets(Sprop).Range("B3").Offset(GetIndexGaz(NomGaz), 6)

End Function
Function GetIndexGaz(NomGaz As String) As Double

Dim i As Integer

'Indices de la table dans la feuille "basdo_gaz"
Select Case NomGaz
    Case "O2"
        GetIndexGaz = 1
    Case "CO2"
        GetIndexGaz = 2
    Case "H2O"
        GetIndexGaz = 3
    Case "N2"
        GetIndexGaz = 4
    Case "SO2"
        GetIndexGaz = 5
    Case "H2"
        GetIndexGaz = 6
    Case "CH4"
        GetIndexGaz = 7
    Case "C2H6"
        GetIndexGaz = 8
    Case "C3H8"
        GetIndexGaz = 9
    Case "C4H10"
        GetIndexGaz = 10
    Case "C5H12"
        GetIndexGaz = 11
    Case "C6H14"
        GetIndexGaz = 12
    Case "C7H16"
        GetIndexGaz = 13
    Case "C8H18"
        GetIndexGaz = 14
    Case "C2H2"
        GetIndexGaz = 15
    Case "C2H4"
        GetIndexGaz = 16
    Case "C3H6"
        GetIndexGaz = 17
    Case "C4H8"
        GetIndexGaz = 18
    Case "C5H10"
        GetIndexGaz = 19
    Case "C6H12"
        GetIndexGaz = 20
    Case "C6H6"
        GetIndexGaz = 21
    Case "C7H8"
        GetIndexGaz = 22
    Case "CO"
        GetIndexGaz = 23
    Case "H2S"
        GetIndexGaz = 24
    Case "C19H30"
        GetIndexGaz = 25
    Case Else
       GetIndexGaz = -1
End Select

End Function
Function GetFracMass(Ngaz As Integer, CombNames() As String, CombCompo() As Double, iGaz As Integer)

'Calcul de la fraction massique d'un élément d'un composé
'iGaz est l'indice du composé

Dim MasseMol As Double 'Masse molaire du composé

MasseMol = 0

For i = 1 To Ngaz
    MasseMol = MasseMol + CombCompo(i) / 100 * GetMasseMolgaz(CombNames(i))
Next i

GetFracMass = CombCompo(iGaz) / 100 * GetMasseMolgaz(CombNames(iGaz)) / MasseMol

End Function
Function GasTemp(Ngaz As Integer, CombNames() As String, CombCompo() As Double, Puissance As Double, Debit As Double)

    '---------------------------------------------------------------
    ' Calcul de la température d'un fluide connaissant :
    '  - La composition en %vol.
    '  - La puissance apportée en W
    '  - Le débit en Nm3/h
    '---------------------------------------------------------------
    ' La température renvoyée est en Celsius, calcul par dichotomie
    '---------------------------------------------------------------
    ' L. FERRANDanvier 2009
    '---------------------------------------------------------------

    Dim Enthalpie As Double ' en J/kg

    Dim Hmin As Double 'en J/kg
    Dim Hmax As Double ' en J/kg
    Dim Tmin As Double ' en K
    Dim Tmax As Double ' en K
    Dim i As Integer

    '-----Calcul de l'enthalpie du fluide

    Enthalpie = Puissance / (Debit / 3600 * CalculMasseVolumique(Ngaz, CombNames, CombCompo, 0, 101325, "[kg/Nm3]"))

    Hmin = 0
    Hmax = 4000000
    Tmin = 0
    Tmax = 2000
    i = 0

    While Abs(Enthalpie - Hmin) > 0.001 * Enthalpie And i < 50
        Hmin = CalculEnthalp(Ngaz, CombNames, CombCompo, Tmin, 0, "[J/kg]")
        Hmax = CalculEnthalp(Ngaz, CombNames, CombCompo, Tmax, 0, "[J/kg]")

        If Enthalpie > 0.5 * (Hmin + Hmax) Then
            Tmin = 0.75 * Tmin + 0.25 * Tmax
        Else
            Tmax = 0.25 * Tmin + 0.75 * Tmax
        End If
    Wend

    If i < 50 Then
        CalculTemperature = 0.5 * (Tmin + Tmax)
    Else
        CalculTemperature = -1
        Call MsgBox("Problème de convergence de l'enthalpie", vbAbortRetryIgnore)
    End If

End Function
