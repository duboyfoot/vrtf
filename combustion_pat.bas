Attribute VB_Name = "combustion_pat"
Public Function GasPCI(Name As String) As Double

    Dim Nbgas As Integer
    Dim Gasname() As String
    Dim gasCompo() As Double


    Call Gaz_GetCompo(SFuel, Name, Nbgas, Gasname, gasCompo)
    
    GasPCI = CalculPCI(Nbgas, Gasname, gasCompo, "[kJ/Nm3]")
    
    

End Function
Public Function GasPCIKG(Name As String) As Double

    Dim Nbgas As Integer
    Dim Gasname() As String
    Dim gasCompo() As Double


    Call Gaz_GetCompo(SFuel, Name, Nbgas, Gasname, gasCompo)
    
    GasPCIKG = CalculPCI(Nbgas, Gasname, gasCompo, "[J/kg]")
    
    

End Function
Function CalculPCI(Ngaz As Integer, CombNames() As String, CombCompo() As Double, Unite As String)

CalculPCI = 0

' Calcul du PCI du mélange en J/kmol
For i = 1 To Ngaz
    CalculPCI = CalculPCI + CombCompo(i) / 100 * GetPCIgaz(CombNames(i))
Next i

' Conversion dans l'unité voulue
Select Case Unite
    Case "[J/kg]"
        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo())
    Case "[kJ/kg]"
        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000
    Case "[kJ/Nm3]"
        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 273, 101325, "[kg/m3]")
    Case "[kcal/kg]"
        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * 0.239
    Case "[kcal/Nm3]"
        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 273, 101325, "[kg/m3]") * 0.239
    Case "[kWh/Nm3]"
        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 273, 101325, "[kg/m3]") * 0.0002778
    Case "[kWh/kg]"
        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * 0.0002778
    Case "[BTU/scf]"
        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 15 + 273, 101325, "[kg/m3]") * 0.9485 / 35.31467
    Case "[kcal/mol]"
        CalculPCI = CalculPCI * 0.000239 / 1000
    Case Else
       CalculPCI = -1
End Select

End Function
Public Function GasVa0(Gasname As String) As Double
    Dim Nbgas As Integer
    Dim FuelCompo() As Double
    Dim FuelName() As String
    Dim nbFuel As Integer
    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim NameAir As String
    NameAir = "air"
    Call Gaz_GetCompo(SFuel, Gasname, nbFuel, FuelName, FuelCompo)
    If (nbFuel = 0) Then
        GasVa0 = -1
        Exit Function
    End If
    Call Gaz_GetCompo(SComburant, NameAir, nbAir, AirName, AirCompo)


    GasVa0 = CalculVa(nbFuel, FuelName, FuelCompo, nbAir, AirName, AirCompo, "[Nm3/Nm3]")
 
    

End Function
Public Function GasVf0(Gasname As String, AirGazRatio) As Double
    Dim Nbgas As Integer
    Dim FuelCompo() As Double
    Dim FuelName() As String
    Dim nbFuel As Integer
    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim NameAir As String
    NameAir = "air"
    Call Gaz_GetCompo(SFuel, Gasname, nbFuel, FuelName, FuelCompo)
    If (nbFuel = 0) Then
        GasVf0 = -1
        Exit Function
    End If
    Call Gaz_GetCompo(SComburant, NameAir, nbAir, AirName, AirCompo)
    GasVf0 = (AirGazRatio - 1) * CalculVa(nbFuel, FuelName, FuelCompo, nbAir, AirName, AirCompo, "[Nm3/Nm3]") + CalculVfh(nbFuel, FuelName, FuelCompo, nbAir, AirName, AirCompo, "[Nm3/Nm3]")

    

End Function
Public Function GasVa0KG(Gasname As String) As Double
    Dim Nbgas As Integer
    Dim FuelCompo() As Double
    Dim FuelName() As String
    Dim nbFuel As Integer
    Dim nbAir As Integer
    Dim AirName() As String
    Dim AirCompo() As Double
    Dim NameAir As String
    NameAir = "air"
    Call Gaz_GetCompo(SFuel, Gasname, nbFuel, FuelName, FuelCompo)
    If (nbFuel = 0) Then
        GasVa0KG = -1
        Exit Function
    End If
    Call Gaz_GetCompo(SComburant, NameAir, nbAir, AirName, AirCompo)


    GasVa0KG = CalculVa(nbFuel, FuelName, FuelCompo, nbAir, AirName, AirCompo, "[kg/kg]")
 
    

End Function
Function CalculVa(Ngaz_fuel As Integer, FuelNames() As String, FuelCompo() As Double, Ngaz_Air As Integer, AirNames() As String, AirCompo() As Double, Unite As String)

' Calcul du pouvoir comburivore en Unite de comburant par Unite de combustible

CalculVa = 0

'Calcul du VO2 du combustible en Nm3 d'O2/Nm3 de combustible
For i = 1 To Ngaz_fuel
    CalculVa = CalculVa + FuelCompo(i) / 100 * GetVO2gaz(FuelNames(i))
Next i

If GetO2(Ngaz_Air, AirNames(), AirCompo()) = 0 Then
    MsgBox "Le comburant doit nécessairement contenir de l'oxygène"
    CalculVa = -1
    Exit Function
Else
    'Calcul du pouvoir comburivore à partir de l'O2 présent dans le comburant
    CalculVa = CalculVa / (GetO2(Ngaz_Air, AirNames(), AirCompo()) / 100)
End If

If Unite = "[Nm3/kg]" Then
    CalculVa = CalculVa / CalculMasseVolumique(Ngaz_fuel, FuelNames, FuelCompo, 273, 101325, "[kg/m3]")
ElseIf Unite = "[scf/lb]" Then
    CalculVa = CalculVa / CalculMasseVolumique(Ngaz_fuel, FuelNames, FuelCompo, 15 + 273, 101325, "[kg/m3]")
ElseIf Unite = "[kg/Nm3]" Then
    CalculVa = CalculVa * CalculMasseVolumique(Ngaz_Air, AirNames, AirCompo, 273, 101325, "[kg/m3]")
ElseIf Unite = "[lb/scf]" Then
    CalculVa = CalculVa * CalculMasseVolumique(Ngaz_Air, AirNames, AirCompo, 15 + 327, 101325, "[kg/m3]")
ElseIf Unite = "[kg/kg]" Or Unite = "[lb/lb]" Then
    CalculVa = CalculVa * CalculMasseVolumique(Ngaz_Air, AirNames, AirCompo, 273, 101325, "[kg/m3]") _
                    / CalculMasseVolumique(Ngaz_fuel, FuelNames, FuelCompo, 273, 101325, "[kg/m3]")
End If

End Function
Function GetVO2gaz(NomGaz As String)

'Spropt est le nom de la feuille des propriétés
Dim i As Integer
GetVO2gaz = -100
i = 0

While GetVO2gaz = -100 And i < 26

    If ThisWorkbook.Sheets(Sprop).Range("B3").Offset(i) = NomGaz Then
        GetVO2gaz = 0.5 * (2 * ThisWorkbook.Sheets(Sprop).Range("B3").Offset(i, 1) + 0.5 * ThisWorkbook.Sheets(Sprop).Range("B3").Offset(i, 2) - ThisWorkbook.Sheets(Sprop).Range("B3").Offset(i, 3) + 2 * ThisWorkbook.Sheets(Sprop).Range("B3").Offset(i, 5))
    Else
        i = i + 1
    End If
Wend

If GetVO2gaz = -100 Then
    Call MsgBox("Le gaz pur : " & NomGaz & " n'existe pas dans la base.", vbCritical)
End If

End Function
Function GetPCIgaz(NomGaz As String)

    ' Recherche du gaz dans la liste et récupération de son PCI dans la colonne J
    GetPCIgaz = ThisWorkbook.Sheets(Sprop).Range("B3").Offset(GetIndexGaz(NomGaz), 8)

End Function
Function CalculMasseMolaire(Ngaz As Integer, CombNames() As String, CombCompo() As Double)

'Ngaz : nombre de gaz contenus dans le composé
'CombNames() : nom de gaz élémentaires contenus dans le composé - l'index zéro n'est pas utilisé
'CombCompo() : pourcentage volumique [%] de chacun des gaz - l'index zéro n'est pas utilisé

CalculMasseMolaire = 0

' Calcul de la masse molaire d'un composé gazeux en kg/kmol
For i = 1 To Ngaz
    CalculMasseMolaire = CalculMasseMolaire + CombCompo(i) / 100 * GetMasseMolgaz(CombNames(i))
Next i

End Function



Function GetMasseMolgaz(NomGaz As String)


    GetMasseMolgaz = ThisWorkbook.Sheets(Sprop).Range("B3").Offset(GetIndexGaz(NomGaz), 6)

End Function
Function CalculVfh(Ngaz_fuel As Integer, FuelNames() As String, FuelCompo() As Double, Ngaz_Air As Integer, AirNames() As String, AirCompo() As Double, Unite As String)

' Calcul du pouvoir fumigène humide stoechiométrique en Nm3 de fumées par Nm3 de combustible
Dim FumNames(5) As String
Dim FumCompo(5) As Double
Dim prop As String


CalculVfh = 0

For i = 1 To Ngaz_fuel
    'Etape 1 : calcul du volume de CO2 généré par la combustion
    CalculVfh = CalculVfh + FuelCompo(i) / 100 * GetVCO2gaz(FuelNames(i), prop)
    'Etape 2 : calcul du volume de H2O généré par la combustion
    CalculVfh = CalculVfh + FuelCompo(i) / 100 * GetVH2Ogaz(FuelNames(i), prop)
    'Etape 3 : calcul du volume de SO2 généré par la combustion
    CalculVfh = CalculVfh + FuelCompo(i) / 100 * GetVSO2gaz(FuelNames(i), prop)
    'Etape 4 : calcul du volume de N2 provenant du combustible
    CalculVfh = CalculVfh + FuelCompo(i) / 100 * GetVN2gaz(FuelNames(i), prop)
Next i

'Etape 5 : calcul du volume de N2 provenant du comburant
CalculVfh = CalculVfh + CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]") * GetN2(Ngaz_Air, AirNames(), AirCompo()) / 100
'Etape 6 : calcul du volume de CO2 provenant du comburant
CalculVfh = CalculVfh + CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]") * GetCO2(Ngaz_Air, AirNames(), AirCompo()) / 100
'Etape 7 : calcul du volume de H2O provenant du comburant
CalculVfh = CalculVfh + CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]") * GetH2O(Ngaz_Air, AirNames(), AirCompo()) / 100
'Etape 8 : calcul du volume de SO2 provenant du comburant
CalculVfh = CalculVfh + CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]") * GetSO2(Ngaz_Air, AirNames(), AirCompo()) / 100



If Unite = "[Nm3/kg]" Then
    CalculVfh = CalculVfh / CalculMasseVolumique(Ngaz_fuel, FuelNames, FuelCompo, 273, 101325, "[kg/m3]")
ElseIf Unite = "[scf/lb]" Then
    CalculVfh = CalculVfh / CalculMasseVolumique(Ngaz_fuel, FuelNames, FuelCompo, 273 + 15, 101325, "[kg/m3]")

ElseIf Unite = "[kg/Nm3]" Then
    FumNames(1) = "CO2"
    FumNames(2) = "H2O"
    FumNames(3) = "O2"
    FumNames(4) = "N2"
    FumNames(5) = "SO2"
    
    FumCompo(1) = CalculFracVolCO2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(2) = CalculFracVolH2O(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(3) = CalculFracVolO2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(4) = CalculFracVolN2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(5) = CalculFracVolSO2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    
    CalculVfh = CalculVfh * CalculMasseVolumique(5, FumNames, FumCompo, 273, 101325, "[kg/m3]")
ElseIf Unite = "[lb/scf]" Then
    FumNames(1) = "CO2"
    FumNames(2) = "H2O"
    FumNames(3) = "O2"
    FumNames(4) = "N2"
    FumNames(5) = "SO2"
    
    FumCompo(1) = CalculFracVolCO2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(2) = CalculFracVolH2O(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(3) = CalculFracVolO2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(4) = CalculFracVolN2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(5) = CalculFracVolSO2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    
    CalculVfh = CalculVfh * CalculMasseVolumique(5, FumNames, FumCompo, 273 + 15, 101325, "[kg/m3]")
ElseIf Unite = "[kg/kg]" Or Unite = "[lb/lb]" Then
    FumNames(1) = "CO2"
    FumNames(2) = "H2O"
    FumNames(3) = "O2"
    FumNames(4) = "N2"
    FumNames(5) = "SO2"
    
    FumCompo(1) = CalculFracVolCO2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(2) = CalculFracVolH2O(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(3) = CalculFracVolO2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(4) = CalculFracVolN2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    FumCompo(5) = CalculFracVolSO2(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, 1)
    
    
    CalculVfh = CalculVfh * CalculMasseVolumique(5, FumNames, FumCompo, 273, 101325, "[kg/m3]") _
                / CalculMasseVolumique(Ngaz_fuel, FuelNames, FuelCompo, 273, 101325, "[kg/m3]")
End If

End Function

Function GetVCO2gaz(NomGaz As String, PropSheet As String)

'PropSheet est le nom de la feuille des propriétés

GetVCO2gaz = ThisWorkbook.Sheets(Sprop).Range("B3").Offset(GetIndexGaz(NomGaz), 1)

End Function
Function GetVH2Ogaz(NomGaz As String, PropSheet As String)

    GetVH2Ogaz = 0.5 * ThisWorkbook.Sheets(Sprop).Range("B3").Offset(GetIndexGaz(NomGaz), 2)

End Function
Function GetVSO2gaz(NomGaz As String, PropSheet As String)


    GetVSO2gaz = ThisWorkbook.Sheets(Sprop).Range("B3").Offset(GetIndexGaz(NomGaz), 5)
    
End Function

Function GetVN2gaz(NomGaz As String, PropSheet As String)
    GetVN2gaz = 0.5 * ThisWorkbook.Sheets(Sprop).Range("B3").Offset(GetIndexGaz(NomGaz), 4)
    
End Function
Function CalculFracVolCO2(Ngaz_fuel As Integer, FuelNames() As String, FuelCompo() As Double, Ngaz_Air As Integer, AirNames() As String, AirCompo() As Double, AirFuelRatio As Double)

'Calcul de la fraction volumique de CO2 dans les fumées issue de la combustion

    CalculFracVolCO2 = 0
    
    'Calcul de la fraction volumique de CO2 présente dans les fumées

    'Etape 1 : calcul du volume de CO2 généré par la combustion
    For i = 1 To Ngaz_fuel
        CalculFracVolCO2 = CalculFracVolCO2 + FuelCompo(i) / 100 * GetVCO2gaz(FuelNames(i), prop)
    Next i

    'Etape 2 : calcul du volume de CO2 provenant du comburant
    CalculFracVolCO2 = CalculFracVolCO2 + AirFuelRatio * GetCO2(Ngaz_Air, AirNames(), AirCompo()) / 100 * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]")

    'Etape 3 : Division par le volume de fumées total
    CalculFracVolCO2 = 100 * CalculFracVolCO2 / (CalculVfh(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]"))

End Function

Function CalculFracVolH2O(Ngaz_fuel As Integer, FuelNames() As String, FuelCompo() As Double, Ngaz_Air As Integer, AirNames() As String, AirCompo() As Double, AirFuelRatio As Double)

'Calcul de la fraction volumique de H2O dans les fumées issue de la combustion

    CalculFracVolH2O = 0
    'Calcul de la fraction volumique de H2O présente dans les fumées


    'Etape 1 : calcul du volume de H2O généré par la combustion
    For i = 1 To Ngaz_fuel
        CalculFracVolH2O = CalculFracVolH2O + FuelCompo(i) / 100 * GetVH2Ogaz(FuelNames(i), prop)
    Next i

    'Etape 2 : calcul du volume de H2O provenant du comburant
    CalculFracVolH2O = CalculFracVolH2O + AirFuelRatio * GetH2O(Ngaz_Air, AirNames(), AirCompo()) / 100 * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]")

    'Etape2 : Division par le volume de fumées total
    CalculFracVolH2O = 100 * CalculFracVolH2O / (CalculVfh(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]"))

End Function

Function CalculFracVolN2(Ngaz_fuel As Integer, FuelNames() As String, FuelCompo() As Double, Ngaz_Air As Integer, AirNames() As String, AirCompo() As Double, AirFuelRatio As Double)

'Calcul de la fraction volumique de O2 dans les fumées issue de la combustion


    CalculFracVolN2 = 0
    'Calcul de la fraction volumique de N2 présente dans les fumées

    'Etape 1 : calcul du volume de N2 généré par la combustion
    For i = 1 To Ngaz_fuel
        CalculFracVolN2 = CalculFracVolN2 + FuelCompo(i) / 100 * GetVN2gaz(FuelNames(i), prop)
    Next i

    'Etape 2 : calcul du volume de N2 provenant du comburant
    CalculFracVolN2 = CalculFracVolN2 + AirFuelRatio * GetN2(Ngaz_Air, AirNames(), AirCompo()) / 100 * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]")

    'Etape3 : Division par le volume de fumées total
    CalculFracVolN2 = 100 * CalculFracVolN2 / (CalculVfh(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]"))

End Function

Function CalculFracVolO2(Ngaz_fuel As Integer, FuelNames() As String, FuelCompo() As Double, Ngaz_Air As Integer, AirNames() As String, AirCompo() As Double, AirFuelRatio As Double)
    
    'Calcul de la fraction volumique de O2 présente dans les fumées

    CalculFracVolO2 = 0

    'Etape 1 : calcul du volume de O2 provenant de l'excès d'air
    CalculFracVolO2 = CalculFracVolO2 + (AirFuelRatio - 1) * GetO2(Ngaz_Air, AirNames(), AirCompo()) / 100 * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]")

    'Etape 2 : Division par le volume de fumées total
    CalculFracVolO2 = 100 * CalculFracVolO2 / (CalculVfh(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]"))

End Function

Function CalculFracVolSO2(Ngaz_fuel As Integer, FuelNames() As String, FuelCompo() As Double, Ngaz_Air As Integer, AirNames() As String, AirCompo() As Double, AirFuelRatio As Double)
'Calcul de la fraction volumique de SO2 présente dans les fumées

    CalculFracVolSO2 = 0

    'Etape 1 : calcul du volume de SO2 généré par la combustion
    For i = 1 To Ngaz_fuel
        CalculFracVolSO2 = CalculFracVolSO2 + FuelCompo(i) / 100 * GetVSO2gaz(FuelNames(i), prop)
    Next i

    'Etape 2 : calcul du volume de SO2 provenant du comburant
    CalculFracVolSO2 = CalculFracVolSO2 + AirFuelRatio * GetSO2(Ngaz_Air, AirNames(), AirCompo()) / 100 * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]")

    'Etape2 : Division par le volume de fumées total
    CalculFracVolSO2 = 100 * CalculFracVolSO2 / (CalculVfh(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), "[Nm3/Nm3]"))


End Function
