Attribute VB_Name = "ModuleCombustion"

'Function CalculMasseMolaire(Ngaz As Integer, CombNames() As String, CombCompo() As Double)
'
''Ngaz : nombre de gaz contenus dans le composé
''CombNames() : nom de gaz élémentaires contenus dans le composé - l'index zéro n'est pas utilisé
''CombCompo() : pourcentage volumique [%] de chacun des gaz - l'index zéro n'est pas utilisé
'
'CalculMasseMolaire = 0
'
'' Calcul de la masse molaire d'un composé gazeux en kg/kmol
'For i = 1 To Ngaz
'    CalculMasseMolaire = CalculMasseMolaire + CombCompo(i) / 100 * GetMasseMolgaz(CombNames(i))
'Next i
'
'End Function
'
'
'
'Function GetMasseMolgaz(Nomgaz As String)
'
'GetMasseMolgaz = 0
'
'' Recherche du gaz dans la liste et récupération de sa masse molaire dans la colonne H
'Dim i As Integer
'i = 1
'While GetMasseMolgaz = 0 And i < 25
'    If Worksheets(SProp).Range("B3").Offset(i) = Nomgaz Then
'        GetMasseMolgaz = Worksheets(SProp).Range("B3").Offset(i, 6)
'    Else
'        i = i + 1
'    End If
'
'Wend
'
'If i = 25 Then
'    MsgBox ("Erreur le composé :" & Nomgaz & " n'existe pas dans la base de donnée")
'    GetMasseMolgaz = 0
'End If
'
'End Function
'
'Function GetFracMass(Ngaz As Integer, CombNames() As String, CombCompo() As Double, iGaz As Integer)
'
''Calcul de la fraction massique d'un élément d'un composé
''iGaz est l'indice du composé
'
'Dim MasseMol As Double 'Masse molaire du composé
'
'MasseMol = 0
'
'For i = 1 To Ngaz
'    MasseMol = MasseMol + CombCompo(i) / 100 * GetMasseMolgaz(CombNames(i))
'Next i
'
'GetFracMass = CombCompo(iGaz) / 100 * GetMasseMolgaz(CombNames(iGaz)) / MasseMol
'
'End Function
'
'Function CalculViscositeDyn(Ngaz As Integer, CombNames() As String, CombCompo() As Double, Tcelsius As Double)
''Calcul de la viscosité dynamique d'un composé constitué de Ngaz composants à la température Tcelsius [°C]
'
''Ngaz : nombre de gaz contenus dans le composé
''CombNames() : nom de gaz élémentaires contenus dans le composé - l'index zéro n'est pas utilisé
''CombCompo() : pourcentage volumique [%] de chacun des gaz - l'index zéro n'est pas utilisé
'
'Dim i As Integer
'Dim Visco As Double
'
'CalculViscositeDyn = 0
'
'For i = 1 To Ngaz
'
'    Visco = GetViscositeDyn(CombNames(i), Tcelsius + 273)
'    If Visco = -1 Then 'Erreur base de données
'        CalculVisositeDyn = -1
'        Exit For
'    Else
'        CalculViscositeDyn = CalculViscositeDyn + GetFracMass(Ngaz, CombNames(), CombCompo(), i) * Visco
'
'    End If
'
'
'Next i
'
'End Function
'
'
'Function GetViscositeDyn(GazPur As String, TKelvin As Double)
''Recherche de la viscosité dynamique d'un gaz pur à la température Tkelvin [K]
'
'GetViscositeDyn = 0
'
'Dim i As Integer
'
'' Recherche du gaz dans la liste et récupération de sa viscosité dynamique à la température spécifiée
'
'i = 3
'While GetViscositeDyn = 0
'    If Worksheets(Viscosite).Cells(8, i) = GazPur Then
'        'On balaye les lignes jusqu'à dépasser la température voulue
'        j = 1
'        While Worksheets(Viscosite).Cells(8 + j, 2) < TKelvin And GetViscositeDyn <> -1
'            If Worksheets(Viscosite).Cells(8 + j, 2) = "" Then
'                Call MsgBox("La température de service (" & TKelvin - 273 & "°C) est supérieure au domaine la base de données (max. 680°C)", vbExclamation, "Erreur de données")
'                GetViscositeDyn = -1
'            Else
'                j = j + 1
'            End If
'        Wend
'
'        If j = 1 Then
'            Call MsgBox("La température de service (" & TKelvin - 273 & "°C) est inférieure au domaine la base de données (min. -20°C)", vbExclamation, "Erreur de données")
'            GetViscositeDyn = -1
'        End If
'
'        If GetViscositeDyn <> -1 Then GetViscositeDyn = Interpolate(Worksheets(Viscosite).Cells(8 + j - 1, 2), Worksheets(Viscosite).Cells(8 + j, 2), Worksheets(Viscosite).Cells(8 + j - 1, i), Worksheets(Viscosite).Cells(8 + j, i), TKelvin)
'
'    Else
'        i = i + 1
'    End If
'
'
'Wend
'
'End Function
'
'
'
'Function GetPsaturation(TKelvin As Double)
'
'    Dim i As Integer
'    GetPsaturation = 0
'
'    ' Récupération de la pression de saturation [pa] à la température spécifiée
'
'    i = 9
'    While GetPsaturation = 0
'        While Worksheets(Satur).Cells(i, 2) <= TKelvin And GetPsaturation <> -1
'            If Worksheets(Satur).Cells(i, 2) = "" Then
'                Call MsgBox("La température de calcul de l'humidité (" & TKelvin - 273 & "°C) est supérieure au domaine la base de données (max. 200°C)", vbExclamation, "Erreur de données")
'                GetPsaturation = -1
'            Else
'                i = i + 1
'            End If
'        Wend
'
'        If i = 9 Then
'            Call MsgBox("La température de calcul de l'humidité (" & TKelvin - 273 & "°C) est inférieure au domaine la base de données (min. 0°C)", vbExclamation, "Erreur de données")
'            GetPsaturation = -1
'        End If
'
'        If GetPsaturation <> -1 Then GetPsaturation = Interpolate(Worksheets(Satur).Cells(i - 1, 2), Worksheets(Satur).Cells(i, 2), Worksheets(Satur).Cells(i - 1, 3), Worksheets(Satur).Cells(i, 2), TKelvin)
'    Wend
'End Function
'
'Function CalculPCI(Ngaz As Integer, CombNames() As String, CombCompo() As Double, unite As String)
'
'CalculPCI = 0
'
'' Calcul du PCI du mélange en J/kmol
'For i = 1 To Ngaz
'    CalculPCI = CalculPCI + CombCompo(i) / 100 * GetPCIgaz(CombNames(i))
'Next i
'
'' Conversion dans l'unité voulue
'Select Case unite
'    Case "[kJ/kg]"
'        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000
'    Case "[kJ/Nm3]"
'        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 0, 101325, "[kg/m3]")
'    Case "[kcal/kg]"
'        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * 0.239
'    Case "[kcal/Nm3]"
'        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 0, 101325, "[kg/m3]") * 0.239
'    Case "[kWh/Nm3]"
'        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 0, 101325, "[kg/m3]") * 0.0002778
'    Case "[kWh/kg]"
'        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * 0.0002778
'    Case "[BTU/scf]"
'        CalculPCI = CalculPCI / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 0, 101325, "[kg/m3]") * 0.9485 / 35.31467
'    Case Else
'       CalculPCI = -1
'End Select
'
'End Function
'
'Function CalculPCS(Ngaz As Integer, CombNames() As String, CombCompo() As Double, unite As String)
'
'CalculPCS = 0
'
'' Calcul du PCS du mélange en J/kmol
'For i = 1 To Ngaz
'    CalculPCS = CalculPCS + CombCompo(i) / 100 * GetPCSgaz(CombNames(i))
'Next i
'
'' Conversion dans l'unité voulue
'Select Case unite
'    Case "[kJ/kg]"
'        CalculPCS = CalculPCS / CalculMasseMolaire(Ngaz, CombNames, CombCompo) / 1000
'    Case "[kJ/Nm3]"
'        CalculPCS = CalculPCS / CalculMasseMolaire(Ngaz, CombNames, CombCompo) / 1000 * CalculMasseVolumique(Ngaz, CombNames, CombCompo, 0, 101325, "[kg/m3]")
'    Case "[kcal/kg]"
'        CalculPCS = CalculPCS / CalculMasseMolaire(Ngaz, CombNames, CombCompo) / 1000 * 0.239
'    Case "[kcal/Nm3]"
'        CalculPCS = CalculPCS / CalculMasseMolaire(Ngaz, CombNames, CombCompo) / 1000 * CalculMasseVolumique(Ngaz, CombNames, CombCompo, 0, 101325, "[kg/m3]") * 0.239
'    Case "[kWh/Nm3]"
'        CalculPCS = CalculPCS / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 0, 101325, "[kg/m3]") * 0.0002778
'    Case "[kWh/kg]"
'        CalculPCS = CalculPCS / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * 0.0002778
'    Case "[BTU/scf]"
'        CalculPCS = CalculPCS / CalculMasseMolaire(Ngaz, CombNames(), CombCompo()) / 1000 * CalculMasseVolumique(Ngaz, CombNames(), CombCompo(), 0, 101325, "[kg/m3]") * 0.9485 / 35.31467
'    Case Else
'       CalculPCS = -1
'End Select
'
'End Function
'
'Function GetPCIgaz(Nomgaz As String)
'
'    GetPCIgaz = -1
'    Dim i As Integer
'
'    i = 1
'    ' Recherche du gaz dans la liste et récupération de son PCI dans la colonne J
'    While GetPCIgaz = -1 And i < 28
'        If Worksheets(SProp).Range("B3").Offset(i) = Nomgaz Then
'            GetPCIgaz = Worksheets(SProp).Range("B3").Offset(i, 8)
'        Else
'            i = i + 1
'        End If
'    Wend
'
'    If GetPCIgaz = -1 Then Call MsgBox("Le gaz " & Nomgaz & " n'a pas été trouvé dans la base", vbCritical)
'
'End Function
'
'Function GetPCSgaz(Nomgaz As String)
'
'    GetPCSgaz = -1
'    Dim i As Integer
'
'    i = 1
'    ' Recherche du gaz dans la liste et récupération de son PCS dans la colonne K
'    While GetPCSgaz = -1 And i < 28
'        If Worksheets(SProp).Range("B3").Offset(i) = Nomgaz Then
'            GetPCSgaz = Worksheets(SProp).Range("B3").Offset(i, 9)
'        Else
'            i = i + 1
'        End If
'    Wend
'
'    If GetPCSgaz = -1 Then Call MsgBox("Le gaz " & Nomgaz & " n'a pas été trouvé dans la base", vbCritical)
'
'End Function
'
'Function CalculMasseVolumique(Ngaz As Integer, CombNames() As String, CombCompo() As Double, Tcelsius As Double, Pabsolue As Double, unite As String)
''Calcul de la masse volumique d'un composé constitué de Ngaz composants à la température Tcelsius
'' et à la pression Pabsolue en Pa
'
''Ngaz : nombre de gaz contenus dans le composé
''CombNames() : nom de gaz élémentaires contenus dans le composé - l'index zéro n'est pas utilisé
''CombCompo() : pourcentage volumique [%] de chacun des gaz - l'index zéro n'est pas utilisé
'    Dim i As Integer
'
'    CalculMasseVolumique = 0
'
'    For i = 1 To Ngaz
'        CalculMasseVolumique = CalculMasseVolumique + GetMasseMolgaz(CombNames(i)) * CombCompo(i) / 100
'    Next i
'
'    CalculMasseVolumique = CalculMasseVolumique / 22.4136 * 273 / (273 + Tcelsius) * Pabsolue / 101325
'
'    If unite = "[lb/cf]" Then
'        CalculMasseVolumique = CalculMasseVolumique * 2.20462 / 35.31467
'    End If
'End Function
'
'
'Function CalculWobbe(PCS_kJNm3 As Double, d As Double, Ngaz As Integer, CombNames() As String, CombCompo() As Double, unite As String)
'
'CalculWobbe = 0
'
'' Conversion dans l'unité voulue
'Select Case unite
'    Case "[MJ/Nm3]"
'        CalculWobbe = PCS_kJNm3 / Sqr(d) / 1000
'    Case "[MJ/kg]"
'        CalculWobbe = PCS_kJNm3 / Sqr(d) / 1000 * CalculMasseVolumique(Ngaz, CombNames, CombCompo, 0, 101325, "kg/m3")
'    Case "[Mcal/Nm3]"
'        CalculWobbe = PCS_kJNm3 / Sqr(d) / 1000 * 0.239
'    Case "[Mcal/kg]"
'        CalculWobbe = PCS_kJNm3 / Sqr(d) / 1000 * CalculMasseVolumique(Ngaz, CombNames, CombCompo, 0, 101325, "kg/m3") * 0.239
'    Case "[BTU/scf]"
'        CalculWobbe = PCS_kJNm3 / Sqr(d) / 1000 * 948.5 * 273 / (273 + 15)
'    Case "[BTU/lb]"
'        CalculWobbe = PCS_kJNm3 / Sqr(d) / 1000 * 948.5 * 273 / (273 + 15) * CalculMasseVolumique(Ngaz, CombNames, CombCompo, 0, 101325, "lb/cf")
'    Case Else
'       CalculWobbe = -1
'End Select
'
'End Function
'
'Function CalculVa(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, unite As String)
'
'' Calcul du pouvoir comburivore en Unite de comburant par Unite de combustible
'
'CalculVa = 0
'
''Calcul du VO2 du combustible en Nm3 d'O2/Nm3 de combustible
'For i = 1 To Ngaz_fuel
'    CalculVa = CalculVa + Fuelcompo(i) / 100 * GetVO2gaz(Fuelnames(i), SProp)
'Next i
'
'If GetO2(Ngaz_Air, Airnames(), Aircompo()) = 0 Then
'    MsgBox "Le comburant doit nécessairement contenir de l'oxygène"
'    CalculVa = -1
'    Exit Function
'Else
'    'Calcul du pouvoir comburivore à partir de l'O2 présent dans le comburant
'    CalculVa = CalculVa / (GetO2(Ngaz_Air, Airnames(), Aircompo()) / 100)
'End If
'
'If unite = "[Nm3/kg]" Then
'    CalculVa = CalculVa / CalculMasseVolumique(Ngaz_fuel, Fuelnames, Fuelcompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[scf/lb]" Then
'    CalculVa = CalculVa / CalculMasseVolumique(Ngaz_fuel, Fuelnames, Fuelcompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[kg/Nm3]" Then
'    CalculVa = CalculVa * CalculMasseVolumique(Ngaz_Air, Airnames, Aircompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[lb/scf]" Then
'    CalculVa = CalculVa * CalculMasseVolumique(Ngaz_Air, Airnames, Aircompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[kg/kg]" Or unite = "[lb/lb]" Then
'    CalculVa = CalculVa * CalculMasseVolumique(Ngaz_Air, Airnames, Aircompo, 0, 101325, "[kg/m3]") _
'                    / CalculMasseVolumique(Ngaz_fuel, Fuelnames, Fuelcompo, 0, 101325, "[kg/m3]")
'End If
'
'End Function
'
'Function CalculVfh(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, unite As String)
'
'' Calcul du pouvoir fumigène humide stoechiométrique en Nm3 de fumées par Nm3 de combustible
'Dim Fumnames(5) As String
'Dim Fumcompo(5) As Double
'
'
'
'CalculVfh = 0
'
'For i = 1 To Ngaz_fuel
'    'Etape 1 : calcul du volume de CO2 généré par la combustion
'    CalculVfh = CalculVfh + Fuelcompo(i) / 100 * GetVCO2gaz(Fuelnames(i), SProp)
'    'Etape 2 : calcul du volume de H2O généré par la combustion
'    CalculVfh = CalculVfh + Fuelcompo(i) / 100 * GetVH2Ogaz(Fuelnames(i), SProp)
'    'Etape 3 : calcul du volume de SO2 généré par la combustion
'    CalculVfh = CalculVfh + Fuelcompo(i) / 100 * GetVSO2gaz(Fuelnames(i), SProp)
'    'Etape 4 : calcul du volume de N2 provenant du combustible
'    CalculVfh = CalculVfh + Fuelcompo(i) / 100 * GetVN2gaz(Fuelnames(i), SProp)
'Next i
'
''Etape 5 : calcul du volume de N2 provenant du comburant
'CalculVfh = CalculVfh + CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]") * GetN2(Ngaz_Air, Airnames(), Aircompo()) / 100
''Etape 6 : calcul du volume de CO2 provenant du comburant
'CalculVfh = CalculVfh + CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]") * GetCO2(Ngaz_Air, Airnames(), Aircompo()) / 100
''Etape 7 : calcul du volume de H2O provenant du comburant
'CalculVfh = CalculVfh + CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]") * GetH2O(Ngaz_Air, Airnames(), Aircompo()) / 100
''Etape 8 : calcul du volume de SO2 provenant du comburant
'CalculVfh = CalculVfh + CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]") * GetSO2(Ngaz_Air, Airnames(), Aircompo()) / 100
'
'
'
'If unite = "[Nm3/kg]" Then
'    CalculVfh = CalculVfh / CalculMasseVolumique(Ngaz_fuel, Fuelnames, Fuelcompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[scf/lb]" Then
'    CalculVfh = CalculVfh / CalculMasseVolumique(Ngaz_fuel, Fuelnames, Fuelcompo, 0, 101325, "[kg/m3]")
'
'ElseIf unite = "[kg/Nm3]" Then
'    Fumnames(1) = "CO2"
'    Fumnames(2) = "H2O"
'    Fumnames(3) = "O2"
'    Fumnames(4) = "N2"
'    Fumnames(5) = "SO2"
'    Fumcompo(1) = CalculFracVolCO2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(2) = CalculFracVolH2O(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(3) = CalculFracVolO2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(4) = CalculFracVolN2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(5) = CalculFracVolSO2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'
'    CalculVfh = CalculVfh * CalculMasseVolumique(5, Fumnames, Fumcompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[lb/scf]" Then
'    Fumnames(1) = "CO2"
'    Fumnames(2) = "H2O"
'    Fumnames(3) = "O2"
'    Fumnames(4) = "N2"
'    Fumnames(5) = "SO2"
'
'    Fumcompo(1) = CalculFracVolCO2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(2) = CalculFracVolH2O(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(3) = CalculFracVolO2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(4) = CalculFracVolN2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(5) = CalculFracVolSO2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'
'    CalculVfh = CalculVfh * CalculMasseVolumique(5, Fumnames, Fumcompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[kg/kg]" Or unite = "[lb/lb]" Then
'    Fumnames(1) = "CO2"
'    Fumnames(2) = "H2O"
'    Fumnames(3) = "O2"
'    Fumnames(4) = "N2"
'    Fumnames(5) = "SO2"
'
'    Fumcompo(1) = CalculFracVolCO2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(2) = CalculFracVolH2O(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(3) = CalculFracVolO2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(4) = CalculFracVolN2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'    Fumcompo(5) = CalculFracVolSO2(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, 1)
'
'
'    CalculVfh = CalculVfh * CalculMasseVolumique(5, Fumnames, Fumcompo, 0, 101325, "[kg/m3]") _
'                / CalculMasseVolumique(Ngaz_fuel, Fuelnames, Fuelcompo, 0, 101325, "[kg/m3]")
'End If
'
'End Function
'
'Function CalculVfs(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, unite As String)
'
''Calcul du pouvoir fumigène sec en Nm3 de fumées /Nm3 de combustible
'
'CalculVfs = CalculVfh(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "Nm3/Nm3")
'
'For i = 1 To Ngaz_fuel
'    CalculVfs = CalculVfs - Fuelcompo(i) / 100 * GetVH2Ogaz(Fuelnames(i), SProp)
'Next i
'
'CalculVfs = CalculVfs - CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "Nm3/Nm3") * GetH2O(Ngaz_Air, Airnames(), Aircompo()) / 100
'
'If unite = "[Nm3/kg]" Then
'    CalculVfs = CalculVfs / CalculMasseVolumique(Ngaz_fuel, Fuelnames, Fuelcompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[scf/lb]" Then
'    CalculVfs = CalculVfs / CalculMasseVolumique(Ngaz_fuel, Fuelnames, Fuelcompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[kg/Nm3]" Then
'    CalculVfs = CalculVfs * CalculMasseVolumique(Ngaz_Air, Airnames, Aircompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[lb/scf]" Then
'    CalculVfs = CalculVfs * CalculMasseVolumique(Ngaz_Air, Airnames, Aircompo, 0, 101325, "[kg/m3]")
'ElseIf unite = "[kg/kg]" Or unite = "[lb/lb]" Then
'    CalculVfs = CalculVfs * CalculMasseVolumique(Ngaz_Air, Airnames, Aircompo, 0, 101325, "[kg/m3]") _
'                / CalculMasseVolumique(Ngaz_fuel, Fuelnames, Fuelcompo, 0, 101325, "[kg/m3]")
'End If
'
'End Function
'
'Function GetVO2gaz(Nomgaz As String, PropSheet As String)
'
''PropSheet est le nom de la feuille des propriétés
'Dim i As Integer
'GetVO2gaz = -100
'i = 0
'
'While GetVO2gaz = -100 And i < 25
'
'    If Worksheets(PropSheet).Range("B3").Offset(i) = Nomgaz Then
'        GetVO2gaz = 0.5 * (2 * Worksheets(PropSheet).Range("B3").Offset(i, 1) + 0.5 * Worksheets(PropSheet).Range("B3").Offset(i, 2) - Worksheets(PropSheet).Range("B3").Offset(i, 3) + 2 * Worksheets(PropSheet).Range("B3").Offset(i, 5))
'    Else
'        i = i + 1
'    End If
'Wend
'
'If GetVO2gaz = -100 Then
'    Call MsgBox("Le gaz pur : " & Nomgaz & " n'existe pas dans la base.", vbCritical)
'End If
'
'End Function
'
'Function GetVCO2gaz(Nomgaz As String, PropSheet As String)
'
''PropSheet est le nom de la feuille des propriétés
'
'Dim i As Integer
'GetVCO2gaz = -100
'i = 0
'
'While GetVCO2gaz = -100 And i < 25
'
'    If Worksheets(PropSheet).Range("B3").Offset(i) = Nomgaz Then
'        GetVCO2gaz = Worksheets(PropSheet).Range("B3").Offset(i, 1)
'    Else
'        i = i + 1
'    End If
'Wend
'
'If GetVCO2gaz = -100 Then
'    Call MsgBox("Le gaz pur : " & Nomgaz & " n'existe pas dans la base.", vbCritical)
'End If
'
'End Function
'
'Function GetVH2Ogaz(Nomgaz As String, PropSheet As String)
'
''PropSheet est le nom de la feuille des propriétés
'Dim i As Integer
'GetVH2Ogaz = -100
'i = 0
'
'While GetVH2Ogaz = -100 And i < 25
'
'    If Worksheets(PropSheet).Range("B3").Offset(i) = Nomgaz Then
'        GetVH2Ogaz = 0.5 * Worksheets(PropSheet).Range("B3").Offset(i, 2)
'    Else
'        i = i + 1
'    End If
'Wend
'
'If GetVH2Ogaz = -100 Then
'    Call MsgBox("Le gaz pur : " & Nomgaz & " n'existe pas dans la base.", vbCritical)
'End If
'
'End Function
'
'Function GetVSO2gaz(Nomgaz As String, PropSheet As String)
'
''PropSheet est le nom de la feuille des propriétés
'Dim i As Integer
'GetVSO2gaz = -100
'i = 0
'
'While GetVSO2gaz = -100 And i < 25
'
'    If Worksheets(PropSheet).Range("B3").Offset(i) = Nomgaz Then
'        GetVSO2gaz = Worksheets(PropSheet).Range("B3").Offset(i, 5)
'    Else
'        i = i + 1
'    End If
'Wend
'
'If GetVSO2gaz = -100 Then
'    Call MsgBox("Le gaz pur : " & Nomgaz & " n'existe pas dans la base.", vbCritical)
'End If
'
'End Function
'
'Function GetVN2gaz(Nomgaz As String, PropSheet As String)
'
''PropSheet est le nom de la feuille des propriétés
'Dim i As Integer
'GetVN2gaz = -100
'i = 0
'
'While GetVN2gaz = -100 And i < 25
'
'    If Worksheets(PropSheet).Range("B3").Offset(i) = Nomgaz Then
'        GetVN2gaz = 0.5 * Worksheets(PropSheet).Range("B3").Offset(i, 4)
'    Else
'        i = i + 1
'    End If
'Wend
'
'If GetVN2gaz = -100 Then
'    Call MsgBox("Le gaz pur : " & Nomgaz & " n'existe pas dans la base.", vbCritical)
'End If
'
'
'End Function
'
'Function GetO2(Ngaz As Integer, CombNames() As String, CombCompo() As Double)
'
'GetO2 = 0
'
'' Recherche du % volumique d'O2 dans le comburant
'Dim j As Integer
'
'j = 0
'For i = 1 To Ngaz
'    If CombNames(i) = "O2" Then
'        j = j + 1
'        GetO2 = GetO2 + CombCompo(i)
'    End If
'Next i
'
'
'End Function
'
'Function GetCO2(Ngaz As Integer, CombNames() As String, CombCompo() As Double)
'
'GetCO2 = 0
'
'' Recherche du % volumique de CO2 dans le comburant
'Dim j As Integer
'
'j = 0
'For i = 1 To Ngaz
'    If CombNames(i) = "CO2" Then
'        j = j + 1
'        GetCO2 = GetCO2 + CombCompo(i)
'    End If
'Next i
'
'
'End Function
'
'Function GetH2O(Ngaz As Integer, CombNames() As String, CombCompo() As Double)
'
'GetH2O = 0
'
'' Recherche du % volumique d'H2O dans le comburant
'Dim j As Integer
'
'j = 0
'For i = 1 To Ngaz
'    If CombNames(i) = "H2O" Then
'        j = j + 1
'        GetH2O = GetH2O + CombCompo(i)
'    End If
'Next i
'
'End Function
'
'Function GetN2(Ngaz As Integer, CombNames() As String, CombCompo() As Double)
'
'GetN2 = 0
'
'' Recherche du % volumique de N2 dans le comburant
'Dim j As Integer
'
'j = 0
'For i = 1 To Ngaz
'    If CombNames(i) = "N2" Then
'        j = j + 1
'        GetN2 = GetN2 + CombCompo(i)
'    End If
'Next i
'
'
'End Function
'
'Function GetSO2(Ngaz As Integer, CombNames() As String, CombCompo() As Double)
'
'GetSO2 = 0
'
'' Recherche du % volumique de SO2 dans le comburant
'Dim j As Integer
'
'j = 0
'For i = 1 To Ngaz
'    If CombNames(i) = "SO2" Then
'        j = j + 1
'        GetSO2 = GetSO2 + CombCompo(i)
'    End If
'Next i
'
'End Function
'
'
'Function CalculFracVolCO2(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, AirFuelRatio As Double)
'
''Calcul de la fraction volumique de CO2 dans les fumées issue de la combustion
'
'    CalculFracVolCO2 = 0
'
'    'Calcul de la fraction volumique de CO2 présente dans les fumées
'
'    'Etape 1 : calcul du volume de CO2 généré par la combustion
'    For i = 1 To Ngaz_fuel
'        CalculFracVolCO2 = CalculFracVolCO2 + Fuelcompo(i) / 100 * GetVCO2gaz(Fuelnames(i), SProp)
'    Next i
'
'    'Etape 2 : calcul du volume de CO2 provenant du comburant
'    CalculFracVolCO2 = CalculFracVolCO2 + AirFuelRatio * GetCO2(Ngaz_Air, Airnames(), Aircompo()) / 100 * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]")
'
'    'Etape 3 : Division par le volume de fumées total
'    CalculFracVolCO2 = 100 * CalculFracVolCO2 / (CalculVfh(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]"))
'
'End Function
'
'Function CalculFracVolH2O(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, AirFuelRatio As Double)
'
''Calcul de la fraction volumique de H2O dans les fumées issue de la combustion
'
'    CalculFracVolH2O = 0
'    'Calcul de la fraction volumique de H2O présente dans les fumées
'
'
'    'Etape 1 : calcul du volume de H2O généré par la combustion
'    For i = 1 To Ngaz_fuel
'        CalculFracVolH2O = CalculFracVolH2O + Fuelcompo(i) / 100 * GetVH2Ogaz(Fuelnames(i), SProp)
'    Next i
'
'    'Etape 2 : calcul du volume de H2O provenant du comburant
'    CalculFracVolH2O = CalculFracVolH2O + AirFuelRatio * GetH2O(Ngaz_Air, Airnames(), Aircompo()) / 100 * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]")
'
'    'Etape2 : Division par le volume de fumées total
'    CalculFracVolH2O = 100 * CalculFracVolH2O / (CalculVfh(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]"))
'
'End Function
'
'Function CalculFracVolN2(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, AirFuelRatio As Double)
'
''Calcul de la fraction volumique de O2 dans les fumées issue de la combustion
'
'
'    CalculFracVolN2 = 0
'    'Calcul de la fraction volumique de N2 présente dans les fumées
'
'    'Etape 1 : calcul du volume de N2 généré par la combustion
'    For i = 1 To Ngaz_fuel
'        CalculFracVolN2 = CalculFracVolN2 + Fuelcompo(i) / 100 * GetVN2gaz(Fuelnames(i), SProp)
'    Next i
'
'    'Etape 2 : calcul du volume de N2 provenant du comburant
'    CalculFracVolN2 = CalculFracVolN2 + AirFuelRatio * GetN2(Ngaz_Air, Airnames(), Aircompo()) / 100 * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]")
'
'    'Etape3 : Division par le volume de fumées total
'    CalculFracVolN2 = 100 * CalculFracVolN2 / (CalculVfh(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]"))
'
'End Function
'
'Function CalculFracVolO2(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, AirFuelRatio As Double)
'
'    'Calcul de la fraction volumique de O2 présente dans les fumées
'
'    CalculFracVolO2 = 0
'
'    'Etape 1 : calcul du volume de O2 provenant de l'excès d'air
'    CalculFracVolO2 = CalculFracVolO2 + (AirFuelRatio - 1) * GetO2(Ngaz_Air, Airnames(), Aircompo()) / 100 * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]")
'
'    'Etape 2 : Division par le volume de fumées total
'    CalculFracVolO2 = 100 * CalculFracVolO2 / (CalculVfh(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]"))
'
'End Function
'
'Function CalculFracVolSO2(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, AirFuelRatio As Double)
''Calcul de la fraction volumique de SO2 présente dans les fumées
'
'    CalculFracVolSO2 = 0
'
'    'Etape 1 : calcul du volume de SO2 généré par la combustion
'    For i = 1 To Ngaz_fuel
'        CalculFracVolSO2 = CalculFracVolSO2 + Fuelcompo(i) / 100 * GetVSO2gaz(Fuelnames(i), SProp)
'    Next i
'
'    'Etape 2 : calcul du volume de SO2 provenant du comburant
'    CalculFracVolSO2 = CalculFracVolSO2 + AirFuelRatio * GetSO2(Ngaz_Air, Airnames(), Aircompo()) / 100 * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]")
'
'    'Etape2 : Division par le volume de fumées total
'    CalculFracVolSO2 = 100 * CalculFracVolSO2 / (CalculVfh(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[Nm3/Nm3]"))
'
'
'End Function
'
'
'Function CalculTadiab(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, Ngaz_WG As Integer, WGNames() As String, WGCompo() As Double, Tfuel As Double, Tair As Double, AirFuelRatio As Double, unite As String)
'
'    'Calcul de la température adiabatique [°C] de la combustion en fonction des conditions de préchauffage et d'excès d'air
'    ' Les températures sont en [°Celsius]
'    ' Les composition sont en [%vol.]
'
'    Dim Pfuel, Pair, Hfumees, ValInf, ValSup As Double
'    Dim PCI As Double
'    Dim DebitFumees As Double
'
'    PCI = CalculPCI(Ngaz_fuel, Fuelnames, Fuelcompo, "[kJ/kg]")
'
'    '-----Hypothèse débit de combustible 1 kg/s------
'
'    'Etape 1 : calcul de la puissance apportée par le combustible préchauffé [W]
'    Pfuel = 1 * CalculEnthalp(Ngaz_fuel, Fuelnames, Fuelcompo, Tfuel, 0, "[J/kg]")
'
'    'Etape 2 : calcul de la puissance apportée par le comburant préchauffé [W]
'    Pair = AirFuelRatio * CalculVa(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo(), "[kg/kg]") * CalculEnthalp(Ngaz_Air, Airnames, Aircompo, Tair, 0, "[J/kg]")
'
'    'Etape 3 : calcul de l'enthalpie des fumées après combustion adiabatique
'
'    DebitFumees = CalculVfh(Ngaz_fuel, Fuelnames, Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), "[kg/kg]") + (AirFuelRatio - 1) * CalculVa(Ngaz_fuel, Fuelnames, Fuelcompo, Ngaz_Air, Airnames, Aircompo, "[kg/kg]")
'
'    Hfumees = (Pfuel + Pair + PCI * 1000) / DebitFumees
'
'    'Etape 4 : calcul de la température adiabatique des fumées à partir de son enthalpie par dichotomie
'    ValInf = 0
'    ValSup = 5000
'
'    While Abs(ValInf - ValSup) > 0.5 ' Condition d'arret = solution trouvée à 0.5°C près
'        temp = CalculEnthalp(Ngaz_WG, WGNames, WGCompo, 0.5 * (ValInf + ValSup), 0, "[J/kg]")
'        If CalculEnthalp(Ngaz_WG, WGNames, WGCompo, 0.5 * (ValInf + ValSup), 0, "[J/kg]") > Hfumees Then
'            ValSup = 0.5 * (ValInf + ValSup)
'            Else: ValInf = 0.5 * (ValInf + ValSup)
'        End If
'    Wend
'
'    CalculTadiab = ValInf
'
'    If unite = "[°K]" Then CalculTadiab = CalculTadiab + 273
'    If unite = "[°F]" Then CalculTadiab = CalculTadiab * 1.8 + 32
'    If unite = "[°R]" Then CalculTadiab = CalculTadiab * 1.8 + 32 + 459.67
'
'End Function
'
'
'Function CalculEnthalp(Ngaz As Integer, CombNames() As String, CombCompo() As Double, Tcelsius As Double, Tref As Double, unite As String)
''Calcul de l'enthalpie d'un composé constitué de Ngaz composants à la température Tcelsius, Tref étant la température de référence
'' Enthalpie en [J/kg]
''Ngaz : nombre de gaz contenus dans le composé
''CombNames() : nom de gaz élémentaires contenus dans le composé - l'index zéro n'est pas utilisé
''CombCompo() : pourcentage volumique [%] de chacun des gaz - l'index zéro n'est pas utilisé
'Dim i As Integer
'CalculEnthalp = 0
'
'For i = 1 To Ngaz
'    CalculEnthalp = CalculEnthalp + Gethgaz(CombNames(i), Tref, Tcelsius, SProp) * GetFracMass(Ngaz, CombNames(), CombCompo(), i)
'Next i
'
'If unite = "[kJ/kg]" Then CalculEnthalp = CalculEnthalp / 1000
'If unite = "[kcal/kg]" Then CalculEnthalp = CalculEnthalp / 1000 * 0.2388459
'If unite = "[BTU/lb]" Then CalculEnthalp = CalculEnthalp * 0.0009478171 / 2.20462
'
'
'End Function
'
'Function CalculTemperature(Ngaz As Integer, CombNames() As String, CombCompo() As Double, Puissance As Double, Debit As Double)
'
'    '---------------------------------------------------------------
'    ' Calcul de la température d'un fluide connaissant :
'    '  - La composition en %vol.
'    '  - La puissance apportée en W
'    '  - Le débit en Nm3/h
'    '---------------------------------------------------------------
'    ' La température renvoyée est en Celsius, calcul par dichotomie
'    '---------------------------------------------------------------
'    ' L. FERRANDanvier 2009
'    '---------------------------------------------------------------
'
'    Dim Enthalpie As Double ' en J/kg
'
'    Dim Hmin As Double 'en J/kg
'    Dim Hmax As Double ' en J/kg
'    Dim Tmin As Double ' en Celsius
'    Dim Tmax As Double ' en Celsius
'    Dim i As Integer
'
'    '-----Calcul de l'enthalpie du fluide
'
'    Enthalpie = Puissance / (Debit / 3600 * CalculMasseVolumique(Ngaz, CombNames, CombCompo, 0, 101325, "[kg/Nm3]"))
'
'    Hmin = 0
'    Hmax = 4000000
'    Tmin = 0
'    Tmax = 2000
'    i = 0
'
'    While Abs(Enthalpie - Hmin) > 0.001 * Enthalpie And i < 50
'        Hmin = CalculEnthalp(Ngaz, CombNames, CombCompo, Tmin, 0, "[J/kg]")
'        Hmax = CalculEnthalp(Ngaz, CombNames, CombCompo, Tmax, 0, "[J/kg]")
'
'        If Enthalpie > 0.5 * (Hmin + Hmax) Then
'            Tmin = 0.75 * Tmin + 0.25 * Tmax
'        Else
'            Tmax = 0.25 * Tmin + 0.75 * Tmax
'        End If
'    Wend
'
'    If i < 50 Then
'        CalculTemperature = 0.5 * (Tmin + Tmax)
'    Else
'        CalculTemperature = -1
'        Call MsgBox("Problème de convergence de l'enthalpie", vbAbortRetryIgnore)
'    End If
'
'End Function
'
'Function Gethgaz(Nomgaz As String, Tref As Double, Tcelsius As Double, PropSheet As String)
'
'' Recherche du gaz dans la liste et calcul de son enthalpie à la température Tcelsius, Tref étant la température de référence
'' Enthalpie en [J/kg]
'Dim i As Integer
'Gethgaz = -1
'i = 0
'While Gethgaz = -1
'    If Worksheets(PropSheet).Range("B3").Offset(i) = Nomgaz Then
'            Gethgaz = 0
'            For j = 1 To 7
'                Gethgaz = Gethgaz + Worksheets(PropSheet).Range("B3").Offset(i, 10 + j) / j * ((Tcelsius + 273) ^ j - (Tref + 273) ^ j)
'            Next j
'    Else
'        i = i + 1
'    End If
'Wend
'
'
'End Function
'
'
'Function CalculDebitFuel(NgazFuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Puissance As Double, unite As String)
'
'' La puissance est en kW
'
'Select Case unite
'    Case "[Nm3/h]"
'        CalculDebitFuel = Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/Nm3]") * 3600
'    Case "[kg/h]"
'        CalculDebitFuel = Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/kg]") * 3600
'    Case "[kg/s]"
'        CalculDebitFuel = Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/kg]")
'    Case "[l/h]"
'        CalculDebitFuel = 0.001 * Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/Nm3]") * 3600
'    Case "[scf/h]"
'        CalculDebitFuel = 35.31467 * (273 + 15) / 273 * Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/Nm3]") * 3600
'    Case Else
'       CalculDebitFuel = 0
'End Select
'
'End Function
'
'Function CalculDebitAir(NgazFuel As Integer, Fuelnames() As String, Fuelcompo() As Double, NgazAir As Integer, Airnames() As String, Aircompo() As Double, Puissance As Double, RatioAirGaz As Double, unite As String)
'
'' La puissance est en kW
'
'Select Case unite
'    Case "[Nm3/h]"
'        CalculDebitAir = Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/Nm3]") * RatioAirGaz * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[Nm3/Nm3]") * 3600
'    Case "[kg/h]"
'        CalculDebitAir = Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/kg]") * RatioAirGaz * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[kg/kg]") * 3600
'    Case "[kg/s]"
'        CalculDebitAir = Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/kg]") * RatioAirGaz * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[kg/kg]")
'    Case "[l/h]"
'        CalculDebitAir = 0.001 * Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/Nm3]") * RatioAirGaz * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[Nm3/Nm3]") * 3600
'    Case "[scf/h]"
'        CalculDebitAir = 35.31467 * (273 + 15) / 273 * Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/Nm3]") * RatioAirGaz * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[Nm3/Nm3]") * 3600
'    Case Else
'        CalculDebitAir = -1
'End Select
'
'End Function
'
'Public Sub WasteGasCompo(NgazFuel As Integer, Fuelnames() As String, Fuelcompo() As Double, _
'                NgazAir As Integer, Airnames() As String, Aircompo() As Double, AirFuelRatio As Double, _
'                Fumnames() As String, ByRef Fumcompo() As Double)
'
'    '------------------------------------------------
'    ' Calculation of Waste gases composition in
'    ' sub-stoechiometric conditions
'    ' at the Temperature Tequilibrium
'    '------------------------------------------------
'    ' P.DUBOIS 2026
'    '------------------------------------------------
'    '
'    '------------------------------------------------
'    '
'    ' Compositions [%vol.]
'    ' Tableau FumNames retourné :
'    '    FumNames(1) = "CO2"
'    '    FumNames(2) = "H2O"
'    '    FumNames(3) = "O2"
'    '    FumNames(4) = "N2"
'    '    FumNames(5) = "SO2"
'    '    FumNames(6) = "CO"
'    '    FumNames(7) = "H2"
'    '------------------------------------------------
'
'
'    Dim KTfum As Double
'    Dim KTadiab As Double
'
'    Dim BB As Double
'    Dim DD As Double
'
'    ReDim Fumnames(5)
'     ReDim Fumcompo(5)
'    '---- Composition des fumées
'    Fumnames(1) = "CO2"
'    Fumnames(2) = "H2O"
'    Fumnames(3) = "O2"
'    Fumnames(4) = "N2"
'    Fumnames(5) = "SO2"
'
'    Fumcompo(1) = CalculFracVolCO2(NgazFuel, Fuelnames(), Fuelcompo(), NgazAir, Airnames(), Aircompo(), 1 + 0.01 * AirFuelRatio)
'    Fumcompo(2) = CalculFracVolH2O(NgazFuel, Fuelnames(), Fuelcompo(), NgazAir, Airnames(), Aircompo(), 1 + 0.01 * AirFuelRatio)
'    Fumcompo(3) = CalculFracVolO2(NgazFuel, Fuelnames(), Fuelcompo(), NgazAir, Airnames(), Aircompo(), 1 + 0.01 * AirFuelRatio)
'    Fumcompo(4) = CalculFracVolN2(NgazFuel, Fuelnames(), Fuelcompo(), NgazAir, Airnames(), Aircompo(), 1 + 0.01 * AirFuelRatio)
'    Fumcompo(5) = CalculFracVolSO2(NgazFuel, Fuelnames(), Fuelcompo(), NgazAir, Airnames(), Aircompo(), 1 + 0.01 * AirFuelRatio)
'End Sub
'
'Function CalculDebitFumees(NgazFuel As Integer, Fuelnames() As String, Fuelcompo() As Double, NgazAir As Integer, Airnames() As String, Aircompo() As Double, Puissance As Double, RatioAirGaz As Double, unite As String)
'
'' La puissance est en kW
'
'
'Select Case unite
'    Case "[Nm3/h]"
'        CalculDebitFumees = Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/Nm3]") _
'                * (CalculVfh(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[Nm3/Nm3]") _
'                + (RatioAirGaz - 1) * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[Nm3/Nm3]")) * 3600
'    Case "[kg/h]"
'        CalculDebitFumees = Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/kg]") _
'                * (CalculVfh(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[kg/kg]") _
'                + (RatioAirGaz - 1) * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[kg/kg]")) * 3600
'    Case "[kg/s]"
'        CalculDebitFumees = Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/kg]") _
'                * (CalculVfh(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[kg/kg]") _
'                + (RatioAirGaz - 1) * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[kg/kg]"))
'    Case "[l/h]"
'        CalculDebitFumees = 0.001 * Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/Nm3]") _
'                * (CalculVfh(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[Nm3/Nm3]") _
'                + (RatioAirGaz - 1) * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[Nm3/Nm3]")) * 3600
'    Case "[scf/h]"
'        CalculDebitFumees = 35.31467 * (273 + 15) / 273 * Puissance / CalculPCI(NgazFuel, Fuelnames, Fuelcompo, "[kJ/Nm3]") _
'                * (CalculVfh(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[Nm3/Nm3]") _
'                + (RatioAirGaz - 1) * CalculVa(NgazFuel, Fuelnames, Fuelcompo, NgazAir, Airnames, Aircompo, "[Nm3/Nm3]")) * 3600
'    Case Else
'        CalculDebitFumees = -1
'End Select
'
'End Function
'
'Function CalculTeqAB(Ngaz_A As Integer, ANames() As String, ACompo() As Double, AMassFlowrate As Double, T_A As Double, _
'    Ngaz_B As Integer, BNames() As String, BCompo() As Double, BMassFlowrate As Double)
'
''Calcul de la température équivalente d'un gaz B [°C] en fonction des conditions d'un gaz A
'' Les températures sont en °Celsius, les débits en kg/s
'
'Dim P_A, Hfumees, ValInf, ValSup As Double
'
''Etape 1 : calcul de la puissance apportée par le gaz A
'P_A = AMassFlowrate * CalculEnthalp(Ngaz_A, ANames(), ACompo(), T_A, 0, "[J/kg]")
'
''Etape 2 : calcul de l'enthalpie du gaz B à la puissance du gaz A
'H_B = P_A / BMassFlowrate
'
''Etape 3 : calcul de la température des fumées à partir de son enthalpie par dichotomie
'' Calcul de la composition des fumées
'
'ValInf = 0
'ValSup = 5000
'
'While Abs(ValInf - ValSup) > 0.5 ' Condition d'arret = solution trouvée à 0.5°C près
'
'    If CalculEnthalp(Ngaz_B, BNames(), BCompo(), 0.5 * (ValInf + ValSup), 0, "[J/kg]") > H_B Then
'        ValSup = 0.5 * (ValInf + ValSup)
'        Else: ValInf = 0.5 * (ValInf + ValSup)
'    End If
'Wend
'
'CalculTeqAB = ValInf
'
'End Function
'Function CalculTeq(Ngaz_fuel As Integer, Fuelnames() As String, Fuelcompo() As Double, Ngaz_Air As Integer, Airnames() As String, Aircompo() As Double, Tfuel As Double, Tair As Double, AirFuelRatio As Double, Puissance As Double)
'
''Calcul de la température équivalente des fumées [°C] en fonction des conditions de préchauffage et d'excès d'air
'' Les températures sont en °Celsius, la puissance en kW
'' AirFuelRatio est le ratio air/gaz
'' Tfuel en [°C]
'' Tair en [°C]
'
'Dim Pfuel, Pair, Hfumees, ValInf, ValSup As Double
'
''Etape 1 : calcul de la puissance apportée par le combustible préchauffé
'Pfuel = CalculDebitFuel(Ngaz_fuel, Fuelnames(), Fuelcompo(), Puissance, "[kg/s]") * CalculEnthalp(Ngaz_fuel, Fuelnames(), Fuelcompo(), Tfuel, 0, "[J/kg]")
'
''Etape 2 : calcul de la puissance apportée par le comburant préchauffé
'Pair = CalculDebitAir(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), Puissance, AirFuelRatio, "[kg/s]") * CalculEnthalp(Ngaz_Air, Airnames(), Aircompo(), Tair, 0, "[J/kg]")
'
''Etape 3 : calcul de l'enthalpie des fumées après combustion adiabatique
'Hfumees = (Pfuel + Pair) / CalculDebitFumees(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), Puissance, AirFuelRatio, "[kg/s]")
'
'
''Etape 4 : calcul de la température des fumées à partir de son enthalpie par dichotomie
'' Calcul de la composition des fumées
'Dim Ngaz_fum As Integer
'Ngaz_fum = 5
'Dim Fumnames() As String
'Dim Fumcompo() As Double
'
'ReDim Fumnames(Ngaz_fum)
'ReDim Fumcompo(Ngaz_fum)
'
'Fumnames(1) = "CO2"
'Fumnames(2) = "H2O"
'Fumnames(3) = "O2"
'Fumnames(4) = "N2"
'Fumnames(5) = "SO2"
'
'Fumcompo(1) = CalculFracVolCO2(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), AirFuelRatio)
'Fumcompo(2) = CalculFracVolH2O(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), AirFuelRatio)
'Fumcompo(3) = CalculFracVolO2(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), AirFuelRatio)
'Fumcompo(4) = CalculFracVolN2(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), AirFuelRatio)
'Fumcompo(5) = CalculFracVolSO2(Ngaz_fuel, Fuelnames(), Fuelcompo(), Ngaz_Air, Airnames(), Aircompo(), AirFuelRatio)
'
'ValInf = 0
'ValSup = 5000
'
'While Abs(ValInf - ValSup) > 0.5 ' Condition d'arret = solution trouvée à 0.5°C près
'
'    If CalculEnthalp(Ngaz_fum, Fumnames(), Fumcompo(), 0.5 * (ValInf + ValSup), 0, "[J/kg]") > Hfumees Then
'        ValSup = 0.5 * (ValInf + ValSup)
'        Else: ValInf = 0.5 * (ValInf + ValSup)
'    End If
'Wend
'
'CalculTeq = ValInf
'
'End Function
'
'Public Sub CalculEchangeur(Ngaz_chaud As Integer, ChaudNames() As String, ChaudCompo() As Double, DebitGazChaud As Double, TGazChaud_in As Double, Ngaz_froid As Integer, FroidNames() As String, FroidCompo() As Double, DebitGazFroid As Double, TGazFroid_in As Double, Efficacite As Double, Rendement As Double, TGazChaud_out As Double, TGazFroid_out As Double)
'
''----------------------------------------------------------------------------------------------------------------
''  Objet : calcul des températures des fluides aux bornes d'un échangeur (système régénératif, à tubes, etc...
''          Entrées :
''          -> Ngaz_chaud, ChaudNames, ChaudCompo : composé gazeux cédant son enthalpie (typiquement les fumées)
''          -> DebitGazChaud [kg/s] : débit du gaz initialement chaud
''          -> TGazChaud_in [K]     : température d'entrée du fluide initialement chaud
''          -> Ngaz_froid, FroidNames, FroidCompo : composé gazeux récupérant l'enthalpie (combustible ou comburant)
''          -> DebitGazFroid [kg/s] : débit du gaz initialement froid
''          -> TGazFroid_in [K]     : température d'entrée du fluide initialement froid
''          -> Efficacite [%]       : efficacité de l'échangeur
''          -> Rendement [%]        : rendement de l'échangeur
''
''          Sorties :
''          -> TGazChaud_out [K]    : température du fluide chaud après avoir cédé son enthalpie
''          -> TGazFroid_out [K]    : température du fluide froid après avoir récupéré l'enthalpie
''-----------------------------------------------------------------------------------------------------------------
''  Méthode : -le calcul est itératif, TGazChaud_out puis TGazFroid_out sont calculés par dichotomie,
''
''-----------------------------------------------------------------------------------------------------------------
''  L.Ferrand4/02/2006
''-----------------------------------------------------------------------------------------------------------------
'Dim i As Integer
'Dim Tinf As Double
'Dim Tsup As Double
'Dim P1 As Double
'Dim P2 As Double
'Dim Eff As Double
'Dim Pchaud As Double
'Dim Pchaudmax As Double
'Dim Pfroidmax As Double
'
''Initialisations
'Tinf = 263
'Tsup = 3273
'TGazChaud_out = Tsup
'
'' Calcul de la température de sortie du gaz initialement chaud
'i = 0
'Eff = 0
'
'While (Sqr((Efficacite / 100 - Eff) * (Efficacite / 100 - Eff)) > 0.01 And i < 100)
'    i = i + 1
'
'    If (Eff < (Efficacite / 100)) Then Tsup = TGazChaud_out
'    If (Eff >= (Efficacite / 100)) Then Tinf = TGazChaud_out
'
'    TGazChaud_out = (Tinf + Tsup) / 2
'
'    Pchaud = DebitGazChaud * CalculEnthalp(Ngaz_chaud, ChaudNames, ChaudCompo, TGazChaud_in - 273, TGazChaud_out - 273, "[J/kg]")
'    Pchaudmax = DebitGazChaud * CalculEnthalp(Ngaz_chaud, ChaudNames, ChaudCompo, TGazChaud_in - 273, TGazFroid_in - 273, "[J/kg]")
'    Pfroidmax = DebitGazFroid * CalculEnthalp(Ngaz_froid, FroidNames, FroidCompo, TGazChaud_in - 273, TGazFroid_in - 273, "[J/kg]")
'
'    Eff = Pchaud / Min(Pchaudmax, Pfroidmax)
'Wend
'
'' Calcul de la température de sortie du gaz initialement froid
'P1 = 1000
'P2 = 0
'Tinf = 263
'Tsup = 3000
'
'TGazFroid_out = Tinf
'
'i = 0
'
'While ((Tsup - Tinf) > 2 And i < 100)
'    i = i + 1
'
'    If (P2 <= P1) Then Tinf = TGazFroid_out
'    If (P2 > P1) Then Tsup = TGazFroid_out
'
'    TGazFroid_out = (Tinf + Tsup) / 2
'
'    P1 = Rendement / 100 * DebitGazChaud * CalculEnthalp(Ngaz_chaud, ChaudNames, ChaudCompo, TGazChaud_in - 273, TGazChaud_out - 273, "[J/kg]")
'    P2 = DebitGazFroid * CalculEnthalp(Ngaz_froid, FroidNames, FroidCompo, TGazFroid_out - 273, TGazFroid_in - 273, "[J/kg]")
'
'Wend
'
'End Sub
'
'
'Function CalculCp(Ngaz As Integer, CombNames() As String, CombCompo() As Double, Tcelsius As Double)
''Calcul du Cp d'un composé constitué de Ngaz composants à la température Tcelsius
'
''Ngaz : nombre de gaz contenus dans le composé
''CombNames() : nom de gaz élémentaires contenus dans le composé - l'index zéro n'est pas utilisé
''CombCompo() : pourcentage volumique [%] de chacun des gaz - l'index zéro n'est pas utilisé
'Dim i As Integer
'
'CalculCp = 0
'
'For i = 1 To Ngaz
'    CalculCp = CalculCp + GetCpgaz(CombNames(i), Tcelsius) * GetFracMass(Ngaz, CombNames(), CombCompo(), i)
'Next i
'
'End Function
'
'Function GetCpgaz(Nomgaz As String, Tcelsius As Double)
'
'Dim i As Integer
'
'' Recherche du gaz dans la liste et calcul de sa chaleur massique à la température Tcelsius
'GetCpgaz = -1
'i = 0
'While GetCpgaz = -1 And i <= NMAX_FUEl
'    i = i + 1
'    If Worksheets(SProp).Range("B3").Offset(i) = Nomgaz Then
'        GetCpgaz = 0
'        For j = 1 To 7
'            GetCpgaz = GetCpgaz + Worksheets(SProp).Range("B3").Offset(i, 10 + j) * (Tcelsius + 273) ^ (j - 1)
'        Next j
'    End If
'Wend
'
'End Function
'
'Public Sub CalculDilution(Ngaz_fum As Integer, Fumnames() As String, Fumcompo() As Double, QFum As Double, TFumCelsius As Double, _
'            TProtecCelsius As Double, TAirDilution As Double, QAirDilution As Double)
''-------------------------------------------------------------------------------------------------
'' Calcul du débit d'air de dilution pour respecter la température de protection du récupérateur
''-------------------------------------------------------------------------------------------------
'' L. Ferrand6/02/2005
''-------------------------------------------------------------------------------------------------
'
'Dim FumRo As Double ' Masse volumique des fumées
'Dim AirRo As Double
'Dim mFum As Double ' Débit de fumées [kg/s]
'Dim mAir As Double  'Debit d'air de dilution [kg/s]
'
'Dim Airnames(2) As String
'Dim Aircompo(2) As Double
'
'Airnames(1) = "O2"
'Airnames(2) = "N2"
'Aircompo(1) = 20.8
'Aircompo(2) = 79.2
'
'FumRo = CalculMasseVolumique(Ngaz_fum, Fumnames, Fumcompo, 0, 101325, "[kg/m3]")
'AirRo = CalculMasseVolumique(2, Airnames, Aircompo, 0, 101325, "[kg/m3]")
'
''Conversion volumique / massique
'mFum = QFum * FumRo / 3600 '[kg/s]
'
'mAir = mFum * (CalculEnthalp(Ngaz_fum, Fumnames, Fumcompo, TFumCelsius, TProtecCelsius, "[J/kg]")) / CalculEnthalp(2, Airnames, Aircompo, TProtecCelsius, TAirDilution, "[J/kg]")
'
'QAirDilution = mAir / AirRo * 3600 '[Nm3/h]
'
'End Sub
'
'Public Sub CalculTsortieFumees(Pechangee As Double, Ngaz_fum As Integer, Fumnames() As String, Fumcompo() As Double, QFum As Double, QAirDilution As Double, TEntreeCelsius As Double, TSortieCelsius As Double, Psortie As Double)
'
''---------------------------------------------------------------------------------------------------
'' Calcul de la température de sortie des fumées après récupérateur, en tenant compte de la dilution
''---------------------------------------------------------------------------------------------------
'' L. Ferrand6/02/2005
''---------------------------------------------------------------------------------------------------
'
''Pechangee : puissance transmise aux fumées+air dilution [W]
'
'Dim FumRo As Double ' Masse volumique des fumées
'Dim AirRo As Double
'Dim mFum As Double ' Débit de fumées [kg/s]
'Dim mAir As Double  'Debit d'air de dilution [kg/s]
'Dim Tinf As Double
'Dim Tsup As Double
'Dim p As Double
'Dim Airnames(2) As String
'Dim Aircompo(2) As Double
'Dim i As Integer
'
'i = 0
'Tinf = 0
'Tsup = TEntreeCelsius
'p = 0
'
'Airnames(1) = "O2"
'Airnames(2) = "N2"
'Aircompo(1) = 20.8
'Aircompo(2) = 79.2
'
'FumRo = CalculMasseVolumique(Ngaz_fum, Fumnames, Fumcompo, 0, 101325, "[kg/m3]")
'AirRo = CalculMasseVolumique(2, Airnames, Aircompo, 0, 101325, "[kg/m3]")
'
''Conversion volumique / massique
'mFum = QFum * FumRo / 3600 '[kg/s]
'mAir = QAirDilution * AirRo / 3600 '[kg/s]
'
'While Tsup - Tinf > 0.5 And i < 500
'    p = mFum * CalculEnthalp(Ngaz_fum, Fumnames, Fumcompo, TEntreeCelsius, 0.5 * (Tinf + Tsup), "[J/kg]") + mAir * CalculEnthalp(2, Airnames, Aircompo, TEntreeCelsius, 0.5 * (Tinf + Tsup), "[J/kg]")
'    If p > Pechangee Then Tinf = 0.5 * (Tinf + Tsup)
'    If p <= Pechangee Then Tsup = 0.5 * (Tinf + Tsup)
'    i = i + 1
'Wend
'
'TSortieCelsius = 0.5 * (Tinf + Tsup)
'Psortie = mFum * CalculEnthalp(Ngaz_fum, Fumnames, Fumcompo, TSortieCelsius, 0, "[J/kg]") + mAir * CalculEnthalp(2, Airnames, Aircompo, TSortieCelsius, 0, "[J/kg]")
'
'End Sub
'
'Public Function CalculSurfaceRecu(Pechangee As Double, FK As Double, T1E As Double, T1S As Double, T2E As Double, T2S As Double) As Double
'
'Dim DTML As Double
'
'DTML = ((T1E - T2S) - (T1S - T2E)) / (Log((T1E - T2S) / (T1S - T2E)))
'
'CalculSurfaceRecu = Pechangee / FK / DTML
'
'End Function
'
'
