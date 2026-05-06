Attribute VB_Name = "ModuleUpdate"
Sub UpdateFuelAir()

    Dim i As Integer
    Dim NbFuel As Integer
    Dim FuelNames() As String
    Dim FuelCompo() As Double
    
    Dim NbAir As Integer
    Dim AirNames() As String
    Dim AirCompo() As Double
    
    Dim PCS As Double
    
    With Worksheets(SGeom)
    
        
        '-------- Recherche du combustible dans la base et récupération de la compo -----
        Call Gaz_GetCompo(SFuel, .Range(CelFuel), NbFuel, FuelNames, FuelCompo)
        
        '------ Masse volumique -----
        .Range(CelDensity) = CalculMasseVolumique(NbFuel, FuelNames, FuelCompo, 0, 101325, "[kg/m3]")
        
        '------ Indice de Wobbe -----
        PCS = CalculPCS(NbFuel, FuelNames, FuelCompo, "[kJ/Nm3]")
        .Range(CelWobbe) = CalculWobbe(PCS, .Range(CelDensity).Offset(0, i) / 1.287, NbFuel, FuelNames, FuelCompo, "[MJ/Nm3]")

        '------ PCI -----
        .Range(CelLHV) = CalculPCI(NbFuel, FuelNames, FuelCompo, "[kJ/Nm3]")
        
      
        
        '---------- Recherche du comburant dans la base -----------
        Call Gaz_GetCompo(SComburant, .Range(CelAir).Offset(j), NbAir, AirNames, AirCompo)
        
        '----- Remplissage du tableau de compo Air ----
        For k = 1 To NbAir
            .Range(CelAirCompo).Offset(k, 0) = AirNames(k)
            .Range(CelAirPourc).Offset(k, 0) = AirCompo(k)
        Next k
        For k = NbAir + 1 To NMAX_AIR
            .Range(CelAirCompo).Offset(k, 0) = ""
            .Range(CelAirPourc).Offset(k, 0) = ""
        Next k
        '----- Remplissage du tableau de compo Air ----
        For k = 1 To NbFuel
            .Range(CelFuelCompo).Offset(k, 0) = FuelNames(k)
            .Range(CelFuelPourc).Offset(k, 0) = FuelCompo(k)
        Next k
        For k = NbFuel + 1 To NMAX_FUEl
            .Range(CelFuelCompo).Offset(k, 0) = ""
            .Range(CelAirPourc).Offset(k, 0) = ""
        Next k
        '----- Va / Vf -----
        .Range(CelVa).Offset(2 * j, i) = CalculVa(NbFuel, FuelNames, FuelCompo, NbAir, AirNames, AirCompo, "[Nm3/Nm3]")
        .Range(CelVf).Offset(2 * j, i) = CalculVfh(NbFuel, FuelNames, FuelCompo, NbAir, AirNames, AirCompo, "[Nm3/Nm3]")
        


       
        
    End With
End Sub
