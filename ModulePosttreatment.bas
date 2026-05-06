Attribute VB_Name = "ModulePosttreatment"
Function ThPostTreatment() As Boolean

    '-----------------------------------------------'
    ' Post-traitement des résultats de Thermette    '
    '-----------------------------------------------'
    ' P.Dubois 04/11/2025
    '-----------------------------------------------'
    ' Modifications:
    ' Initial
    '
    '-----------------------------------------------'
    
    Dim i As Integer
    Dim j As Integer
    Dim izone As Integer
    Dim NbZone As Integer
    
    Dim iLine As String
    
    Dim NbSect As Integer
    Dim NbSrip As Integer
    Dim IMaxRangee As Integer
    Dim NbRangee() As Integer
    Dim FluxMur() As Double ' W
    Dim FluxTube() As Double ' W
    Dim FluxStrip() As Double ' W
    Dim TempStrip() As Double ' °c
    Dim TempFum() As Double ' °c
    Dim FluxFum() As Double ' °c
    Dim EStripIn As Double
    Dim EStripOut As Double
    Dim FluxMurFour As Double
    Dim FluxMurZone() As Double
    Dim FluxTubeFour As Double
    Dim FluxTubeZone() As Double
    Dim FluxFumeeFour As Double
    Dim FluxFumZone() As Double
    Dim TempFumZone() As Double
    
    
    Dim DebGaz() As Double
    Dim DebAir() As Double
    Dim DebFum() As Double
    Dim DebGazZone() As Double
    Dim DebAirZone() As Double
    Dim DebFumZone() As Double
    Dim DebGazFour As Double
    Dim DebAirFour As Double
    Dim DebFumFour As Double
    
    Dim GazPCI As Double
    Dim GazVa0 As Double
    Dim GazDensite As Double
    Dim AirDensite As Double
    Dim FumDensite As Double
    Dim AirFuelRatio As Double
    
    Dim NbFuel As Variant
    Dim FuelNames As Variant
    Dim FuelCompo As Variant
    Dim NbAir As Integer
    Dim AirNames() As String
    Dim AirCompo() As Double
    Dim FumNames() As String
    Dim FumCompo() As Double
    
    Dim unite As String
    
    NbSect = Worksheets(SGeom).Range(CelNbSect)
'    NbStrip = Worksheets(SResult).Range(CelNbStrip)
    ReDim NbRangee(NbSect)
'    ReDim TempStrip(NbStrip)
    
    
    strChemFich = Worksheets(SFiles).Range(CThermette) + "\VRTF_resu"
    IMaxRangee = 0
    For i = 1 To NbSect
        NbRangee(i) = Worksheets(SGeom).Range(CelSectDesc).Offset(i, 0).Value
        IMaxRangee = Max(IMaxRangee, NbRangee(i))
    Next i
    i = 1
    NbZone = 0
    intFile = FreeFile
    Open strChemFich For Input As intFile
    While (Worksheets(SGeom).Range(CelReelZone).Offset(i, 0).Value > 0)
        NbZone = Max(NbZone, Worksheets(SGeom).Range(CelReelZone).Offset(i, 0))
   
    '---- Tableaux des résultats -----
  

    

    Line Input #intFile, iLine
 
    If InStr(iLine, "CALCUL TERMINE NORMALEMENT") = 0 Then
        Call MsgBox("Solution not converged. Post-treatment aborted.", vbCritical, "FOX")
        ThPostTreatment = False
        Exit Function
    Else
        ThPostTreatment = True
    End If
    ReDim TempStrip(NbSect, IMaxRangee)
    ReDim FluxMur(NbSect, IMaxRangee)
    ReDim FluxStrip(NbSect, IMaxRangee)
    ReDim FluxTube(NbSect, IMaxRangee)
    ReDim FluxFum(NbSect, IMaxRangee)
    ReDim TempFum(NbSect, IMaxRangee)
    ReDim FluxTubeZone(NbZone)
    ReDim FluxMurZone(NbZone)
    ReDim FluxFumZone(NbZone)
    ReDim TempFumZone(NbZone)
    ReDim DebGaz(NbSect, IMaxRangee)
    ReDim DebAir(NbSect, IMaxRangee)
    ReDim DebGazZone(NbZone)
    ReDim DebAirZone(NbZone)
    ReDim DebFum(NbSect, IMaxRangee)
    ReDim DebFumZone(NbZone)
    
    
    While (InStr(iLine, "Valeurs des sorties") = 0)
        Line Input #intFile, iLine
    Wend
    
    While Not EOF(intFile)
        Line Input #intFile, iLine

       
  
          
    '---------- Identification de la sortie / sollicitation -------
 
    
    '---- Température peau sup
        If (InStr(iLine, "SStrip_S") > 0) Then
            i = CInt(Mid(iLine, InStr(1, iLine, "SStrip_S") + 8, 1))
            j = CInt(Mid(iLine, InStr(1, iLine, "SStrip_S") + 11, 2))
            TempStrip(i, j) = CInt(Mid(iLine, InStr(1, iLine, ":") + 1, Len(iLine) - InStr(1, iLine, ":")))
            
        ElseIf (InStr(iLine, "SFluxMur_S") > 0) Then
            i = CInt(Mid(iLine, InStr(1, iLine, "SFluxMur_S") + 10, 1))
            j = CInt(Mid(iLine, InStr(1, iLine, "SFluxMur_S") + 13, 2))
            FluxMur(i, j) = CDbl(Mid(iLine, InStr(1, iLine, ":") + 1, Len(iLine) - InStr(1, iLine, ":")))
                    
        ElseIf (InStr(iLine, "SFluxStrip_S") > 0) Then
            i = CInt(Mid(iLine, InStr(1, iLine, "SFluxStrip_S") + 12, 1))
            j = CInt(Mid(iLine, InStr(1, iLine, "SFluxStrip_S") + 15, 2))
            FluxStrip(i, j) = CDbl(Mid(iLine, InStr(1, iLine, ":") + 1, Len(iLine) - InStr(1, iLine, ":")))
        
        ElseIf (InStr(iLine, "SFluxTube_S") > 0) Then
            i = CInt(Mid(iLine, InStr(1, iLine, "SFluxTube_S") + 11, 1))
            j = CInt(Mid(iLine, InStr(1, iLine, "SFluxTube_S") + 14, 2))
            FluxTube(i, j) = -CDbl(Mid(iLine, InStr(1, iLine, ":") + 1, Len(iLine) - InStr(1, iLine, ":")))
            
        ElseIf (InStr(iLine, "S_Prod_IN") > 0) Then
            EStripIn = CDbl(Mid(iLine, InStr(1, iLine, ":") + 1, Len(iLine) - InStr(1, iLine, ":")))
            
        ElseIf (InStr(iLine, "S_Prod_OUT") > 0) Then
            EStripOut = CDbl(Mid(iLine, InStr(1, iLine, ":") + 1, Len(iLine) - InStr(1, iLine, ":")))


             
        End If
         i = i + 1
    Wend
    Close #intOutFile
     Wend
    '----------- caractéristiques du Gaz
    GazPCI = Worksheets(SGeom).Range(CelLHV)
    GazVa0 = Worksheets(SGeom).Range(CelVa)
   
    '----------------- calcul des caractéristiques des fumées
    
    '-------- Recherche du combustible dans la base et récupération de la compo -----
    Application.Run "module_commun.xlam!Gaz_GetCompo", SFuel, Worksheets(SGeom).Range(CelFuel), NbFuel, Fuelname, FuelCompo
    Application.Run "module_commun.xlam!Gaz_GetCompo", SComburant, Worksheets(SGeom).Range(CelAir), NbAir, AirNames, AirCompo


    '---------- Recherche du comburant dans la base -----------
'    Application.Run "module_commun.xlam!Gaz_GetCompo", SComburant, Worksheets(SGeom).Range(CelAir), NbAir, AirNames, AirCompo
    AirFuelRatio = Worksheets(SGeom).Range(CelAirExces)
    Application.Run "module_commun.xlam!WasteGasCompo", FumNames(), FumCompo(), 293, NbFuel, FuelNames(), FuelCompo(), NbAir, AirNames(), AirCompo(), AirFuelRatio
    FumDensite = Application.Run("module_commun.xlam!CalculMasseVolumique", 5, FumNames, FumCompo, 0, 101325, "[kg/m3]")
    AirDensite = Application.Run("module_commun.xlam!CalculMasseVolumique", NbAir, AirNames, AirCompo, 0, 101325, "[kg/m3]")
    GazDensite = Application.Run("module_commun.xlam!CalculMasseVolumique", NbFuel, FuelNames, FuelCompo, 0, 101325, "[kg/m3]")

'------------ affichage dans les taleaux
     k = 1
    While (Worksheets(SGeom).Range(CelReelZone).Offset(k, 0).Value > 0)
         i = Worksheets(SGeom).Range(CelSect).Offset(k, 0).Value
         j = Worksheets(SGeom).Range(CelRangee).Offset(k, 0).Value
         
         izone = Worksheets(SGeom).Range(CelReelZone).Offset(k, 0).Value
         
         RendementTube = Worksheets(SGeom).Range(CelTubeEff).Offset(k, 0).Value
         
        ' flux dans les zones
         FluxTubeZone(izone) = FluxTubeZone(izone) + FluxTube(i, j)
         FluxTubeFour = FluxTubeFour + FluxTubeZone(izone)
         
          ' flux dans les murs
         FluxMurZone(izone) = FluxMurZone(izone) + FluxMur(i, j)
         FluxMurFour = FluxMurFour + FluxMurZone(izone)
         
         ' Debit Gaz air fumee en Nm3/s
         DebGaz(i, j) = FluxTube(i, j) / GazPCI / RendementTube
         DebAir(i, j) = DebGaz(i, j) * GazVa0 * (1 + AirFuelRatio / 100)
         DebFum(i, j) = (DebGaz(i, j) * GazDensite + DebAir(i, j) * AirDensite) / FumDensite
         unite = "[Nm3/h]"
        
        ' flux dans les fumees
         
         FluxFum(i, j) = (1 - RendementTube) / RendementTube * FluxTube(i, j)
         FluxFumZone(izone) = FluxFumZone(izone) + FluxFum(i, j)
         FluxFumFour = FluxFumeeFour + FluxFumZone(izone)
         ' temperature es fumees
'        TempFum(i, j) = CalculTemperature(5, Fumnames(), Fumcompo(), FluxFum(i, j), DebFum(i, j))
         TempFum(i, j) = Application.Run("module_commun.xlam!GasTemp", 5, FumNames(), FumCompo(), FluxFum(i, j), DebFum(i, j))

         k = k + 1
    Wend
'------------ calcul enthalpie et température  des fumées
'    For i = 1 To NbZone
'        FluxFumee
End Function



