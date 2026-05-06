Attribute VB_Name = "ModuleSover"
Public Sub SolveCombustion()

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim iso As Integer
    Dim jso As Integer
    
'    Dim frm As frmMessage ' Message de progression
'
'    Set frm = New frmMessage
'
'    FrmMessage.Title = "FOX"
'    FrmMessage.Show vbModeless

    
'    frm.Show vbModeless ' Ouvre mais rend la main au code
'
    '----- Déclaration des variables ---
    
    Dim NbSect As Integer
    
    '----- Données sur les réactifs
    Dim Combustibles() As String        ' Liste des combustibles des zones
    Dim Comburants() As String          ' Liste des comburants des zones
    
    Dim Combustible_NGaz() As Integer    ' Liste des nombres de gaz élémentaires composant les combustibles
    Dim Combustible_Names() As Variant   ' Liste des noms des gaz élémentaires composant les combustibles
    Dim Combustible_Compo() As Variant   ' Liste des pourcentages volumiques des gaz élémentaires composant les combustibles
    
    Dim Comburant_NGaz() As Integer    ' Liste des nombres de gaz élémentaires composant les Comburants
    Dim Comburant_Names() As Variant   ' Liste des noms des gaz élémentaires composant les Comburants
    Dim Comburant_Compo() As Variant   ' Liste des pourcentages volumiques des gaz élémentaires composant les Comburants
    
    Dim Compo() As Double ' Tableaux temporaires
    Dim Name() As String
                   
    Dim Ngaz_fuel As Integer
    Dim FuelCompo() As Double
    Dim FuelNames() As String
    
    Dim Ngaz_Air As Integer
    Dim AirCompo() As Double
    Dim AirNames() As String
    
    Dim nbFum As Integer
    Dim FumCompo() As Double
    Dim FumNames() As String
    
    Dim Puissances() As Double          ' Puissances nominales des brûleurs [W]
    Dim ExcesAir() As Double            ' Excès d'air des brûleurs en [%]
    
    Dim Puissance As Double             ' Variables temporaires
    Dim n As Double
    
    Dim DebitsFumees() As Double        ' Débits nominaux de fumees des brûleurs [kg/s]
    Dim DebitsAir() As Double            ' Débits nominaux de comburant [Nm3/h]
    Dim DebitsFuel() As Double           ' Débits nominaux de combustible [Nm3/h]
    
    Dim Debchaud(4) As Double
    Dim Debfroid(4) As Double

    Dim IB() As Integer          ' Index des zones où il y a des brûleurs
    Dim Nbruleur() As Integer

    Dim Tair() As Double
    Dim Tfuel() As Double
    
    Dim TairTakeOverP() As Double
    Dim TfuelTakeOverP() As Double
    

    '----- Données sur les fumées
    Dim ChaleurMassique() As Double     ' Pôlynome de la chaleur massique des fumées dans chaque zone
    Dim MasseVolumique() As Double      ' Masse volumique à 0° et 1atm des fumées dans chaque zone

    
    '------- Régénératifs ---
    Dim EffRegeFuel() As Double         ' Efficacité des régénérateurs de combustibles [%]
    Dim AspiRegeFuel() As Double        ' Taux d'aspiration des régénérateurs de combustibles [%]
    Dim EffRegeAir() As Double          ' Efficacité des régénérateurs de comburants [%]
    Dim AspiRegeAir() As Double         ' Taux d'aspiration des régénérateurs de comburants [%]
    
    Dim PointsTzone(4) As Double            ' Pour le calcul des pôlynomes de préchauffage des réactifs
    Dim PointsTairchaud(4) As Double
    Dim PointsTfuelchaud(4) As Double
    Dim PointsTfumeesfroides(4) As Double
    Dim PointsTfumeesEqAir(4) As Double
    Dim PointsTfumeesEqFuel(4) As Double
    

    ' Polynômes de Températures
    Dim poly(4) As Double

    Dim PolyTairchaud() As Double           ' Calcul de la température air préchauffé
    Dim PolyTfuelchaud() As Double          ' Calcul de la température Fuel préchauffé
    Dim PolyTfumeesfroidesAir() As Double   ' Calcul de la température fumées sortie régénérateur Air
    Dim PolyTfumeesfroidesFuel() As Double  ' Calcul de la température fumées sortie régénérateur Fuel
    
    Dim PolyTfumeesEqAir() As Double        ' Calcul de la température équivalente des fumées côté air
    Dim PolyTfumeesEqFuel() As Double       ' Calcul de la température équivalente des fumées côté fuel
    
    
    ' 4 points pour faire passer le polynôme du Cp et de l'enthalpie
    Dim temp(4) As Double ' Températures en Celsius
    Dim cp(4) As Double
    Dim Enthalp(4) As Double
    
    temp(1) = 273 + 0
    temp(2) = 273 + 400
    temp(3) = 273 + 800
    temp(4) = 273 + 1200
        
    Dim mAir() As Double ' kg/s
    Dim mFuel() As Double 'kg/s
    Dim EnthalpAir()
    Dim EnthalpFuel()
    Dim EnthalpFum()
    
    '------------------------------------------
    ' Données sur l'indexation entre brûleurs
    '------------------------------------------
    Dim nbindex As Integer                  ' Nombre d'indexations
    Dim Maitre_zone() As Integer            ' Tableau des n° de zones des brûleurs servant de source aux zones indexées
    Dim Esclave_zone() As Integer           ' Tableau des n° de zones des brûleurs à indexer sur un autre
    Dim Flowratio() As Double               ' Rapport de débit appliqué
    Dim esclave As Boolean

    '------------------------------------------
    ' Données sur les puissances imposées
    '------------------------------------------
    Dim Putile As Double
    Dim nbPimposee As Integer                   ' Nombre de puissances imposées
    Dim Pimposee_zone() As Integer              ' Tableau des n° de zones des brûleurs à puissance imposée
    Dim Pimposee() As Double                    ' Puissance imposée en [W]
    Dim ratio() As Double                       ' Ratio à appliquer dans les zones sup
    
    
    '------------------------------------------
    ' Données sur les puissances installées
    '------------------------------------------
    Dim nbPinstal As Integer                    ' Nombre de puissances installées
    Dim Pinstal_zone() As Integer               ' Tableau des n° de zones des brûleurs à puissance imposée
    Dim Pinstal() As Double                     ' Puissance installée en [W]
    Dim ratio_instal() As Double                   ' Ratio à appliquer dans les zones sup
        
    '------------------------------------------
    ' Données pour le post-traitement
    '------------------------------------------
    
    Dim QAirRecuTot As Double
    Dim QFumRecuTot As Double
    Dim QAirRecu() As Double
    Dim mAirRecu() As Double
    Dim mFuelRecu() As Double
    Dim mAirRege() As Double
    Dim mFuelRege() As Double
    Dim mFumRecu() As Double        ' Débit de fumées en [kg/s]
    Dim QfumRecu() As Double        ' Débit de fumées en [Nm3/h]
    Dim mFumRegeAir() As Double        ' Débit de fumées régénérés coté air en [kg/s]
    Dim mFumRegeFuel() As Double        ' Débit de fumées régénérés coté air en [kg/s]
    

    
    NbSect = Worksheets(SGeom).Range(CelZone)
    
    '----- Initialisation de valeurs -------
    ReDim Nbruleur(1)               ' Nombre de brûleurs zones sup (0) et inf (1)
    ReDim IB(NbSect)            ' Liste des indice des zones équipées de brûleurs sup (x,0) et inf (x,1)
    ReDim Puissances(NbSect)
    ReDim ratio_instal(NbSect)
    
    '----- Redimensionnement des tableaux ---
    ReDim PolyTairchaud(NbSect, 4) ' 4 coefficients, pour chaque zone sup et inf
    ReDim PolyTfuelchaud(NbSect, 4)
    ReDim PolyTfumeesfroidesAir(NbSect, 4)
    ReDim PolyTfumeesfroidesFuel(NbSect, 4)
    ReDim PolyTfumeesEqAir(NbSect, 4)
    ReDim PolyTfumeesEqFuel(NbSect, 4)
    
    ReDim Combustibles(NbSect)
    ReDim Comburants(NbSect)
    ReDim Combustible_NGaz(NbSect)
    ReDim Comburant_NGaz(NbSect)
    ReDim Combustible_Names(NbSect, NMAX_FUEl)
    ReDim Comburant_Names(NbSect, NMAX_AIR)
    ReDim Combustible_Compo(NbSect, NMAX_FUEl)
    ReDim Comburant_Compo(NbSect, NMAX_AIR)
   
    
    ReDim Tfuel(NbSect)
    ReDim Tair(NbSect)
    
    ReDim TfuelTakeOverP(NbSect)
    ReDim TairTakeOverP(NbSect)
    
    ReDim DebitsFumees(NbSect)        ' Débits nominaux de fumees des brûleurs [kg/s]
    ReDim DebitsAir(NbSect)             ' Débits nominaux de comburant [Nm3/h]
    ReDim DebitsFuel(NbSect)            ' Débits nominaux de combustible [Nm3/h]
    
    
    ReDim EffRegeFuel(NbSect)
    ReDim AspiRegeFuel(NbSect)
    ReDim EffRegeAir(NbSect)
    ReDim AspiRegeAir(NbSect)

    ReDim ExcesAir(NbSect)
    
    ReDim QAirRecu(NbSect)
    ReDim mAirRecu(NbSect)
    ReDim mFuelRecu(NbSect)
    ReDim mAirRege(NbSect)
    ReDim mFuelRege(NbSect)
    ReDim mFumRecu(NbSect)
    ReDim QfumRecu(NbSect)
    ReDim mFumRegeAir(NbSect)
    ReDim mFumRegeFuel(NbSect)
    
    
    ReDim ChaleurMassique(NbSect, 4)
    ReDim MasseVolumique(NbSect)
    
    ReDim mAir(NbSect)
    ReDim mFuel(NbSect)
    ReDim EnthalpAir(NbSect, 4)
    ReDim EnthalpFuel(NbSect, 4)
    ReDim EnthalpFum(NbSect, 4)
    
    

    
    With Worksheets(SBurner)

        For i = 1 To NbSect
                
                    
            '-------- Calculs des polynômes températures  -----
            
            Nbruleur(j) = Nbruleur(j) + 1
            
            '--- Indice brûleur
            IB(Nbruleur(j)) = i
            
            
            '------ Choix de la puissance nominale des brûleurs
            Puissances(i) = .Range(CelInstalledP).Offset(i, 0) * 1000
              
        Next i
        
        '---------------------------------------------------------------------
    
        
        
        For i = 1 To NbSect
                     
                    
            '-------------------------------------------------------------------------------------'
            '   Etape 1 :remplissage des tableaux concernant les réactifs et les brûleurs         '
            '-------------------------------------------------------------------------------------'
                
            'Initialisation des polynomes
            
            For k = 1 To 4
                PolyTairchaud(i, k) = 0
                PolyTfuelchaud(i, k) = 0
                PolyTfumeesfroidesAir(i, k) = 0
                PolyTfumeesfroidesFuel(i, k) = 0
                PolyTfumeesEqAir(i, k) = 0
                PolyTfumeesEqFuel(i, k) = 0
            Next k
                        
                        
                        
            '--- Combustible
            Combustibles(i) = .Range(CelFuel)
            
            '--- Température au TOP
         
            Tfuel(i) = .Range(CelFuelTop) + 273
            TfuelTakeOverP(i) = .Range(CelFuelTop) + 273
           
                        
            '--- Valeur polynome si pas de régénératif sur le gaz
            PolyTfuelchaud(i, 1) = Tfuel(i)
            
           
                        
            '----- Comburants -------
                
            Comburants(i) = .Range(CelAir)
            ExcesAir(i) = (.Range(CelRatio).Offset(i, 0) - 1) * 100
            
            
            TairTakeOverP(i) = .Range(CelAirTop) + 273
                            
            Tair(i) = .Range(CelAirTop) + 273
                         
                                            
                        
            '--- Valeur polynome si pas de régénératif sur l'air
            
            PolyTairchaud(i, 1) = Tair(i) ' Si pas de régénératif, cette valeur sera prise
        
            '---- Si régénération sur le comburant
             If (Worksheets(SBurner).Range(CelTypAir).Offset(i, 0) = Worksheets(SOption).Range(CelOptionAir)) Then
                           
                EffRegeAir(i) = Worksheets(SBurner).Range(CelRegenAireff)
                AspiRegeAir(i) = Worksheets(SBurner).Range(CelRegenAirAspi)
            Else
                EffRegeAir(i) = 0
                AspiRegeAir(i) = 0
            End If
                        
            
            '------ Remplissage des tableaux de combustible
            
            Call Gaz_GetCompo(SFuel, Combustibles(i), Combustible_NGaz(i), Name, Compo)
            For k = 1 To Combustible_NGaz(i)
                Combustible_Names(i, k) = Name(k)
                Combustible_Compo(i, k) = Compo(k)
            Next k
            '------ Remplissage des tableaux de comburant
            Call Gaz_GetCompo(SComburant, Comburants(i), Comburant_NGaz(i), Name, Compo)
            For k = 1 To Comburant_NGaz(i)
                Comburant_Names(i, k) = Name(k)
                Comburant_Compo(i, k) = Compo(k)
            Next k
              
        Next i
                
        
                                  
        '---------------------------------------------------------------------
        For i = 1 To NbSect
            
                    
            '----- Calcul des débits nominaux
            ' Initialisation des variables
            Ngaz_fuel = Combustible_NGaz(i)
            Ngaz_Air = Comburant_NGaz(i)
            ReDim FuelNames(Ngaz_fuel)
            ReDim FuelCompo(Ngaz_fuel)
            ReDim AirNames(Ngaz_Air)
            ReDim AirCompo(Ngaz_Air)
            
            ReDim FumNames(5)
            ReDim FumCompo(5)
                    
            For k = 1 To Ngaz_fuel
                FuelNames(k) = Combustible_Names(i, k)
                FuelCompo(k) = Combustible_Compo(i, k)
            Next k
            
            For k = 1 To Ngaz_Air
                AirNames(k) = Comburant_Names(i, k)
                AirCompo(k) = Comburant_Compo(i, k)
            Next k
            
            Puissance = Puissances(i) / 1000
            n = ExcesAir(i)
        
            ' Calcul du débit de fumées nominal de la zone en [kg/s]
            DebitsFumees(i) = CalculDebitFumees(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[kg/s]")
            ' Calcul du débit de comburant nominal de la zone en [Nm3/h]
            DebitsAir(i) = CalculDebitAir(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[Nm3/h]")
            ' Calcul du débit de combustible de la zone en [Nm3/h]
            DebitsFuel(i) = CalculDebitFuel(Ngaz_fuel, FuelNames, FuelCompo, Puissance, "[Nm3/h]")

            ' Calcul du débit de comburant nominal de la zone en [kg/s]
            mAir(i) = CalculDebitAir(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[kg/s]")
            ' Calcul du débit de combustible de la zone en [kg/s]
            mFuel(i) = CalculDebitFuel(Ngaz_fuel, FuelNames, FuelCompo, Puissance, "[kg/s]")

            ' Calcul du polynome d'enthalpie sur l'air
            Enthalp(1) = CalculEnthalp(Ngaz_Air, AirNames, AirCompo, temp(1) - 273, 0, "[J/kg]")
            Enthalp(2) = CalculEnthalp(Ngaz_Air, AirNames, AirCompo, temp(2) - 273, 0, "[J/kg]")
            Enthalp(3) = CalculEnthalp(Ngaz_Air, AirNames, AirCompo, temp(3) - 273, 0, "[J/kg]")
            Enthalp(4) = CalculEnthalp(Ngaz_Air, AirNames, AirCompo, temp(4) - 273, 0, "[J/kg]")
            
            Call Polint3(temp, Enthalp, poly)
                   
            For k = 1 To 4
                EnthalpAir(i, k) = poly(k)
            Next k
                    
                    
            ' Calcul du polynome d'enthalpie sur le gaz
            Enthalp(1) = CalculEnthalp(Ngaz_fuel, FuelNames, FuelCompo, temp(1) - 273, 0, "[J/kg]")
            Enthalp(2) = CalculEnthalp(Ngaz_fuel, FuelNames, FuelCompo, temp(2) - 273, 0, "[J/kg]")
            Enthalp(3) = CalculEnthalp(Ngaz_fuel, FuelNames, FuelCompo, temp(3) - 273, 0, "[J/kg]")
            Enthalp(4) = CalculEnthalp(Ngaz_fuel, FuelNames, FuelCompo, temp(4) - 273, 0, "[J/kg]")
            
            Call Polint3(temp, Enthalp, poly)
                   
            For k = 1 To 4
                EnthalpFuel(i, k) = poly(k)
            Next k
                    
            '------- Débits pour le post-traitement
                                
            '----- Si pas de régénératif sur l'air
            If AspiRegeAir(i) = 0 Then
                mAirRecu(i) = CalculDebitAir(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[kg/s]")
                QAirRecu(i) = DebitsAir(i)
            End If
            
            '----- Si pas de rege sur le gaz
            If AspiRegeFuel(i) = 0 Then
                mFuelRecu(i) = CalculDebitFuel(Ngaz_fuel, FuelNames, FuelCompo, Puissance, "[kg/s]")
            End If
            
            QfumRecu(i) = 0 ' initialisation
            mFumRecu(i) = 0 ' initialisation
            
            '----- Si régé sur l'air
            If AspiRegeAir(i) <> 0 And AspiRegeFuel(i) = 0 Then
                mAirRege(i) = CalculDebitAir(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[kg/s]")
                mFumRegeAir(i) = AspiRegeAir(i) / 100 * CalculDebitFumees(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[kg/s]")
                QfumRecu(i) = (100 - AspiRegeAir(i)) / 100 * CalculDebitFumees(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[Nm3/h]")
                mFumRecu(i) = (100 - AspiRegeAir(i)) / 100 * CalculDebitFumees(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[kg/s]")
            End If
            
       
            
            '----- Si ni régé air, ni régé fuel
            If AspiRegeAir(i) = 0 And AspiRegeFuel(i) = 0 Then
                QfumRecu(i) = CalculDebitFumees(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[Nm3/h]")
                mFumRecu(i) = CalculDebitFumees(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance, 1 + 0.01 * n, "[kg/s]")
            End If

            '----- Composition des fumées
            FumNames(1) = "CO2"
            FumNames(2) = "H2O"
            FumNames(3) = "O2"
            FumNames(4) = "N2"
            FumNames(5) = "SO2"

            FumCompo(1) = CalculFracVolCO2(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), 1 + 0.01 * n)
            FumCompo(2) = CalculFracVolH2O(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), 1 + 0.01 * n)
            FumCompo(3) = CalculFracVolO2(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), 1 + 0.01 * n)
            FumCompo(4) = CalculFracVolN2(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), 1 + 0.01 * n)
            FumCompo(5) = CalculFracVolSO2(Ngaz_fuel, FuelNames(), FuelCompo(), Ngaz_Air, AirNames(), AirCompo(), 1 + 0.01 * n)
            
            ' Calcul du polynome d'enthalpie sur les fumées
            Enthalp(1) = CalculEnthalp(5, FumNames, FumCompo, temp(1) - 273, 0, "[J/kg]")
            Enthalp(2) = CalculEnthalp(5, FumNames, FumCompo, temp(2) - 273, 0, "[J/kg]")
            Enthalp(3) = CalculEnthalp(5, FumNames, FumCompo, temp(3) - 273, 0, "[J/kg]")
            Enthalp(4) = CalculEnthalp(5, FumNames, FumCompo, temp(4) - 273, 0, "[J/kg]")
            
            Call Polint3(temp, Enthalp, poly)
                   
            For k = 1 To 4
                EnthalpFum(i, k) = poly(k)
            Next k
            
            
            Worksheets(SCombustion).Range(CelCFumNb).Offset(0, i) = 5
            For k = 1 To 5
                Worksheets(SCombustion).Range(CelCFumNames).Offset(5 * j + k - 1, i) = FumNames(k)
                Worksheets(SCombustion).Range(CelCFumCompo).Offset(5 * j + k - 1, i) = FumCompo(k)
            Next k
                    
                    
            '----- Température équivalente des fumées
            
            '-----------------------------------------------------------------------------------
            ' 1er cas : pas de système régénératif, la température équivalente est constante
            '-----------------------------------------------------------------------------------
            If (AspiRegeAir(i) = 0 And mFuel(i) > 0) Then
                PolyTfumeesEqAir(i, 1) = CalculTeqAB(Ngaz_Air, AirNames, AirCompo, mAir(i), Tair(i) - 273, 5, FumNames, FumCompo, mAir(i)) + 273
            End If
           
            If (mFuel(i) > 0) Then
                PolyTfumeesEqFuel(i, 1) = CalculTeqAB(Ngaz_fuel, FuelNames, FuelCompo, mFuel(i), Tfuel(i) - 273, 5, FumNames, FumCompo, mFuel(i)) + 273
            End If
                    
            '-----------------------------------------------------------------------------------
            ' 2ème cas : système régénératif sur l'air
            '-----------------------------------------------------------------------------------
            If (AspiRegeAir(i) <> 0) Then


    '            On calcule :
                ' - les propriétés des fumées de la zone
                ' - la température de préchauffage de l'air pour plusieurs températures de zone
                ' - la température des fumées équivalente côté air correspondante
                ' On détermine les coefficients des polynômes de la loi TfumeesEqAir = f(Tzone)
                            
                          
                ' 4 points pour identifier les pôlynomes de degré 3 en Kelvin
                For k = 1 To 4
                    PointsTzone(k) = 773 + (k - 1) * 300 'Un point tous les 300 Kelvin
                    
                    'Calcul de l'échangeur
                    Debchaud(k) = mFumRegeAir(i) 'CalculDebitFumees(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance + 0.01 * n, "[kg/s]")
                    Debfroid(k) = mAirRege(i) 'CalculDebitAir(Ngaz_fuel, FuelNames, FuelCompo, Ngaz_Air, AirNames, AirCompo, Puissance + 0.01 * n, "[kg/s]")
  
                    Call CalculEchangeur(5, FumNames, FumCompo, Debchaud(k), PointsTzone(k), Ngaz_Air, AirNames, AirCompo, Debfroid(k), Tair(i), EffRegeAir(i), 100, PointsTfumeesfroides(k), PointsTairchaud(k))
                    PointsTfumeesEqAir(k) = CalculTeqAB(Ngaz_Air, AirNames, AirCompo, mAirRege(i), PointsTairchaud(k), 5, FumNames, FumCompo, mAirRege(i))
                    
                Next k
                            
                Call Polint3(PointsTzone, PointsTairchaud, poly)
                
                For k = 1 To 4
                    PolyTairchaud(i, k) = poly(k)
                Next k
                
                Call Polint3(PointsTzone, PointsTfumeesfroides, poly)
                For k = 1 To 4
                    PolyTfumeesfroidesAir(i, k) = poly(k)
                Next k
    
                Call Polint3(PointsTzone, PointsTfumeesEqAir, poly)
                For k = 1 To 4
                    PolyTfumeesEqAir(i, k) = poly(k)
                Next k
             
            End If
                                
                                 
                    
                    '------------------------------------------------------'
                    'Etape 3 : Propriétés thermophysiques des fumées       '
                    '------------------------------------------------------'

    
                    cp(1) = CalculCp(5, FumNames(), FumCompo(), temp(1) - 273)
                    cp(2) = CalculCp(5, FumNames(), FumCompo(), temp(2) - 273)
                    cp(3) = CalculCp(5, FumNames(), FumCompo(), temp(3) - 273)
                    cp(4) = CalculCp(5, FumNames(), FumCompo(), temp(4) - 273)
                    
                    Call Polint3(temp, cp, poly)
                           
                    For k = 1 To 4
                        ChaleurMassique(i, k) = poly(k)
                    Next k

                    MasseVolumique(i) = CalculMasseVolumique(5, FumNames, FumCompo, 0, 101325, "[kg/Nm3]")
                    
                                
                    
                    '------------------------------------------------------------------------------
                    '------- Ecriture des résultats dans la feuille Combustion pour ces zones
                    '------------------------------------------------------------------------------
                    
                    Worksheets(SCombustion).Range(CelCFuel).Offset(0, i) = Combustibles(i)
                    
                    Worksheets(SCombustion).Range(CelCFuelNb).Offset(0, i) = Combustible_NGaz(i)
                    
                    For k = 1 To Combustible_NGaz(i)
                        Worksheets(SCombustion).Range(CelCFuelNames).Offset(k - 1, i) = Combustible_Names(i, k)
                        Worksheets(SCombustion).Range(CelCFuelCompo).Offset(k - 1, i) = Combustible_Compo(i, k)
                    Next k
                    
                    Worksheets(SCombustion).Range(CelCAir).Offset(0, i) = Comburants(i)
                    
                    Worksheets(SCombustion).Range(CelCAirNb).Offset(0, i) = Comburant_NGaz(i)
                    
                    For k = 1 To Comburant_NGaz(i)
                        Worksheets(SCombustion).Range(CelCAirNames).Offset(k - 1, i) = Comburant_Names(i, k)
                        Worksheets(SCombustion).Range(CelCAirCompo).Offset(k - 1, i) = Comburant_Compo(i, k)
                    Next k
                    
                    For k = 1 To 4
                        Worksheets(SCombustion).Range(CelCPolyTairchaud).Offset(k - 1, i) = PolyTairchaud(i, k)
                    Next k
                    
                    For k = 1 To 4
                        Worksheets(SCombustion).Range(CelCPolyTfuelchaud).Offset(k - 1, i) = PolyTfuelchaud(i, k)
                    Next k
                    
                    
                    For k = 1 To 4
                        Worksheets(SCombustion).Range(CelCPolyTfumeesfroidesAir).Offset(k - 1, i) = PolyTfumeesfroidesAir(i, k)
                    Next k
                    
                    For k = 1 To 4
                        Worksheets(SCombustion).Range(CelCPolyTfumeesfroidesFuel).Offset(k - 1, i) = PolyTfumeesfroidesFuel(i, k)
                    Next k
                    
                    
                    For k = 1 To 4
                        Worksheets(SCombustion).Range(CelCPolyTfumeesEqAir).Offset(k - 1, i) = PolyTfumeesEqAir(i, k)
                    Next k
                    
                    For k = 1 To 4
                        Worksheets(SCombustion).Range(CelCPolyTfumeesEqFuel).Offset(k - 1, i) = PolyTfumeesEqFuel(i, k)
                    Next k
                    
                    
                    Worksheets(SCombustion).Range(CelCTfuel).Offset(0, i) = Tfuel(i)
                    Worksheets(SCombustion).Range(CelCTair).Offset(0, i) = Tair(i)
                    Worksheets(SCombustion).Range(CelCDebitFuel).Offset(0, i) = DebitsFuel(i)
                    Worksheets(SCombustion).Range(CelCDebitAir).Offset(0, i) = DebitsAir(i)
                    
                    Worksheets(SCombustion).Range(CelCDebitFumees).Offset(0, i) = DebitsFumees(i)
                                       
                    
                    Worksheets(SCombustion).Range(CelCEffRegeFuel).Offset(0, i) = EffRegeFuel(i)
                    Worksheets(SCombustion).Range(CelCAspiRegeFuel).Offset(0, i) = AspiRegeFuel(i)
                    Worksheets(SCombustion).Range(CelCEffRegeAir).Offset(0, i) = EffRegeAir(i)
                    Worksheets(SCombustion).Range(CelCAspiRegeAir).Offset(0, i) = AspiRegeAir(i)
                    
                    Worksheets(SCombustion).Range(CelCExcesAir).Offset(0, i) = ExcesAir(i)
                    Worksheets(SCombustion).Range(CelCQAirRecu).Offset(0, i) = QAirRecu(i)
                    Worksheets(SCombustion).Range(CelCmAirRecu).Offset(0, i) = mAirRecu(i)
                    Worksheets(SCombustion).Range(CelCmFuelRecu).Offset(0, i) = mFuelRecu(i)
                    Worksheets(SCombustion).Range(CelCmAirRege).Offset(0, i) = mAirRege(i)
                    Worksheets(SCombustion).Range(CelCmFuelRege).Offset(0, i) = mFuelRege(i)
                    Worksheets(SCombustion).Range(CelCmFumRecu).Offset(0, i) = mFumRecu(i)
                    Worksheets(SCombustion).Range(CelCQFumRecu).Offset(0, i) = QfumRecu(i)
                    Worksheets(SCombustion).Range(CelCmFumRegeAir).Offset(0, i) = mFumRegeAir(i)
                    Worksheets(SCombustion).Range(CelCmFumRegeFuel).Offset(0, i) = mFumRegeFuel(i)
                    
                    For k = 1 To 4
                        Worksheets(SCombustion).Range(CelCChaleurMassique).Offset(k - 1, i) = ChaleurMassique(i, k)
                    Next k
                    
                    Worksheets(SCombustion).Range(CelCMasseVolumique).Offset(0, i) = MasseVolumique(i)
                            
                    For k = 1 To 4
                        Worksheets(SCombustion).Range(CelCEnthalpieAir).Offset(k - 1, i) = EnthalpAir(i, k)
                        Worksheets(SCombustion).Range(CelCEnthalpieFuel).Offset(k - 1, i) = EnthalpFuel(i, k)
                        Worksheets(SCombustion).Range(CelCEnthalpieFum).Offset(k - 1, i) = EnthalpFum(i, k)
                    Next k
                    
                    Worksheets(SCombustion).Range(CelCmAir).Offset(0, i) = mAir(i)
                    Worksheets(SCombustion).Range(CelCmFuel).Offset(0, i) = mFuel(i)
                    
                    Worksheets(SCombustion).Range(CelCTairTakeOverP).Offset(0, i) = TairTakeOverP(i)
                    Worksheets(SCombustion).Range(CelCTfuelTakeOverP).Offset(0, i) = TfuelTakeOverP(i)
                            
                
                
                
                Worksheets(SCombustion).Range(CelCPuis).Offset(0, i) = Puissances(i)
                Worksheets(SCombustion).Range(CelCRatio).Offset(0, i) = ratio_instal(i)
                    


        Next i


                  
                  
                  
    End With
    
    '--------- Ecriture des nb bruleurs en en-tête de la feuille combustion
    
    
    Worksheets(SCombustion).Range(CelCNbBurners).Offset(0, 0) = Nbruleur(0)
    Worksheets(SCombustion).Range(CelCNbBurners).Offset(1, 0) = Nbruleur(1)
    
    For i = 1 To NMAX_ZONES
        For j = 0 To 1
            Worksheets(SCombustion).Range(CelCIBBurners).Offset(0, i) = ""
        Next j
    Next i
    
'    For i = 1 To Max(Nbruleur(0), Nbruleur(1))
'
'        Worksheets(SCombustion).Range(CelCIBBurners).Offset(0, i) = IB(i)
'
'    Next i

       FrmMessage.CloseForm ' Ferme la fenêtre d'avancement


End Sub
