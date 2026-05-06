Attribute VB_Name = "Module_thermette"
Public Nomgaz() As String
Public CompGaz() As Double
Public Ngaz As Integer


Public Sub ThTablesPropAcier(strChem As String, Bisra As String, BisraCoSheet As String, BisraRoSheet As String, BisraCpSheet As String, CelOrig As String)
 Dim strFic As String
    Dim intFile As Integer
    
  
    Dim i As Integer
    Dim j As Integer
    Dim NbLigne As Integer

    j = 1
    While Bisra <> Worksheets(BisraCpSheet).Range(CelOrig).Offset(0, j)
        j = j + 1
    Wend
    NbLigne = Worksheets(BisraCpSheet).Range(CelOrig).Offset(-1, 0)
          

    strFic = strChem + "\AcierCp" '"c:\thermette\AcierCp"
    intFile = FreeFile
    Open strFic For Output As #intFile 'ouvert en écriture
    NbLigne = Worksheets(BisraCpSheet).Range(CelOrig).Offset(-1, 0)
    Print #intFile, CStr(NbLigne)
    For i = 1 To NbLigne
        Print #intFile, Worksheets(BisraCpSheet).Range(CelOrig).Offset(i, 0); Worksheets(BisraCpSheet).Range(CelOrig).Offset(i, j)
    Next i
    
    Close #intFile
   
    strFic = strChem + "\AcierCo" '"c:\thermette\AcierCo"
    intFile = FreeFile
    Open strFic For Output As #intFile 'ouvert en écriture

    NbLigne = Worksheets(BisraCpSheet).Range(CelOrig).Offset(-1, 0)
    Print #intFile, CStr(NbLigne)
    For i = 1 To NbLigne
        Print #intFile, Worksheets(BisraCoSheet).Range(CelOrig).Offset(i, 0); Worksheets(BisraCoSheet).Range(CelOrig).Offset(i, j)
    Next i
    
    
    Close #intFile
    
    strFic = strChem + "\AcierRo" '"c:\thermette\AcierRo"
    intFile = FreeFile
    Open strFic For Output As #intFile 'ouvert en écriture

    NbLigne = Worksheets(BisraRoSheet).Range(CelOrig).Offset(-1, 0)
    Print #intFile, CStr(NbLigne)
    For i = 1 To NbLigne
        Print #intFile, Worksheets(BisraRoSheet).Range(CelOrig).Offset(i, 0); Worksheets(BisraRoSheet).Range(CelOrig).Offset(i, j)
    Next i
    Close #intFile
   

End Sub


Sub Thnetwork()
'--------------------------'
'Données Produits  '
'--------------------------'

Dim Ep_prod As Double    'épaisseur produit
Dim Larg_prod As Double  'largeur produit
Dim Prod As Double  'largeur produit
Dim Acier As String  'type de produit
Dim T_entreebande As Double
Dim strChemFich As String
'--------------------------'
'Données Zones  '
'--------------------------'
Dim NbSect As Integer     'nombre de sections
Dim NbRangee() As Integer    'nombre de rangee
Dim H_zone As Double       'hauteur des zones
Dim Larg_Four As Double    'largeur des zones
Dim Gap_B() As Double      'écartement entre bandes d'une même zone

Dim Ep_mur As Double
Dim Cond_mur As Double
'--------------------------'
'Description rangees  '
'--------------------------'
Dim Type_tube() As String
Dim a() As Double
Dim b() As Double
Dim c() As Double
Dim d() As Double
Dim Nbtube() As Double
Dim T_tube() As Double
Dim Prem_Abs() As Double
Dim Distance_BT() As Double
Dim Pas_tube() As Double
Dim emis() As Double       'emissivité bande pour indice som2

'-------------------------------
'initialisation
'-------------------------------
ChemThermette = Worksheets(SFiles).Range(CThermette)
ChemModray = Worksheets(SFiles).Range(CModray)

'--------------------Création fichier base données fluides------------------'
      
    NomFic = Worksheets(SFiles).Range(CThermette) + "\fluides.prp"
    intOutFile = FreeFile
    
    Open NomFic For Output As intOutFile
    
    '-------------Acier----------------------------'
    Print #intOutFile, "#BISRA 0 3000"
    Print #intOutFile, "table(T,AcierCo)"
    Print #intOutFile, "table(T,AcierRo)"
    Print #intOutFile, "table(T,AcierCp)"
    Print #intOutFile, ""
    Close intOutFile

With Worksheets(SGeom)

    NbSect = .Range(CelNbSect)

    ReDim NbRangee(NbSect)
    ReDim Gap_B(NbSect)
    ReDim Type_tube(NbSect, 10)
    ReDim a(NbSect, 10)
    ReDim b(NbSect, 10)
    ReDim c(NbSect, 10)
    ReDim d(NbSect, 10)
    ReDim Nbtube(NbSect, 10)
    ReDim T_tube(NbSect, 10)
    ReDim Prem_Abs(NbSect, 10)
    ReDim Distance_BT(NbSect, 10)
    ReDim Pas_tube(NbSect, 10)
    ReDim emis(NbSect, 10)


    For i = 1 To NbSect
        NbRangee(i) = .Range(CelSectDesc).Offset(i, 0).Value
    Next i


'--------------------------'
'affectation des valeurs aux variables           '
'--------------------------'

           Ep_prod = .Range(CelEpProd).Value * 0.001
           Larg_prod = .Range(CelLargProd).Value
           Prod = .Range(CelProd).Value * 1000 / 3600
           T_entreebande = .Range(CelTempEntree).Value + 273
           Acier = .Range(CelTypeProd).Value
           H_zone = .Range(CelHautSect).Value
           Larg_Four = .Range(CelLargSect).Value
           Ep_mur = .Range(CelMurEp).Value
           Cond_mur = .Range(CelMurCond).Value

           For i = 1 To NbSect
               Gap_B(i) = .Range(CelDiamRoul).Offset(i, 0).Value
           Next i
           k = 1
           While (.Range(CelReelZone).Offset(k, 0).Value > 0)
            i = .Range(CelSect).Offset(k, 0).Value
            j = .Range(CelRangee).Offset(k, 0).Value
          
       
                   Type_tube(i, j) = .Range(CelTubeType).Offset(k, 0).Value
                   a(i, j) = .Range(CelA).Offset(k, 0).Value
                   b(i, j) = .Range(CelB).Offset(k, 0).Value / 2
                   c(i, j) = .Range(CelC).Offset(k, 0).Value
                   d(i, j) = .Range(CelD).Offset(k, 0).Value
                   Prem_Abs(i, j) = .Range(CelPreAbs).Offset(k, 0).Value
                   Distance_BT(i, j) = .Range(CelDiamRoul).Offset(i, 0).Value / 2
                   Pas_tube(i, j) = .Range(CelTubePitch).Offset(k, 0).Value
                   Nbtube(i, j) = .Range(CelNBTube).Offset(k, 0).Value
                   T_tube(i, j) = .Range(CelTubeTemp).Offset(k, 0).Value + 273
                   emis(i, j) = .Range(CelTubeEmis).Offset(k, 0).Value
                   k = k + 1
             Wend
    End With

'------------------------------------------------'
' Etape 1 : Construction du réseau thermette     '
'------------------------------------------------'
    strChemFich = ChemThermette + "\VRTF_reseau"
    intOutFile = FreeFile
    Open strChemFich For Output As intOutFile
    Close #intOutFile

    ' Ecriture des tables des propriétés thermophysiques de l'acier
    Call ThTablesPropAcier(ChemThermette, Acier, SBisraCo, SBisraRo, SBisraCp, CelDataOrig)

    ' Ecriture des sollicitations
    Call ThRes_Sollicitations(strChemFich, NbSect, NbRangee(), Prod, T_entreebande, T_tube())
    ' Ecriture des branches parois
    Call ThRes_Parois(strChemFich, NbSect, NbRangee(), H_zone, Larg_Four, Ep_mur, Cond_mur, Gap_B())
    ' Ecriture des branches bandes
    Call ThRes_Bandes(strChemFich, NbSect, NbRangee(), Larg_prod, Ep_prod, T_entreebande, Acier, H_zone, Prod)
    ' Ecriture des volumes tubes
    Call ThRes_Tubes(strChemFich, NbSect, NbRangee(), a(), b(), c(), d(), Nbtube(), Type_tube())
    ' Ecriture des echanges
    Call ThRes_RadEx(strChemFich)
'--------------------------------------------------------'
' Etape 2 : Construction du fichier calcul thermette     '
'--------------------------------------------------------'
    strChemFich = ChemThermette + "\VRTF_calc"
    intOutFile = FreeFile
    Open strChemFich For Output As intOutFile
    Close #intOutFile
    
    Call ThCalc(strChemFich, NbSect, NbRangee(), Nbtube())
    
    '--------------------------------------------------------'
    ' Etape 3 : Lancement du solveur Thermette               '
    '--------------------------------------------------------'
    Call GetCurrentDirectory(lstr, strCurrentDir)
    SetCurrentDirectory (Worksheets(SFiles).Range(CThermette) + "\")
    
    ShellWait (Worksheets(SFiles).Range(CThermette) + "\Thermette.exe " + _
            Worksheets(SFiles).Range(CThermette) + "\VRTF_reseau " + _
            Worksheets(SFiles).Range(CThermette) + "\VRTF_calc " + _
            Worksheets(SFiles).Range(CThermette) + "\VRTF_resu")
    
'    Call ThGetResults(Worksheets(SFiles).Range(CThermette) + "\VRTF_resu")
    
    SetCurrentDirectory (strCurrentDir)
    End Sub
  
Public Sub ThRes_RadEx(strChemFich As String)


    '------------------------------------------------------
    ' Ecriture des échanges radiatifs
    '------------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim intOutFile As Integer
    
    
    intOutFile = FreeFile
    Open strChemFich For Append As intOutFile

    Print #intOutFile, ""
    With Worksheets(SMesh)

        For j = 1 To .Range("A1")
            Print #intOutFile, .Range("A1").Offset(j, 0)
        Next j

    
    End With
    Close #intOutFile

End Sub

Public Sub ThRes_Sollicitations(strChemFich As String, NbSect As Integer, NbRangee() As Integer, Prod As Double, T_entreebande As Double, T_tube() As Double)
    
    '----------------------------------------------------------------------------------------------------'
    ' Ecriture des sollicitations dans le fichier réseau Thermette
    '----------------------------------------------------------------------------------------------------'
    ' A. Ceccotti, 29/06/2006
    '----------------------------------------------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' NbSect : nombre de zones du four
    ' prod : production en t/h
    ' T_tube : température tubes de la rangée
    ' T_entreebande : température entrée bande en °C
    '-----------------------------------------------------------------------------------------------------'
    
    Dim i As Integer
    Dim k As Integer
    Dim temp1 As String
    Dim temp2 As String
    Dim temp3 As String
    
    intOutFile = FreeFile
    Open strChemFich For Append As intOutFile
    
    'Température ambiante
    Print #intOutFile, "**sollicit Text 293"
    Print #intOutFile, "**sollicit Tref 273"
    
    'Débit bande en kg/s
    Print #intOutFile, "**sollicit prod "; Prod
    Print #intOutFile, "**sollicit inv_prod "; -1 * Prod
    
    'Température entrée bande
    Print #intOutFile, "**sollicit T_entreebande "; T_entreebande
    
    'Températures des tubes [K]
    For i = 1 To NbSect
        temp1 = i
        
        For j = 1 To NbRangee(i)

            temp2 = j
            temp3 = "**sollicit T_tube_Z" + temp1 + "_R" + temp2
            Print #intOutFile, temp3; T_tube(i, j)
        Next j

    Next i

        
Close #intOutFile

End Sub

Public Sub ThRes_Parois(strChemFich As String, NbSect As Integer, NbRangee() As Integer, H_zone As Double, Larg_zone As Double, Ep_mur As Double, cond As Double, Gap_B() As Double)
    '-----------------------------------------------------------------'
    ' Ecriture des parois dans le fichier réseau Thermette
    '-----------------------------------------------------------------'
    ' A. Ceccotti, 29/06/2006
    '-----------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' NbSect : nombre de zones du four
    ' Nbrangee() : nombre de rangées des zones
    ' H_zone : Hauteur interne zones [m]
    ' Larg_zone : Largeur zones [m]
    ' Ep_mur : épaisseur mur [m]
    ' Cond : conductivité mur
    ' Gap_B(): distance entre bandes
    '-----------------------------------------------------------------'
    
    Dim i As Integer
    Dim j As Integer
    Dim Ihaut As Boolean
    Dim temp1 As String
    Dim temp2 As String
    Dim temp3 As String
    Dim temp4 As String
    Dim temp5 As String
    Dim temp6 As String
    Dim temp7 As String
    Dim temp8 As String
    
    intOutFile = FreeFile
    Open strChemFich For Append As intOutFile
    
    For i = 1 To NbSect
    
        If (i = 2 Or i = 3 Or i = 6 Or i = 7 Or i = 10) Then Ihaut = True Else Ihaut = False
        
        For j = 1 To NbRangee(i)
         

            If j = 1 Then  'mur gauche
                           
                Print #intOutFile, "" 'Saut de ligne
                Print #intOutFile, "**branche Bmur_S" + CStr(i) + "_R" + CStr(j) + "Gauche mur_S" + CStr(i) + "_R" + CStr(j) + "Gauche_0 mur_S" + CStr(i) + "_R" + CStr(j) + "Gauche_1"
                Print #intOutFile, "*geom 0 "; Ep_mur; " rect "; H_zone; Larg_zone
                Print #intOutFile, "*prop brique "; cond; " 2000 400 273 "
                Print #intOutFile, "*maill 1 5 1"
                If (i = 1 Or i = 2 Or i = NbSect Or i = NbSect - 1) Then
                    Print #intOutFile, "*sortie conv Text 10+sfrad(T,Text)"
                End If
                Print #intOutFile, "*fin"
            End If
            
            If j = NbRangee(i) Then  'mur droit
    
                         
                Print #intOutFile, "" 'Saut de ligne
                Print #intOutFile, "**branche Bmur_S" + CStr(i) + "_R" + CStr(j) + "Droit mur_S" + CStr(i) + "_R" + CStr(j) + "Droit_0 mur_S" + CStr(i) + "_R" + CStr(j) + "Droit_1"
                Print #intOutFile, "*geom 0 "; Ep_mur; " rect "; H_zone; Larg_zone
                Print #intOutFile, "*prop brique "; cond; " 2000 400 273 "
                Print #intOutFile, "*maill 1 5 1"
                If (i = 1 Or i = 2 Or i = NbSect Or i = NbSect - 1) Then
                    Print #intOutFile, "*sortie conv Text 10+sfrad(T,Text)"
                End If
                Print #intOutFile, "*fin"
            End If
            
    ' mur haut ou bas

            If (Ihaut) Then
            
                Print #intOutFile, "" 'Saut de ligne
                Print #intOutFile, "**branche Bmur_S" + CStr(i) + "_R" + CStr(j) + "Haut mur_S" + CStr(i) + "_R" + CStr(j) + "Haut_0 mur_S" + CStr(i) + "_R" + CStr(j) + "Haut_1"
                Print #intOutFile, "*geom 0 "; Ep_mur; " rect "; Larg_zone; Gap_B(i)
                Print #intOutFile, "*prop brique "; cond; " 2000 400 273 "
                Print #intOutFile, "*maill 1 5 1"
                Print #intOutFile, "*sortie conv Text 10+sfrad(T,Text)"
                 Print #intOutFile, "*fin"
             
             Else
             
                Print #intOutFile, "" 'Saut de ligne
                Print #intOutFile, "**branche Bmur_S" + CStr(i) + "_R" + CStr(j) + "Bas mur_S" + CStr(i) + "_R" + CStr(j) + "Bas_0 mur_S" + CStr(i) + "_R" + CStr(j) + "Bas_1"
                Print #intOutFile, "*geom 0 "; Ep_mur; " rect "; Larg_zone; Gap_B(i)
                Print #intOutFile, "*prop brique "; cond; " 2000 400 273 "
                Print #intOutFile, "*maill 1 5 1"
                Print #intOutFile, "*sortie conv Text 10+sfrad(T,Text)"
                Print #intOutFile, "*fin"
             
             End If
             
              ' mur lat
              
            Print #intOutFile, "" 'Saut de ligne
            Print #intOutFile, "**branche Bmur_S" + CStr(i) + "_R" + CStr(j) + "Lat mur_S" + CStr(i) + "_R" + CStr(j) + "Lat_0 mur_S" + CStr(i) + "_R" + CStr(j) + "Lat_1"
            Print #intOutFile, "*geom 0 "; Ep_mur; " rect "; H_zone; 2 * Gap_B(i)
            Print #intOutFile, "*prop brique "; cond; " 2000 400 273 "
            Print #intOutFile, "*maill 1 5 1"
            Print #intOutFile, "*sortie conv Text 10+sfrad(T,Text)"
            Print #intOutFile, "*fin"

        Next j
    
    Next i
    
    Close #intOutFile
 
End Sub
Public Sub ThRes_Bandes(strChemFich As String, NbSect As Integer, NbRangee() As Integer, Larg_prod As Double, Ep_prod As Double, T_entreebande As Double, Acier As String, H_zone As Double, Prod As Double)

    '-----------------------------------------------------------------------------------'
    ' Ecriture de la bande dans le fichier réseau Thermette                   '
    '-----------------------------------------------------------------------------------'
    ' A. Ceccotti, 29/06/2006
    '-----------------------------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' NbSect : nombre de zones du four
    ' Nbrangee() : Longeurs des zones [m]
    ' larg_prod : largeur des produits [m]
    ' Ep_prod : épaisseur des produits [m]
    ' H_zone : Hauteur interne zones [m]
    ' Acier : type d'acier BISRA
    ' prod : débit massique
    '-----------------------------------------------------------------------------------'
    ' On décrit les bande en branches verticales dans lesquelles passe un débit d'acier
    '-----------------------------------------------------------------------------------'
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim temp1 As String
    Dim temp2 As String
    Dim temp3 As String
    Dim temp4 As String
    Dim temp5 As String
    Dim temp6 As String
    Dim temp_a As String
    Dim temp_b As String
    Dim temp_c As String

    
    intOutFile = FreeFile
    Open strChemFich For Append As intOutFile
    
    k = 0

For i = 1 To NbSect
    For j = 1 To NbRangee(i) - 1
        
            
        temp1 = "B_S" + CStr(i) + "_R" + CStr(j)
        

        Print #intOutFile, " " ' Saut de ligne
        Print #intOutFile, "**branche " + temp1 + " NS" + CStr(k) + " NS" + CStr(k + 1)
        Print #intOutFile, "*geom 0 "; H_zone; " rect "; Larg_prod; " "; Ep_prod
        ' Propriétés de l'acier
        Print #intOutFile, "*prop BISRA table(T,AcierCo) table(T,AcierRo) table(T,AcierCp) 293"
        Print #intOutFile, "*maill 1 6 1"
        
        Print #intOutFile, "*debmas prod"
        If (i = 1 And j = 1) Then
            Print #intOutFile, "*entree temp T_entreebande"
        End If
        Print #intOutFile, "*fin"
        k = k + 1

    Next j
Next i
    Worksheets(SResult).Range(CelNbStrip) = k - 1

    
    Close #intOutFile
End Sub

Public Sub ThRes_Tubes(strChemFich As String, NbSect As Integer, NbRangee() As Integer, a() As Double, b() As Double, c() As Double, d() As Double, Nbtube() As Double, Type_tube() As String)
    '---------------------------------------------------------------------------------------------------'
    ' Ecriture des Volumes tubes par zones et rangées dans le fichier réseau Thermette                   '
    '---------------------------------------------------------------------------------------------------'
    ' A. Ceccotti, 29/06/2006
    '---------------------------------------------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' NbSect : nombre de zones du four
    ' a,b,c,d : paramètres géométriques pour les tubes sup et inf
    ' type_tube : indique la forme du tube U ou W
    ' Nbtube : nombre de tube pour une rangée de la zone
    '-----------------------------------------------------------------------------------------------------'

Dim F(100), g(100) As Integer
Dim i, j, k, som As Integer
Dim temp1 As String
Dim temp2 As String
Dim temp3 As String
Dim temp4 As String

intOutFile = FreeFile
Open strChemFich For Append As intOutFile
   
For i = 1 To NbSect
    For j = 1 To NbRangee(i)
     
        Print #intOutFile, "" 'Saut de ligne
        temp1 = i
        temp2 = j
        temp4 = "Tube_S" + temp1 + "_R" + temp2
        Print #intOutFile, "**volume " + temp4
                    
        Select Case Type_tube(i, j)
            Case "U"
                Print #intOutFile, "*geom "; (a(i, j) * 3.1415 * (b(i, j) * b(i, j))) * 2 + (3.1415 * c(i, j) * 3.1415 * (b(i, j) * b(i, j))) * Nbtube(i, j); " "; ((3.1415 * b(i, j) * a(i, j) * 4) + (2 * 3.1415 * b(i, j) * 3.1415 * c(i, j))) * Nbtube(i, j)
            Case "W"
                Print #intOutFile, "*geom "; ((a(i, j) * 3.1415 * (b(i, j) * b(i, j))) * 2 + (3.1415 * c(i, j) * (3.1415 * (b(i, j) * b(i, j)))) * 3 + (d(i, j) * 3.1415 * (b(i, j) * b(i, j))) * 2) * Nbtube(i, j); " "; ((3.1415 * b(i, j) * a(i, j) * 4) + ((2 * 3.1415 * b(i, j) * 3.1415 * c(i, j)) * 3) + (3.1415 * b(i, j) * d(i, j) * 4)) * Nbtube(i, j)
            Case Else
                MsgBox "Attention, seuls les tubes U,W sont disponibles. Le tube U sera utilisé par défaut", vbInformation
                Print #intOutFile, "*geom "; (a(i, j) * 3.1415 * (b(i, j) * b(i, j))) * 2 + (3.1415 * c(i, j) * 3.1415 * (b(i, j) * b(i, j))) * Nbtube(i, j); " "; ((3.1415 * b(i, j) * a(i, j) * 4) + (2 * 3.1415 * b(i, j) * 3.1415 * c(i, j))) * Nbtube(i, j)
        End Select
        
        Print #intOutFile, "*prop Materiau_tube 51.6 10 100 273"
        Print #intOutFile, "*soll temp T_tube_Z" + temp1 + "_R" + temp2
        Print #intOutFile, "*fin"
   
    Next j
            
Next i
    
    Print #intOutFile, "" 'Saut de ligne
    
    Close #intOutFile

End Sub


Public Sub ThCalc(strChemFich As String, NbSect As Integer, NbRangee() As Integer, Nbtube() As Double)

    '---------------------------------------------------------------------------------------------------'
    ' Ecriture du fichier calcul thermette (définition des sorties)                                      '
    '---------------------------------------------------------------------------------------------------'
    ' P.Dubois
    '---------------------------------------------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' Nbrangee() : nombre de rangée de la zone
    ' NbSect : indice de la zone
    ' Nbtube() : nbre tube
    '---------------------------------------------------------------------------------------------------'
    intOutFile = FreeFile
    Open strChemFich For Append As intOutFile

    Dim i, j, k, Ns As Integer
    Dim temp1 As String
    
    Print #intOutFile, "*calcul statique"
    '=============
    ' nombre de sorties
    '=============
    ' bande
    For i = 1 To NbSect
        For j = 1 To NbRangee(i) - 1
            Ns = Ns + 2
        Next j
    Next i
    Ns = Ns + 2
    'mur et tube
    For i = 1 To NbSect

        For j = 1 To NbRangee(i)
            Ns = Ns + 2
        Next j
    Next i
    
    Print #intOutFile, "*sorties " + CStr(Ns)
    
    '=============
    ' sorties
    '=============
     ' bande
    Print #intOutFile, "S_Prod_IN EntM(BISRA,273,Tn(NS0))*prod 1 0"
    k = 1
    For i = 1 To NbSect
        For j = 1 To NbRangee(i) - 1
                
            temp1 = "SStrip_S" + CStr(i) + "_R" + CStr(j)
            Print #intOutFile, temp1 + " Tn(NS" + CStr(k) + ") 1 0 "
            
            temp1 = "SFluxStrip_S" + CStr(i) + "_R" + CStr(j)
            Print #intOutFile, temp1 + " EntM(BISRA,Tn(NS" + CStr(k - 1) + "),Tn(NS" + CStr(k) + "))*prod 1 0 "
            k = k + 1

        Next j
            
    Next i
    
    Print #intOutFile, "S_Prod_OUT EntM(BISRA,273,Tn(NS" + CStr(k - 1) + "))*prod 1 0"

    ' murs

      For i = 1 To NbSect
    
        If (i = 2 Or i = 3 Or i = 6 Or i = 7 Or i = 10) Then Ihaut = True Else Ihaut = False
        
        For j = 1 To NbRangee(i)
         
            temp1 = "SFluxMur_S" + CStr(i) + "_R" + CStr(j) + " "
            If j = 1 Then  'mur gauche
                temp1 = temp1 + "Fcond(Bmur_S" + CStr(i) + "_R" + CStr(j) + "Gauche,0)+"
            End If
            
            If j = NbRangee(i) Then  'mur droit
                temp1 = temp1 + "Fcond(Bmur_S" + CStr(i) + "_R" + CStr(j) + "Droit,0)+"
            End If
            
    ' mur haut ou bas

            If (Ihaut) Then
                 temp1 = temp1 + "Fcond(Bmur_S" + CStr(i) + "_R" + CStr(j) + "Haut,0)+"
             Else
                 temp1 = temp1 + "Fcond(Bmur_S" + CStr(i) + "_R" + CStr(j) + "Bas,0)+"
             End If
             
              ' mur lat
              Print #intOutFile, temp1 + "Fcond(Bmur_S" + CStr(i) + "_R" + CStr(j) + "Lat,0) 1 0"

        Next j
    
    Next i
    
    
    ' Tubes



      For i = 1 To NbSect
    
              
        For j = 1 To NbRangee(i)
         
            temp1 = "SFluxTube_S" + CStr(i) + "_R" + CStr(j) + " "
            Print #intOutFile, temp1 + "Fvnet(Tube_S" + CStr(i) + "_R" + CStr(j) + ") 1 0"
            
        Next j
    
    Next i
  
  

'---------------------------------------------------------------------------------------------------------------------------'
    Close #intOutFile
    
End Sub

