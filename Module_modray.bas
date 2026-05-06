Attribute VB_Name = "Module_modray"
Public Sub ModrayVrtf()

Dim strChemFich As String   ' Fichiers thermette
Dim intOutFile As Integer
Dim i As Integer
Dim j As Integer
Dim TotaleRangee As Integer 'somme rangee

Dim izone As Integer
Dim jrangee As Integer
Dim temp1 As String
Dim temp2 As String
Dim temp3 As String
Dim temp4 As String
Dim temp5 As String
Dim temp6 As String
Dim Reponse As Integer 'message box

'--------------------------'
'Données Produits  '
'--------------------------'
Dim Ep_prod As Double    'épaisseur produit
Dim Larg_prod As Double  'largeur produit


'--------------------------'
'Données Zones  '
'--------------------------'
Dim NbSect As Integer     'nombre de sections
Dim NbRangee() As Integer    'nombre de rangee
Dim H_zone As Double       'hauteur des zones
Dim Larg_Four As Double    'largeur des zones
Dim Gap_B() As Double      'écartement entre bandes d'une même zone


'--------------------------'
'Description rangees  '
'--------------------------'
Dim Type_tube() As String
Dim a() As Double
Dim b() As Double
Dim c() As Double
Dim d() As Double
Dim Nbtube() As Double
Dim Prem_Abs() As Double
Dim Distance_BT() As Double
Dim Pas_tube() As Double
Dim emis() As Double       'emissivité bande pour indice som2


'-------------------------------
'initialisation
'-------------------------------
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
    ReDim Prem_Abs(NbSect, 10)
    ReDim Distance_BT(NbSect, 10)
    ReDim Pas_tube(NbSect, 10)
    ReDim emis(NbSect, 10)

    ChemThermette = Worksheets(SFiles).Range(CThermette)
    ChemModray = Worksheets(SFiles).Range(CModray)
        With Worksheets(SGeom)
            For i = 1 To NbSect
                NbRangee(i) = .Range(CelSectDesc).Offset(i, 0).Value
            Next i


'--------------------------'
'affectation des valeurs aux variables           '
'--------------------------'

            Ep_prod = .Range(CelEpProd).Value * 0.001
            Larg_prod = .Range(CelLargProd).Value
            Acier = .Range(CelTypeProd).Value
            H_zone = .Range(CelHautSect).Value
            Larg_Four = .Range(CelLargSect).Value
        

            For i = 1 To NbSect
                Gap_B(i) = .Range(CelDiamRoul).Offset(i, 0).Value
            Next i
            k = 1
            For i = 1 To NbSect
                For j = 1 To NbRangee(i)
        
                    Type_tube(i, j) = .Range(CelTubeType).Offset(k, 0).Value
                    a(i, j) = .Range(CelA).Offset(k, 0).Value
                    b(i, j) = .Range(CelB).Offset(k, 0).Value / 2
                    c(i, j) = .Range(CelC).Offset(k, 0).Value
                    d(i, j) = .Range(CelD).Offset(k, 0).Value
                    Prem_Abs(i, j) = .Range(CelPreAbs).Offset(k, 0).Value
                    Distance_BT(i, j) = .Range(CelDiamRoul).Offset(i, 0).Value / 2
                    Pas_tube(i, j) = .Range(CelTubePitch).Offset(k, 0).Value
                    Nbtube(i, j) = .Range(CelNBTube).Offset(k, 0).Value
                    emis(i, j) = .Range(CelTubeEmis).Offset(k, 0).Value
                    k = k + 1
                Next j
     
        Next i
    End With

'------------------------------------------------'
' Etape 0 : Vérification des données             '
'------------------------------------------------'
    Dim validite As Integer
    Dim X As Integer
    Dim y As Integer
    X = 0
    y = 0
    validite = 0

    If Larg_prod >= Larg_Four Then
    validite = 1
    MsgBox "Error, check strip width", vbInformation
    End If
       
    If Larg_zone >= H_zone Then
    validite = 1
    MsgBox "Error, check section width", vbInformation
    End If


    If Round(Ep_prod, 3) > 0.007 Then
    validite = 1
    MsgBox "Error, strip thickness must be lower than 7 mm", vbInformation
    End If

       
    For i = 1 To NbSect
    If Round(Gap_B(i), 2) > 0.8 Then X = 1
    Next i
    If X = 1 Then MsgBox "Be carefull, roll diameter is normally close to 0.8 m", vbInformation
    
    X = 0
    For i = 1 To NbSect Step 2
    If NbRangee(i) <> NbRangee(i + 1) Then X = 1
    If Gap_B(i) <> Gap_B(i + 1) Then y = 1
    Next i
    If X = 1 Or y = 1 Then
    validite = 1
    MsgBox "Error row number or roll diameter", vbInformation
    End If


    If validite = 0 Then

      k = 2

'--------------------------------------------------------'
' Etape 1 : Construction des fichier Modray              '
'--------------------------------------------------------'

'écriture des parois pour chaque rangée de chaque zone
      
        For i = 1 To NbSect
                
            For j = 1 To NbRangee(i)
                jrangee = j
                temp2 = j

                          
                strChemFich = ChemModray + "\VRTF_S"
                intOutFile = FreeFile
                Open strChemFich For Output As intOutFile
                Close #intOutFile
        
                Call MdRes_objet(strChemFich, NbRangee(i), i, j, Nbtube(i, j), Type_tube(i, j))
                Call MdRes_parois(strChemFich, NbRangee(i), i, j, H_zone, Larg_Four, Gap_B(i))
                Call MdRes_bande(strChemFich, NbRangee(i), i, j, H_zone, Larg_Four, Gap_B(i), Ep_prod, Larg_prod, emis(i, j))
                Call MdRes_tube(strChemFich, Larg_Four, NbRangee(i), i, j, Nbtube(i, j), Type_tube(i, j), Prem_Abs(i, j), Distance_BT(i, j), Pas_tube(i, j), a(i, j), b(i, j), c(i, j), d(i, j))
                Call MdRes_groupe(strChemFich, NbRangee(i), i, j, Nbtube(i, j), Type_tube(i, j))
                
'-----------------------------------------'
' Etape 2 : Lancement du calcul modray1   '
'-----------------------------------------'
                Call GetCurrentDirectory(lstr, strCurrentDir)
                SetCurrentDirectory (Worksheets(SFiles).Range(CModray) + "\")
                Call ShellWait(ChemModray + "\Modray1.exe " + strChemFich + " " + strChemFich + "_m1")
'-----------------------------------------'
' Etape 2 : Lancement du calcul modray2   '
'-----------------------------------------'
                Call ShellWait(ChemModray + "\Modray2.exe " + strChemFich + "_m1")
                        intFile = FreeFile
'-----------------------------------------'
' Etape 3 : récupération des donnees   '
'-----------------------------------------'
                    '---- 1ere étape : on récupère le texte existant dans le fichier
                    Open ChemModray + "\Res_VRTF_R" For Input As intFile
                

                    While Not EOF(intFile)
                       
                        Line Input #intFile, text1
                        Worksheets(SMesh).Cells(k, 1) = text1
                        k = k + 1
                       
                    Wend
                     Close #intFile
                Next j
        Next i

        Worksheets(SMesh).Cells(1, 1) = k - 2
End If
SetCurrentDirectory (strCurrentDir)
End With
End Sub
Public Sub MdRes_objet(strChemFich As String, NbRangee As Integer, Isect As Integer, Irangee As Integer, Nbtube As Double, Type_tube As String)

    '---------------------------------------------------------------------------------------------------'
    ' Ecriture début fichier Modray1  (nombre d'objets ...)                                      '
    '---------------------------------------------------------------------------------------------------'
    ' A. Ceccotti, 20/03/2006
    '---------------------------------------------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' Nbrangee :indice de la rangée
    ' izone : indice de la zone
    ' Nbtube :
    ' type_tube :
    '---------------------------------------------------------------------------------------------------'
    intOutFile = FreeFile
    Open strChemFich For Append As intOutFile

    Dim i As Integer
    Dim Nobj As Integer       'nbre d'objets
    Dim temp1 As String
    Dim temp2 As String
    
    temp1 = izone
    temp2 = jrangee
    
    Print #intOutFile, "**Nbandes 1"
    Print #intOutFile, "0 100"
    Print #intOutFile, "" 'Saut de ligne
    
    Print #intOutFile, "**Calculs groupe 1"
    Print #intOutFile, "1 diffus Res_VRTF_R"
    Print #intOutFile, "**Ordre_rotations XYZ"
    Print #intOutFile, "" 'Saut de ligne
    
    
    If Type_tube = "I" Then
    ElseIf Type_tube = "U" Then Nobj = 3 * Nbtube
    ElseIf Type_tube = "W" Then Nobj = 7 * Nbtube
    End If
    ' objet bande et mur
     If Irangee = 1 Or Irangee = NbRangee Then Nobj = Nobj + 7 Else Nobj = Nobj + 9
    
    Print #intOutFile, "**nobjets "; Nobj
    Print #intOutFile, "" 'Saut de ligne
    
              
    Close #intOutFile

End Sub

Public Sub MdRes_parois(strChemFich As String, NbRangee As Integer, Isect As Integer, jrangee As Integer, H_zone As Double, Larg_zone As Double, Gap_B As Double)

    '---------------------------------------------------------------------------------------------------'
    ' Ecriture dans le premier fichier Modray objets parois (définition géométrique)                                      '
    '---------------------------------------------------------------------------------------------------'
    ' A. Ceccotti, 29/06/2006
    '---------------------------------------------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' Nbrangee() : nombre de rangée de la zone
    ' izone : indice de la zone
    ' jrangee : indice de la rangée de la zone
    ' H_zone : hauteur zone
    ' Larg_zone : largeur four
    ' Gap_B(): espace entre bande d'une même rangée
    '---------------------------------------------------------------------------------------------------'
    intOutFile = FreeFile
    Open strChemFich For Append As intOutFile
       
    Dim Ihaut As Boolean
    Dim temp1 As String
    Dim temp2 As String
    Dim temp3 As String
    temp1 = Isect
    temp2 = jrangee
    
   
    If (Isect = 2 Or Isect = 3 Or Isect = 6 Or Isect = 7 Or Isect = 10) Then Ihaut = True Else Ihaut = False
    If jrangee = 1 Then  'première rangée de la zone
                
        temp3 = "mur_S" + temp1 + "_R" + temp2 + "Gauche_0" 'mur coté gauche
        
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet " + temp3
        Print #intOutFile, "*geom 0 0 0 0 0 0 paral "; Gap_B; Larg_zone; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        Print #intOutFile, "*proprad 0.85 0"
        Print #intOutFile, "*enceinte"
        Print #intOutFile, "*seulrayon gauche"
        Print #intOutFile, "*noeudthermette 0"
        Print #intOutFile, "*fin"
                
    ElseIf jrangee = NbRangee Then  'première rangée de la zone
                
        temp3 = "mur_S" + temp1 + "_R" + temp2 + "Droit_0" 'mur coté droit
        
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet " + temp3
        Print #intOutFile, "*geom 0 0 0 0 0 0 paral "; Gap_B; Larg_zone; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        Print #intOutFile, "*proprad 0.85 0"
        Print #intOutFile, "*enceinte"
        Print #intOutFile, "*seulrayon droite"
        Print #intOutFile, "*noeudthermette 0"
        Print #intOutFile, "*fin"
    
    End If
    'murs latéraux
        temp3 = "mur_S" + temp1 + "_R" + temp2 + "Lat_0" 'mur coté gauche
        
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet " + temp3
        Print #intOutFile, "*geom 0 0 0 0 0 0 paral "; Gap_B; Larg_zone; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        Print #intOutFile, "*proprad 0.85 0"
        Print #intOutFile, "*enceinte"
        Print #intOutFile, "*norayon droite"
        Print #intOutFile, "*norayon gauche"
        Print #intOutFile, "*norayon base"
        Print #intOutFile, "*norayon sommet"
        Print #intOutFile, "*noeudthermette 0"
        Print #intOutFile, "*fin"

                
    'la section est supérieure le mur est en haut la séparation en bas
   
        temp3 = "mur_S" + temp1 + "_R" + temp2 + "Haut_0" 'mur coté gauche
        
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet " + temp3
        Print #intOutFile, "*geom 0 0 0 0 0 0 paral "; Gap_B; Larg_zone; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        If (Ihaut) Then
            Print #intOutFile, "*proprad 0.85 0"
        Else
            Print #intOutFile, "*proprad 0 0"
        End If
        Print #intOutFile, "*enceinte"
        Print #intOutFile, "*seulrayon sommet"
        Print #intOutFile, "*noeudthermette 0"
        Print #intOutFile, "*fin"
        
        temp3 = "mur_S" + temp1 + "_R" + temp2 + "Bas_0" 'mur coté gauche
        
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet " + temp3
        Print #intOutFile, "*geom 0 0 0 0 0 0 paral "; Gap_B; Larg_zone; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        If (Ihaut) Then
            Print #intOutFile, "*proprad 0 0"
        Else
            Print #intOutFile, "*proprad 0.85 0"
        End If
        Print #intOutFile, "*enceinte"
        Print #intOutFile, "*seulrayon base"
        Print #intOutFile, "*noeudthermette 0"
        Print #intOutFile, "*fin"
        
    
 Close #intOutFile
        
End Sub

Public Sub MdRes_bande(strChemFich As String, NbRangee As Integer, Isect As Integer, jrangee As Integer, H_zone As Double, Larg_Four As Double, Gap_B As Double, Ep_prod As Double, Larg_prod As Double, emis As Double)

    '---------------------------------------------------------------------------------------------------'
    ' Ecriture dans le premier fichier Modray objets bandes (définition géométrique)                                      '
    '---------------------------------------------------------------------------------------------------'
    ' A. Ceccotti, 29/06/2006
    '---------------------------------------------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' Nbrangee() : nombre de rangée de la zone
    ' izone : indice de la zone
    ' jrangee : indice de la rangée de la zone
    ' H_zone : hauteur zone
    ' Larg_zone : largeur four
    ' Gap_B(): espace entre bande d'une même rangée
    ' Ep_prod : épaisseur bande
    '---------------------------------------------------------------------------------------------------'
    intOutFile = FreeFile
    Open strChemFich For Append As intOutFile
       
    Dim Delta As Double
    Dim temp1 As String
    Dim temp2 As String
    Dim temp3 As String
    Dim temp4 As String
    Dim temp5 As String
    Dim temp6 As String
    
    temp1 = Isect
    temp2 = jrangee - 1
    temp3 = jrangee
    temp5 = jrangee - 1 + som3
    Delta = (Larg_Four - Larg_prod) / 2
    If jrangee > 1 Then
            
        temp4 = "B_S" + temp1 + "_R" + temp2
        
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet "; temp4
        Print #intOutFile, "*geom 0 "; (Larg_Four - Larg_prod) / 2; 0; " 0 0 0 paral "; Ep_prod; Larg_prod; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        Print #intOutFile, "*proprad "; emis; " 0"
        Print #intOutFile, "*seulrayon droite"
        Print #intOutFile, "*branchethermette "; Larg_prod * H_zone
        Print #intOutFile, "*fin"

        temp4 = "B_S" + temp1 + "_R" + temp2 + "_Sepa"
        
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet "; temp4
        Print #intOutFile, "*geom 0 0 0 0 0 0 paral "; Ep_prod; Delta; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        Print #intOutFile, "*proprad 0 0"
        Print #intOutFile, "*seulrayon droite"
        Print #intOutFile, "*fin"
          
        temp4 = "B_S" + temp1 + "_R" + temp2 + "_Sepa2"
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet "; temp4
        Print #intOutFile, "*geom 0 "; Larg_prod + Delta; " 0 0 0 0 paral "; Ep_prod; Delta; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        Print #intOutFile, "*proprad 0 0"
        Print #intOutFile, "*seulrayon droite"
        Print #intOutFile, "*fin"

    End If

    If jrangee < NbRangee Then
            
        temp4 = "B_S" + temp1 + "_R" + temp3
        
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet "; temp4
        Print #intOutFile, "*geom  "; Gap_B; (Larg_Four - Larg_prod) / 2; 0; " 0 0 0 paral "; Ep_prod; Larg_prod; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        Print #intOutFile, "*proprad "; emis; " 0"
        Print #intOutFile, "*seulrayon gauche"
        Print #intOutFile, "*branchethermette "; Larg_prod * H_zone
        Print #intOutFile, "*fin"

        temp4 = "B_S" + temp1 + "_R" + temp3 + "_Sepa"
        
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet "; temp4
        Print #intOutFile, "*geom "; Gap_B; " 0 0 0 0 0 paral "; Ep_prod; Delta; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        Print #intOutFile, "*proprad 0 0"
        Print #intOutFile, "*seulrayon gauche"
        Print #intOutFile, "*fin"
          
        temp4 = "B_S" + temp1 + "_R" + temp2 + "_Sepa2"
        Print #intOutFile, "" 'Saut de ligne
        Print #intOutFile, "**objet "; temp4
        Print #intOutFile, "*geom  "; Gap_B; Larg_prod + Delta; " 0 0 0 0 paral "; Ep_prod; Delta; H_zone
        Print #intOutFile, "*maill  1 5 1 5 1 5 1"
        Print #intOutFile, "*proprad 0 0"
        Print #intOutFile, "*seulrayon gauche"
        Print #intOutFile, "*fin"

    End If
 
    
 Close #intOutFile
        
End Sub

Public Sub MdRes_tube(strChemFich As String, Larg_Four As Double, NbRangee As Integer, Isect As Integer, Irangee As Integer, Nbtube As Double, Type_tube As String, Prem_Abs As Double, Distance_BT As Double, Pas_tube As Double, a As Double, b As Double, c As Double, d As Double)
    
    '---------------------------------------------------------------------------------------------------'
    ' Ecriture dans le premier fichier Modray objets bandes (définition géométrique)                                      '
    '---------------------------------------------------------------------------------------------------'
    ' A. Ceccotti, 29/06/2006
    '---------------------------------------------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' izone : indice de la zone
    ' jrangee : indice rangee
    ' type_tube : type des tubes U,W
    ' a : paramètre géométrique pour les tubes
    ' b : paramètre géométrique pour les tubes
    ' c : paramètre géométrique pour les tubes
    ' d : paramètre géométrique pour les tubes
    ' Nbtube : nombre de tubes sup
    ' Prem_Abs : abscisse vertical du premier tube dans la rangee de la zone
    ' Distance_BT : distance bande tube dans la rangee de la zone
    ' Pas_tube: pas entre les tubes dans la rangee de la zone
    '---------------------------------------------------------------------------------------------------'
intOutFile = FreeFile
Open strChemFich For Append As intOutFile

 Dim temp1 As String
    Dim temp2 As String
    Dim temp3 As String
    Dim temp4 As String
    Dim temp5 As String
    Dim temp6 As String
    
    temp1 = Isect
    temp2 = Irangee

    For i = 1 To Nbtube
        
        temp3 = i
        Select Case Type_tube
                
            Case "U"
                If (Int(i / 2) = i / 2) Then        ' pour inverser les tubes
                   
                    Print #intOutFile, "**objet Tubea" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                    Print #intOutFile, "*geom "; Distance_BT; "0.01"; (Prem_Abs - c) + ((i - 1) * Pas_tube); " -90 0 0 cylin "; b; a; "360"
                    Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                    Print #intOutFile, "*proprad 0.85 0"
                    Print #intOutFile, "*norayon sommet"
                    Print #intOutFile, "*fin"
                    Print #intOutFile, "" 'Saut de ligne
                    
                    Print #intOutFile, "**objet Tubeb" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                    Print #intOutFile, "*geom "; Distance_BT; a + 0.01; (Prem_Abs - c) + ((i - 1) * Pas_tube); " -90 90 0 coude "; b; c; "180 360"
                    Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                    Print #intOutFile, "*proprad 0.85 0"
                    Print #intOutFile, "*norayon base"
                    Print #intOutFile, "*norayon sommet"
                    Print #intOutFile, "*fin"
                    Print #intOutFile, "" 'Saut de ligne
                    
                    Print #intOutFile, "**objet Tubec" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                    Print #intOutFile, "*geom "; Distance_BT; "0.01"; (Prem_Abs + c) + ((i - 1) * Pas_tube); " -90 0 0 cylin "; b; a; "360"
                    Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                    Print #intOutFile, "*proprad 0.85 0"
                    Print #intOutFile, "*norayon sommet"
                    Print #intOutFile, "*fin"
                    Print #intOutFile, "" 'Saut de ligne
                Else
                    
                    Print #intOutFile, "**objet Tubea" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                    Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01; (Prem_Abs - c) + ((i - 1) * Pas_tube); " 90 0 0 cylin "; b; a; "360"
                    Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                    Print #intOutFile, "*proprad 0.85 0"
                    Print #intOutFile, "*norayon sommet"
                    Print #intOutFile, "*fin"
                    Print #intOutFile, "" 'Saut de ligne
                    
                    Print #intOutFile, "**objet Tubeb" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                    Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01 - a; (Prem_Abs - c) + ((i - 1) * Pas_tube); " 90 90 0 coude "; b; c; "180 360"
                    Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                    Print #intOutFile, "*proprad 0.85 0"
                    Print #intOutFile, "*norayon base"
                    Print #intOutFile, "*norayon sommet"
                    Print #intOutFile, "*fin"
                    Print #intOutFile, "" 'Saut de ligne
                    
                    Print #intOutFile, "**objet Tubec" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                    Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01; (Prem_Abs + c) + ((i - 1) * Pas_tube); " 90 0 0 cylin "; b; a; "360"
                    Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                    Print #intOutFile, "*proprad 0.85 0"
                    Print #intOutFile, "*norayon sommet"
                    Print #intOutFile, "*fin"
                    Print #intOutFile, "" 'Saut de ligne
                End If
                
            Case "W"
                If (Int(i / 2) = i / 2) Then        ' pour inverser les tubes
                     Print #intOutFile, "" 'Saut de ligne
                     Print #intOutFile, "**objet Tubea" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; "0.01"; (Prem_Abs - (3 * c)) + ((i - 1) * Pas_tube); "-90 0 0 cylin "; b; a; "360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                    
                     Print #intOutFile, "**objet Tubeb" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; a + 0.01; (Prem_Abs - (3 * c)) + ((i - 1) * Pas_tube); "-90 90 0 coude "; b; c; "180 360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                     
                     Print #intOutFile, "**objet Tubec" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; a + 0.01 - d; (Prem_Abs - c) + ((i - 1) * Pas_tube); "-90 0 0 cylin "; b; d; "360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                     
                     Print #intOutFile, "**objet Tubed" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; a + 0.01 - d; (Prem_Abs - c) + ((i - 1) * Pas_tube); " -90 -270 180 coude "; b; c; "180 360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                    
                     Print #intOutFile, "**objet Tubee" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; a + 0.01 - d; (Prem_Abs + c) + ((i - 1) * Pas_tube); "-90 0 0 cylin "; b; d; "360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                    
                     Print #intOutFile, "**objet Tubef" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; a + 0.01; (Prem_Abs + c) + ((i - 1) * Pas_tube); "-90 90 0 coude "; b; c; "180 360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                     
                     Print #intOutFile, "**objet Tubeg" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; "0.01"; (Prem_Abs + (3 * c)) + ((i - 1) * Pas_tube); "-90 0 0 cylin "; b; a; "360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                Else
                    Print #intOutFile, "" 'Saut de ligne
                     Print #intOutFile, "**objet Tubea" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01; (Prem_Abs - (3 * c)) + ((i - 1) * Pas_tube); "90 0 0 cylin "; b; a; "360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                    
                     Print #intOutFile, "**objet Tubeb" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01 - a; (Prem_Abs - (3 * c)) + ((i - 1) * Pas_tube); "90 90 0 coude "; b; c; "180 360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                     
                     Print #intOutFile, "**objet Tubec" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01 - a + d; (Prem_Abs - c) + ((i - 1) * Pas_tube); "90 0 0 cylin "; b; d; "360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                     
                     Print #intOutFile, "**objet Tubed" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01 - a + d; (Prem_Abs - c) + ((i - 1) * Pas_tube); " 90 -270 180 coude "; b; c; "180 360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                    
                     Print #intOutFile, "**objet Tubee" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01 - a + d; (Prem_Abs + c) + ((i - 1) * Pas_tube); "90 0 0 cylin "; b; d; "360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                    
                     Print #intOutFile, "**objet Tubef" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01 - a; (Prem_Abs + c) + ((i - 1) * Pas_tube); "90 90 0 coude "; b; c; "180 360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon base"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                     
                     Print #intOutFile, "**objet Tubeg" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     Print #intOutFile, "*geom "; Distance_BT; Larg_Four - 0.01; (Prem_Abs + (3 * c)) + ((i - 1) * Pas_tube); "90 0 0 cylin "; b; a; "360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                End If
                     
                Case "I"
                    Print #intOutFile, "" 'Saut de ligne
                    Print #intOutFile, "**objet Tubea" + "_S" + temp1 + "_R" + temp2 + "_N" + temp3
                     If i = 1 Then Print #intOutFile, "*geom "; "0.01"; Distance_BT; Prem_Abs - c; 0; 90; " 0 cylin "; b; a - 0.01; "360"
                     If i <> 1 Then Print #intOutFile, "*geom "; "0.01"; Distance_BT; (Prem_Abs - c) + ((i - 1) * Pas_tube); 0; 90; " 0 cylin "; b; a - 0.01; "360"
                     Print #intOutFile, "*maill  1 8 1 1 1 3 1"
                     Print #intOutFile, "*proprad 0.85 0"
                     Print #intOutFile, "*norayon sommet"
                     Print #intOutFile, "*fin"
                     Print #intOutFile, "" 'Saut de ligne
                    
        
                
                
        End Select
            
    Next i





Close #intOutFile

End Sub

Public Sub MdRes_groupe(strChemFich As String, NbRangee As Integer, Isect As Integer, Irangee As Integer, Nbtube As Double, Type_tube As String)

    '---------------------------------------------------------------------------------------------------'
    ' Ecriture des groupes fichier Modray1                                        '
    '---------------------------------------------------------------------------------------------------'
    ' A. Ceccotti, 20/03/2006
    '---------------------------------------------------------------------------------------------------'
    ' strChemFich : chemin du fichier Réseau
    ' Nbrangee :indice de la rangée
    ' izone : indice de la zone
    ' Nbtube :
    ' type_tube :
    '---------------------------------------------------------------------------------------------------'

    Dim i As Integer
    Dim Ngpe As Integer       'nbre groupe
    Dim temp1 As String
    Dim temp2 As String
    Dim temp3 As String
    Dim temp4 As String
    Dim temp5 As String
    Dim temp6 As String
    Dim temp7 As String
    Dim temp8 As String
    Dim temp9 As String
    Dim temp10 As String
    
    
    temp1 = Isect
    temp2 = Irangee

    If Nbtube > 0 Then
       intOutFile = FreeFile
        Open strChemFich For Append As intOutFile

         Print #intOutFile, "**Groupes "; 1
         Print #intOutFile, "" 'Saut de ligne
         If Type_tube = "W" Then
             temp10 = ""
             For i = 1 To Nbtube
                 temp3 = i
                 temp4 = "Tubea_S" + temp1 + "_R" + temp2 + "_N" + temp3
                 temp5 = "Tubeb_S" + temp1 + "_R" + temp2 + "_N" + temp3 + " " + temp4
                 temp6 = "Tubec_S" + temp1 + "_R" + temp2 + "_N" + temp3 + " " + temp5
                 temp7 = "Tubed_S" + temp1 + "_R" + temp2 + "_N" + temp3 + " " + temp6
                 temp8 = "Tubee_S" + temp1 + "_R" + temp2 + "_N" + temp3 + " " + temp7
                 temp9 = "Tubef_S" + temp1 + "_R" + temp2 + "_N" + temp3 + " " + temp8
                 temp10 = temp10 + " " + "Tubeg_S" + temp1 + "_R" + temp2 + "_N" + temp3 + " " + temp9
             Next i
                 
             Print #intOutFile, "Tube_S" + temp1 + "_R" + temp2; 7 * Nbtube; temp10
             Print #intOutFile, "" 'Saut de ligne
                 
                             
         ElseIf Type_tube = "U" Then
             temp6 = ""
             For i = 1 To Nbtube
                 temp3 = i
                 temp4 = "Tubea_S" + temp1 + "_R" + temp2 + "_N" + temp3 + " "
                 temp5 = "Tubeb_S" + temp1 + "_R" + temp2 + "_N" + temp3 + " " + temp4
                 temp6 = temp6 + " " + "Tubec_S" + temp1 + "_R" + temp2 + "_N" + temp3 + " " + temp5
              Next i
                     Dim n As Integer
                     Dim s As String
                     Dim str As String
                     
                     Print #intOutFile, "Tube_S" + temp1 + "_R" + temp2; 3 * Nbtube; temp6
                     Print #intOutFile, "" 'Saut de ligne
                 
         End If
        
                 
         Close #intOutFile
    End If


End Sub
