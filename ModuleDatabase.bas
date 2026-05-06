Attribute VB_Name = "ModuleDatabase"

'----------------------------------------
' Refractories Database Functions
'----------------------------------------


Public Function Refrac_GetDensity(RefracName As String)

    '----------------------------------------------------------------
    ' Look-up for Density of Refractory density in Database sheet
    '----------------------------------------------------------------
    ' L. Ferrand, 01/2009
    '----------------------------------------------------------------

    Dim i As Integer

    With Worksheets(SRefracRo)

        i = 1
        While RefracName <> .Range(CelDataOrig).Offset(0, i) And i <= .Range(CelDataOrig)
            i = i + 1
        Wend
        If i > .Range(CelDataOrig) Then
            Call MsgBox("The material " & RefracName & " was not found in the database. Density = 0", vbExclamation, Fox)
            Refrac_GetDensity = 0

            Exit Function
        End If
        Refrac_GetDensity = .Range(CelDataOrig).Offset(1, i)

    End With

End Function


Public Function Refrac_GetCp(RefracName As String)

    '----------------------------------------------------------------
    ' Look-up for Density of Refractory density in Database sheet
    '----------------------------------------------------------------
    ' L. Ferrand, 01/2009
    '----------------------------------------------------------------

    Dim i As Integer

    With Worksheets(SRefracCp)

        i = 1
        While RefracName <> .Range(CelDataOrig).Offset(0, i) And i <= .Range(CelDataOrig)
            i = i + 1
        Wend
        If i > .Range(CelDataOrig) Then
            Call MsgBox("The material " & RefracName & " was not found in the database. Density = 0", vbExclamation, Fox)
            Refrac_GetCp = 0

            Exit Function
        End If
        Refrac_GetCp = .Range(CelDataOrig).Offset(1, i)

    End With

End Function


Sub Refrac_GetCoTable(RefracName As String, nval As Integer, vecT() As Double, vecCo() As Double)
    '----------------------------------------------------------------
    ' Look-up for Conductivity of Refractory in Database sheet
    '----------------------------------------------------------------
    ' L. Ferrand, 01/2009
    '----------------------------------------------------------------

    Dim i As Integer

    With Worksheets(SRefracCo)

        i = 1
        While RefracName <> .Range(CelDataOrig).Offset(0, i) And i <= .Range(CelDataOrig)
            i = i + 1
        Wend
        If i > .Range(CelDataOrig) Then
            Call MsgBox("The material " & RefracName & " was not found in the database.", vbExclamation, Fox)

            Exit Sub
        End If

        nval = .Range(CelDataOrig).Offset(-1, 0)

        ReDim vecT(nval)
        ReDim vecCo(nval)

        For j = 1 To nval
            vecT(j) = .Range(CelDataOrig).Offset(j, 0)
            vecCo(j) = .Range(CelDataOrig).Offset(j, i)

        Next j

    End With

End Sub


Sub Refrac_GetEmissTable(RefracName As String, nval As Integer, vecT() As Double, vecEmiss() As Double)
    '----------------------------------------------------------------
    ' Look-up for Emissivity of Refractory in Database sheet
    '----------------------------------------------------------------
    ' L. Ferrand, 01/2009
    '----------------------------------------------------------------

    Dim i As Integer

    With Worksheets(SRefracEmiss)

        i = 1
        While RefracName <> .Range(CelDataOrig).Offset(0, i) And i <= .Range(CelDataOrig)
            i = i + 1
        Wend
        If i > .Range(CelDataOrig) Then
            Call MsgBox("The material " & RefracName & " was not found in the database.", vbExclamation, Fox)

            Exit Sub
        End If

        nval = .Range(CelDataOrig).Offset(-1, 0)

        ReDim vecT(nval)
        ReDim vecEmiss(nval)

        For j = 1 To nval
            vecT(j) = .Range(CelDataOrig).Offset(j, 0)
            vecEmiss(j) = .Range(CelDataOrig).Offset(j, i)

        Next j

    End With

End Sub


Function Refrac_GetEmiss(RefracName As String, TKelvin As Double)
    '----------------------------------------------------------------
    ' Look-up for Emissivity of Refractory in Database sheet
    '----------------------------------------------------------------
    ' L. Ferrand, 01/2009
    '----------------------------------------------------------------

    Dim i As Integer
    Dim j As Integer


    With Worksheets(SRefracEmiss)

        i = 1
        While RefracName <> .Range(CelDataOrig).Offset(0, i) And i <= .Range(CelDataOrig)
            i = i + 1
        Wend
        If i > .Range(CelDataOrig) Then
            Call MsgBox("The material " & RefracName & " was not found in the database.", vbExclamation, Fox)

            Exit Function
        End If

        ' Look-up for temperature
        j = 1

        While TKelvin > .Range(CelDataOrig).Offset(j, 0) And j <= .Range(CelDataOrig).Offset(-1, 0)
            j = j + 1
        Wend

        Refrac_GetEmiss = Interpolate(.Range(CelDataOrig).Offset(j, 0), .Range(CelDataOrig).Offset(j + 1, 0), .Range(CelDataOrig).Offset(j, i), .Range(CelDataOrig).Offset(j + 1, i), TKelvin)

    End With

End Function


'----------------------------------------
' Fuel/Comburant Grades Database Functions
'----------------------------------------


'Public Sub Gaz_GetCompo(SBasdo As String, Name As String, ByRef NbPureGaz As Integer, ByRef PureGazNames() As String, PureGazCompo() As Double)
'
'
'    '----------------------------------------------------------------
'    ' Look-up for Gaz compostion in Database sheet SBasdo
'    '----------------------------------------------------------------
'    ' L. Ferrand, 01/2009
'    '----------------------------------------------------------------
'    Dim i As Integer
'    Dim j As Integer
'
'
'    With Worksheets(SBasdo)
'
'        i = 1
'        While Name <> .Range(CelDataOrig).Offset(0, i) And i <= .Range(CelDataOrig)
'            i = i + 1
'        Wend
'        If i > .Range(CelDataOrig) Then
'            Call MsgBox("The Gaz " & Name & " was not found in the database.", vbExclamation, Fox)
'
'            Exit Sub
'        End If
'
'        j = 0
'        NbPureGaz = 0
'        While j <= .Range(CelDataOrig).Offset(-1, 0)
'            j = j + 1
'
'            If .Range(CelDataOrig).Offset(j, i) <> 0 Then
'                NbPureGaz = NbPureGaz + 1
'                ReDim Preserve PureGazNames(NbPureGaz)
'                ReDim Preserve PureGazCompo(NbPureGaz)
'                PureGazNames(NbPureGaz) = .Range(CelDataOrig).Offset(j, 0)
'                PureGazCompo(NbPureGaz) = .Range(CelDataOrig).Offset(j, i)
'            End If
'        Wend
'
'    End With
'
'End Sub



Public Sub RefracLoadDatabase(SName As String, ListTarget As Range, nbTargetCell As Integer, TargetCell() As String, IsMessage As Boolean)

    Dim intFile As Integer
    Dim strFic As String
    Dim Text As String
    Dim MatName As String
    Dim nbMat As Integer
    
    Dim CurrentSheet As String
    
    strFic = Worksheets(SFiles).Range(CelPathRefrac)
    
    intFile = FreeFile
    
    '-- Ouverture de la base
    Open strFic For Input As intFile
    nbMat = 0
       
    While Not EOF(intFile)
    
        Line Input #intFile, Text
            
         
        '---- Test si Nom matériau
        If Left(Text, 1) = "*" Then
            
            nbMat = nbMat + 1
            '---- Récupération du nom
            MatName = Mid(Text, 2)
            Worksheets(SRefracRo).Range(CelDataOrig).Offset(0, nbMat) = MatName
            Worksheets(SRefracCo).Range(CelDataOrig).Offset(0, nbMat) = MatName
            Worksheets(SRefracCp).Range(CelDataOrig).Offset(0, nbMat) = MatName
            Worksheets(SRefracEmiss).Range(CelDataOrig).Offset(0, nbMat) = MatName
            Worksheets(SRefracPoro).Range(CelDataOrig).Offset(0, nbMat) = MatName
        ElseIf (Left(Text, 2) = "Ro") Then
            FeuilleName = SRefracRo
            i = 1
        ElseIf (Left(Text, 2) = "Cp") Then
            FeuilleName = SRefracCp
            i = 1
        ElseIf (Left(Text, 2) = "Co") Then
            FeuilleName = SRefracCo
            i = 1
        ElseIf (Left(Text, 4) = "Poro") Then
            FeuilleName = SRefracPoro
            i = 1
        ElseIf (Left(Text, 4) = "Emis") Then
            FeuilleName = SRefracEmiss
            i = 1
         Else
            If (i > 0) Then
                Worksheets(FeuilleName).Range(CelDataOrig).Offset(i, nbMat) = Text
                i = i + 1
            End If
         End If
      
    Wend
            '---- Nombre d'éléments dans la base
        Worksheets(SRefracRo).Range(CelDataOrig) = nbMat
        Worksheets(SRefracCp).Range(CelDataOrig) = nbMat
        Worksheets(SRefracCo).Range(CelDataOrig) = nbMat
        Worksheets(SRefracEmiss).Range(CelDataOrig) = nbMat
        Worksheets(SRefracPoro).Range(CelDataOrig) = nbMat
    '---------- Ecriture de la liste des matériaux dans la cellule cible
    For i = 1 To nbMat
        ListTarget.Offset(i) = Worksheets(SRefracRo).Range(CelDataOrig).Offset(0, i)
    Next i
    
    
    
    Close #intFile
    
    Dim RangeText As String
    
    CurrentSheet = ActiveSheet.Name
        
    Worksheets(SName).Activate
        
    For i = 1 To nbTargetCell
        
        
        Range(TargetCell(i)).Select
        
        RangeText = "=" + ListTarget.Offset(1).Address + ":" + ListTarget.Offset(nbMat).Address
        
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=RangeText
        End With
            
    Next i
    
    Worksheets(CurrentSheet).Activate
    If (IsMessage) Then
    Call MsgBox("The refractory database has been loaded.", vbInformation)
    End If
    
End Sub




Function RefracAddDatabase() As Boolean
    
    Dim intFile As Integer
    Dim strFic As String
    Dim i As Integer
    Dim iname As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim Text() As String
    
    Dim ListName() As String
    Dim Name As String
    Dim IndexNew As Integer
    Dim OKnew As Boolean
    
    Dim Valid As Boolean
    
    strFic = Worksheets(SFiles).Range(CelPathRefrac)
    
    intFile = FreeFile
    
    
    With Worksheets(SRefracPanel)
        
        
        '----- Test de la validité des données ------
    iname = Worksheets(SRefracRo).Range(CelDataOrig) + 1
    ReDim Preserve ListName(iname)
        
        '--- Test si le nom existe déjà
        i = 0
        While i < Worksheets(SRefracRo).Range(CelDataOrig)
            i = i + 1
            If Range(CelRefracPanelName) = Worksheets(SRefracRo).Range(CelDataOrig).Offset(0, i) Then
                Valid = False
                Call MsgBox("This Name already exists. Please choose another one.")
                Range(CelRefracPanelName) = "New_Name"
                
            End If
            ListName(i) = "*" + CStr(Worksheets(SRefracRo).Range(CelDataOrig).Offset(0, i))
        Wend
        Valid = True
        If .Range(CelRefracPanelName) = "" Then
            Valid = False
            Call MsgBox("Please enter a valid name.", vbInformation)
        End If
        
        If .Range(CelRefracPanelRo) <= 0 Then
            Valid = False
            Call MsgBox("Please enter a density.", vbInformation)
        End If
        
        If .Range(CelRefracPanelCp) <= 0 Then
            Valid = False
            Call MsgBox("Please enter a valid specific heat.", vbInformation)
        End If
        
        For i = 1 To NDATA
            If .Range(CelRefracPanelCoTable).Offset(i) <= 0 Then
                Valid = False
                Call MsgBox("Please enter a valid conductivity.", vbInformation)
                Exit For
            End If
            
            If .Range(CelRefracPanelEmissTable).Offset(i) <= 0 Or .Range(CelRefracPanelEmissTable).Offset(i) > 1 Then
                Valid = False
                Call MsgBox("Please enter a valid emissivity.", vbInformation)
                Exit For
            End If
            
        Next i
        
        
        If Valid = True Then
        

            
            '--- Ajout du nouveau nom dans la liste
 
            ListName(iname) = "*" + CStr(.Range(CelRefracPanelName))
            
            '--- Classement de la liste dans l'ordre alphabétique
            Call SortArray(ListName)
            
            '--- Récupération de l'index du nouvel élément dans la liste
            IndexNew = 1
            While ListName(IndexNew) <> "*" + CStr(.Range(CelRefracPanelName))
                IndexNew = IndexNew + 1
            Wend
            
            Open strFic For Output As intFile
            
            For j = 1 To iname
                If j = IndexNew Then
                        Print #intFile, ListName(IndexNew)
                        Print #intFile, "Ro"
                        Print #intFile, .Range(CelRefracPanelRo)
                        Print #intFile, "Cp"
                        Print #intFile, .Range(CelRefracPanelCp)
                        Print #intFile, "Poro"
                        Print #intFile, .Range(CelRefracPanelPoro)

                        Print #intFile, "Co"
                        
                        For k = 1 To NDATA
                            Print #intFile, .Range(CelRefracPanelCoTable).Offset(k)
                        Next k
                          
                        Print #intFile, "Emiss"
                        For k = 1 To NDATA
                            Print #intFile, .Range(CelRefracPanelEmissTable).Offset(k)
                       Next k
                Else
                
                    k = 1
                    While ListName(j) <> "*" + Worksheets(SRefracRo).Range(CelDataOrig).Offset(0, k)
                      k = k + 1
                    Wend
                     Print #intFile, ListName(j)
                     Print #intFile, "Ro"
                     Print #intFile, Sheets(SRefracRo).Range(CelDataOrig).Offset(1, k)
                     Print #intFile, "Cp"
                     Print #intFile, Sheets(SRefracCp).Range(CelDataOrig).Offset(1, k)
                     Print #intFile, "Poro"
                     Print #intFile, Sheets(SRefracPoro).Range(CelDataOrig).Offset(1, k)
    
                     Print #intFile, "Co"
                     
                     For l = 1 To NDATA
                         Print #intFile, Sheets(SRefracCo).Range(CelDataOrig).Offset(1 + l, k)
                     Next l
                       
                     Print #intFile, "Emiss"
                     For l = 1 To NDATA
                         Print #intFile, Sheets(SRefracEmiss).Range(CelDataOrig).Offset(1 + l, k)
                    Next l
                End If
            Next
            
   
            Close #intFile
        
            
        End If
        
    End With
    
    
    RefracAddDatabase = Valid
End Function



Function RefracRemoveDatabase(Item As String) As Boolean
    
    Dim intFile As Integer
    Dim strFic As String
    Dim Text() As String
    Dim i As Integer
      
    strFic = Worksheets(SFiles).Range(CelPathRefrac)
    
    intFile = FreeFile
    
    
    RefracRemoveDatabase = False
    
    
    '-- Ouverture de la base
    Open strFic For Input As intFile
    
    i = 0
    While Not EOF(intFile)
        i = i + 1
        ReDim Preserve Text(i)
        Line Input #intFile, Text(i)
        
        If Text(i) = "*" + Item Then
            Line Input #intFile, Text(i)
            
            While (InStr(1, Text(i), "*", vbTextCompare) = 0) And (Not EOF(intFile))
                Line Input #intFile, Text(i)
            Wend
            If EOF(intFile) Then i = i - 2
            RefracRemoveDatabase = True
        End If
        
    Wend
    
    Close #intFile
    
    '-- Ouverture de la base en écriture (écrasement)
    Open strFic For Output As intFile
    
        For j = 1 To i
            Print #intFile, Text(j)
        Next j
        
    Close #intFile
    
    If RefracRemoveDatabase = True Then MsgBox "The refractory : " & Item & " has been removed from the database.", vbInformation
    
End Function



Sub SortArray(ByRef strArray() As String)

    Dim lLoop As Long, lLoop2 As Long

    Dim str1 As String, str2 As String

    'Sort array

    For lLoop = 1 To UBound(strArray)

       For lLoop2 = lLoop To UBound(strArray)

            If UCase(strArray(lLoop2)) < UCase(strArray(lLoop)) Then

                str1 = strArray(lLoop)

                str2 = strArray(lLoop2)

                strArray(lLoop) = str2

                strArray(lLoop2) = str1

            End If

        Next lLoop2

    Next lLoop
    
End Sub


Public Sub FuelLoadDatabase()
'Public Sub FuelLoadDatabase(SName As String, ListTarget As Range, nbTargetCell As Integer, TargetCell() As String)
    Dim strFic As String
    Dim Text As String
    Dim Fuelname As String
    Dim NbFuel As Integer
    Dim NbPureGaz As Variant
    
    Dim CurrentSheet As String
    Dim Cel1 As String, Cel2 As String
    
    strFic = Worksheets(SFiles).Range(CelPathFuel)
    
    intFile = FreeFile
    
    '-- Effacement des éventuelles données existantes
    Cel1 = Worksheets(SFuel).Range(CelDataOrig).Offset(1, 1).Address
    Cel2 = Worksheets(SFuel).Range(CelDataOrig).Offset(24, 150).Address
    
    Worksheets(SFuel).Range(Cel1, Cel2).ClearContents
   
    '-- Ouverture de la base
    Open strFic For Input As intFile
    NbFuel = 0
       
    While Not EOF(intFile)
    
        Line Input #intFile, Text
            
         
        '---- Test si Nom fuel
        If Left(Text, 1) = "*" Then
            
            NbFuel = NbFuel + 1
            '---- Récupération du nom
            Fuelname = Mid(Text, 2)
            Worksheets(SFuel).Range(CelDataOrig).Offset(0, NbFuel) = Fuelname
            
            '---- Récupération du nombre de gaz purs
            Line Input #intFile, NbPureGaz
            
            NbPureGaz = CInt(NbPureGaz)
            
            For i = 1 To NbPureGaz
                '--- Nom du Gaz pur
                Line Input #intFile, Text
                
                j = 1
                '--- recherche du gaz pur dans la liste et écriture de la fraction vol. correspondante
                While Worksheets(SFuel).Range(CelDataOrig).Offset(j) <> Text
                    j = j + 1
                Wend
                    
                '--- Fraction vol.
                Line Input #intFile, Text
                Worksheets(SFuel).Range(CelDataOrig).Offset(j, NbFuel) = Text
                
            Next i
            
        End If
        
        '---- Nombre d'éléments dans la base
        Worksheets(SFuel).Range(CelDataOrig) = NbFuel
        
        
    Wend
    
'    '---------- Ecriture de la liste des matériaux dans la cellule cible
'    For i = 1 To nbFuel
'        ListTarget.Offset(i) = Worksheets(SFuel).Range(CelDataOrig).Offset(0, i)
'    Next i
'
'
'
'    Close #intFile
'
'    Dim RangeText As String
'
'    CurrentSheet = ActiveSheet.Name
'
'    Worksheets(SName).Activate
'
'    For i = 1 To nbTargetCell
'
'        Range(TargetCell(i)).Select
'
'        RangeText = "=" + ListTarget.Offset(1).Address + ":" + ListTarget.Offset(nbFuel).Address
'
'        With Selection.Validation
'            .Delete
'            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'            xlBetween, Formula1:=RangeText
'        End With
'
'
'    Next i
'
'    Worksheets(CurrentSheet).Activate
    
    Call MsgBox("The Fuel database has been loaded.", vbInformation)
    
End Sub



Function FuelAddDatabase() As Boolean
    
    Dim intFile As Integer
    Dim strFic As String
    Dim i As Integer
    Dim j As Integer
    Dim Text() As String
    
    Dim nbPure As Integer
    Dim PureGazName() As String
    Dim PureGazCompo() As Double
    
    
    Dim ListName() As String
    Dim Name As String
    Dim IndexNew As Integer
    Dim OKnew As Boolean
    
    Dim Valid As Boolean
    
    strFic = Worksheets(SFiles).Range(CelPathFuel)
    
    intFile = FreeFile
    
    
    With Worksheets(SFuelPanel)
        
        Valid = True
        '----- Test de la validité des données ------
            
        
        '--- Test si le nom existe déjà
        i = 0
        While i < Worksheets(SFuel).Range(CelDataOrig)
            i = i + 1
            If Range(CelFuelPanelName) = Worksheets(SFuel).Range(CelDataOrig).Offset(0, i) Then
                Valid = False
                Call MsgBox("This Name already exists. Please choose another one.")
                Range(CelFuelPanelName) = "New_Name"
                
            End If
        Wend
        
        If .Range(CelFuelPanelName) = "" Then
            Valid = False
            Call MsgBox("Please enter a valid name.", vbInformation)
        End If
        
        If CDbl(.Range(CelFuelPanelSum100)) <> 100 Then
            Valid = False
            Call MsgBox("Please enter valid composition (sum = 100).", vbInformation)
        End If
        
        If Valid = True Then
        
            ' On récupère la liste des noms, pour placer le nouveau dans l'ordre alphabétique
            '-- Ouverture de la base
            Open strFic For Input As intFile
            
            i = 0
            j = 0
            While Not EOF(intFile)
                Line Input #intFile, Name
                '--- Tableau temporaire contenant toute la base
                j = j + 1
                ReDim Preserve Text(j)
                Text(j) = Name
                
                '--- Recherche du nom
                If InStr(1, Name, "*", vbTextCompare) <> 0 Then
                    i = i + 1
                    ReDim Preserve ListName(i)
                    ListName(i) = Name
                End If
            Wend
            
            Close #intFile
            
            '--- Ajout du nouveau nom dans la liste
            ReDim Preserve ListName(i + 1)
            ListName(i + 1) = "*" + CStr(.Range(CelFuelPanelName))
            
            '--- Classement de la liste dans l'ordre alphabétique
            Call SortArray(ListName)
            
            '--- Récupération de l'index du nouvel élément dans la liste
            IndexNew = 1
            While ListName(IndexNew) <> "*" + CStr(.Range(CelFuelPanelName))
                IndexNew = IndexNew + 1
            Wend
            
            If IndexNew = UBound(ListName) Then
                EndOfList = True
            Else
                EndOfList = False
            End If
            
            '--- Ecriture de la base dans le bon ordre
            Open strFic For Output As intFile
            OKnew = False
            
            For i = 1 To UBound(Text)
                                
                '--- Nom détecté, est-il juste après l'indice de ListName
                If EndOfList = False And InStr(1, Text(i), "*", vbTextCompare) <> 0 And OKnew = False Then
                    If Text(i) = ListName(IndexNew + 1) Then
                        i = i - 1 '--- on recule d'un pas
                        OKnew = True
                        '-- On écrit le nouveau
                        Print #intFile, ListName(IndexNew)
                        nbPure = 0
                        For j = 1 To NMAX_FUEl
                            If .Range(CelFuelPanelCompoTable).Offset(j) <> 0 Then
                                nbPure = nbPure + 1
                                ReDim Preserve PureGazName(nbPure)
                                ReDim Preserve PureGazCompo(nbPure)
                                
                                PureGazName(nbPure) = .Range(CelFuelPanelCompoTable).Offset(j, -1)
                                PureGazCompo(nbPure) = .Range(CelFuelPanelCompoTable).Offset(j)
                                
                            End If
                        Next j
                        Print #intFile, nbPure
                        For j = 1 To nbPure
                            Print #intFile, PureGazName(j)
                            Print #intFile, PureGazCompo(j)
                        Next j
                        Print #intFile, ""
                
                    Else
                        Print #intFile, Text(i)
                    End If
                Else
                    Print #intFile, Text(i)
                End If
            Next i
            
            If EndOfList = True Then
                '-- On écrit le nouveau
                Print #intFile, ""
                Print #intFile, ListName(IndexNew)
                nbPure = 0
                For j = 1 To NMAX_FUEl
                    If .Range(CelFuelPanelCompoTable).Offset(j) <> 0 Then
                        nbPure = nbPure + 1
                        ReDim Preserve PureGazName(nbPure)
                        ReDim Preserve PureGazCompo(nbPure)
                                
                        PureGazName(nbPure) = .Range(CelFuelPanelCompoTable).Offset(j, -1)
                        PureGazCompo(nbPure) = .Range(CelFuelPanelCompoTable).Offset(j)
                                
                    End If
                Next j
                Print #intFile, nbPure
                For j = 1 To nbPure
                    Print #intFile, PureGazName(j)
                    Print #intFile, PureGazCompo(j)
                Next j
    
            End If
            Close #intFile
        
            
        End If
        
    End With
    
    
    FuelAddDatabase = Valid
End Function


Function FuelRemoveDatabase(Item As String) As Boolean
    
    Dim intFile As Integer
    Dim strFic As String
    Dim Text() As String
    Dim i As Integer
      
    strFic = Worksheets(SFiles).Range(CelPathFuel)
    
    intFile = FreeFile
    
    
    FuelRemoveDatabase = False
    
    
    '-- Ouverture de la base
    Open strFic For Input As intFile
    
    i = 0
    While Not EOF(intFile)
        i = i + 1
        ReDim Preserve Text(i)
        Line Input #intFile, Text(i)
        
        If Text(i) = "*" + Item Then
            Line Input #intFile, Text(i)
            
            While (InStr(1, Text(i), "*", vbTextCompare) = 0) And (Not EOF(intFile))
                Line Input #intFile, Text(i)
            Wend
            If EOF(intFile) Then i = i - 2
            FuelRemoveDatabase = True
        End If
        
    Wend
    
    Close #intFile
    
    '-- Ouverture de la base en écriture (écrasement)
    Open strFic For Output As intFile
    
        For j = 1 To i
            Print #intFile, Text(j)
        Next j
        
    Close #intFile
    
    If FuelRemoveDatabase = True Then MsgBox "The fuel : " & Item & " has been removed from the database.", vbInformation
    
End Function


Public Sub AirLoadDatabase(SName As String, ListTarget As Range, nbTargetCell As Integer, TargetCell() As String)
    
    Dim intFile As Integer
    Dim strFic As String
    Dim Text As String
    Dim AirName As String
    Dim NbAir As Integer
    Dim NbPureGaz As Variant
    
    Dim CurrentSheet As String
    Dim Cel1 As String, Cel2 As String
    
    strFic = Worksheets(SFiles).Range(CelPathAir)
    
    intFile = FreeFile
    
    '-- Effacement des éventuelles données existantes
    Cel1 = Worksheets(SComburant).Range(CelDataOrig).Offset(1, 1).Address
    Cel2 = Worksheets(SComburant).Range(CelDataOrig).Offset(5, 150).Address
    
    Worksheets(SComburant).Range(Cel1, Cel2).ClearContents
   
    '-- Ouverture de la base
    Open strFic For Input As intFile
    NbAir = 0
       
    While Not EOF(intFile)
    
        Line Input #intFile, Text
            
         
        '---- Test si Nom air
        If Left(Text, 1) = "*" Then
            
            NbAir = NbAir + 1
            '---- Récupération du nom
            AirName = Mid(Text, 2)
            Worksheets(SComburant).Range(CelDataOrig).Offset(0, NbAir) = AirName
            
            '---- Récupération du nombre de gaz purs
            Line Input #intFile, NbPureGaz
            
            NbPureGaz = CInt(NbPureGaz)
            
            For i = 1 To NbPureGaz
                '--- Nom du Gaz pur
                Line Input #intFile, Text
                
                j = 1
                '--- recherche du gaz pur dans la liste et écriture de la fraction vol. correspondante
                While Worksheets(SComburant).Range(CelDataOrig).Offset(j) <> Text
                    j = j + 1
                Wend
                    
                '--- Fraction vol.
                Line Input #intFile, Text
                Worksheets(SComburant).Range(CelDataOrig).Offset(j, NbAir) = Text
                
            Next i
            
        End If
        
        '---- Nombre d'éléments dans la base
        Worksheets(SComburant).Range(CelDataOrig) = NbAir
        
        
    Wend
    
    '---------- Ecriture de la liste des matériaux dans la cellule cible
'    For i = 1 To nbAir
'        ListTarget.Offset(i) = Worksheets(SComburant).Range(CelDataOrig).Offset(0, i)
'    Next i
'
    
    
    Close #intFile
    
'    Dim RangeText As String
'
'    CurrentSheet = ActiveSheet.Name
'
'    Worksheets(SName).Activate
'
'    For i = 1 To nbTargetCell
'
'        Range(TargetCell(i)).Select
'
'        RangeText = "=" + ListTarget.Offset(1).Address + ":" + ListTarget.Offset(nbAir).Address
'
'        With Selection.Validation
'            .Delete
'            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'            xlBetween, Formula1:=RangeText
'        End With
'
'
'    Next i
'
'    Worksheets(CurrentSheet).Activate
    
    Call MsgBox("The Comburant database has been loaded.", vbInformation)
    
End Sub

Function AirAddDatabase() As Boolean
    
    Dim intFile As Integer
    Dim strFic As String
    Dim i As Integer
    Dim j As Integer
    Dim Text() As String
    
    Dim nbPure As Integer
    Dim PureGazName() As String
    Dim PureGazCompo() As Double
    
    
    Dim ListName() As String
    Dim Name As String
    Dim IndexNew As Integer
    Dim OKnew As Boolean
    
    Dim Valid As Boolean
    
    strFic = Worksheets(SFiles).Range(CelPathAir)
    
    intFile = FreeFile
    
    
    With Worksheets(SAirPanel)
        
        Valid = True
        '----- Test de la validité des données ------
            
        
        '--- Test si le nom existe déjà
        i = 0
        While i < Worksheets(SComburant).Range(CelDataOrig)
            i = i + 1
            If Range(CelAirPanelName) = Worksheets(SComburant).Range(CelDataOrig).Offset(0, i) Then
                Valid = False
                Call MsgBox("This Name already exists. Please choose another one.")
                Range(CelAirPanelName) = "New_Name"
                
            End If
        Wend
        
        If .Range(CelAirPanelName) = "" Then
            Valid = False
            Call MsgBox("Please enter a valid name.", vbInformation)
        End If
        
        If CDbl(.Range(CelAirPanelSum100)) <> 100 Then
            Valid = False
            Call MsgBox("Please enter valid composition (sum = 100).", vbInformation)
        End If
        
        If Valid = True Then
        
            ' On récupère la liste des noms, pour placer le nouveau dans l'ordre alphabétique
            '-- Ouverture de la base
            Open strFic For Input As intFile
            
            i = 0
            j = 0
            While Not EOF(intFile)
                Line Input #intFile, Name
                '--- Tableau temporaire contenant toute la base
                j = j + 1
                ReDim Preserve Text(j)
                Text(j) = Name
                
                '--- Recherche du nom
                If InStr(1, Name, "*", vbTextCompare) <> 0 Then
                    i = i + 1
                    ReDim Preserve ListName(i)
                    ListName(i) = Name
                End If
            Wend
            
            Close #intFile
            
            '--- Ajout du nouveau nom dans la liste
            ReDim Preserve ListName(i + 1)
            ListName(i + 1) = "*" + CStr(.Range(CelAirPanelName))
            
            '--- Classement de la liste dans l'ordre alphabétique
            Call SortArray(ListName)
            
            '--- Récupération de l'index du nouvel élément dans la liste
            IndexNew = 1
            While ListName(IndexNew) <> "*" + CStr(.Range(CelAirPanelName))
                IndexNew = IndexNew + 1
            Wend
            
            If IndexNew = UBound(ListName) Then
                EndOfList = True
            Else
                EndOfList = False
            End If
            
            '--- Ecriture de la base dans le bon ordre
            Open strFic For Output As intFile
            OKnew = False
            
            For i = 1 To UBound(Text)
                                
                '--- Nom détecté, est-il juste après l'indice de ListName
                If EndOfList = False And InStr(1, Text(i), "*", vbTextCompare) <> 0 And OKnew = False Then
                    If Text(i) = ListName(IndexNew + 1) Then
                        i = i - 1 '--- on recule d'un pas
                        OKnew = True
                        '-- On écrit le nouveau
                        Print #intFile, ListName(IndexNew)
                        nbPure = 0
                        For j = 1 To NMAX_AIR
                            If .Range(CelAirPanelCompoTable).Offset(j) <> 0 Then
                                nbPure = nbPure + 1
                                ReDim Preserve PureGazName(nbPure)
                                ReDim Preserve PureGazCompo(nbPure)
                                
                                PureGazName(nbPure) = .Range(CelAirPanelCompoTable).Offset(j, -1)
                                PureGazCompo(nbPure) = .Range(CelAirPanelCompoTable).Offset(j)
                                
                            End If
                        Next j
                        Print #intFile, nbPure
                        For j = 1 To nbPure
                            Print #intFile, PureGazName(j)
                            Print #intFile, PureGazCompo(j)
                        Next j
                        Print #intFile, ""
                
                    Else
                        Print #intFile, Text(i)
                    End If
                Else
                    Print #intFile, Text(i)
                End If
            Next i
            
            If EndOfList = True Then
                '-- On écrit le nouveau
                Print #intFile, ""
                Print #intFile, ListName(IndexNew)
                nbPure = 0
                For j = 1 To NMAX_AIR
                    If .Range(CelAirPanelCompoTable).Offset(j) <> 0 Then
                        nbPure = nbPure + 1
                        ReDim Preserve PureGazName(nbPure)
                        ReDim Preserve PureGazCompo(nbPure)
                                
                        PureGazName(nbPure) = .Range(CelAirPanelCompoTable).Offset(j, -1)
                        PureGazCompo(nbPure) = .Range(CelAirPanelCompoTable).Offset(j)
                                
                    End If
                Next j
                Print #intFile, nbPure
                For j = 1 To nbPure
                    Print #intFile, PureGazName(j)
                    Print #intFile, PureGazCompo(j)
                Next j
    
            End If
            Close #intFile
        
            
        End If
        
    End With
    
    
    AirAddDatabase = Valid
End Function


Function AirRemoveDatabase(Item As String) As Boolean
    
    Dim intFile As Integer
    Dim strFic As String
    Dim Text() As String
    Dim i As Integer
      
    strFic = Worksheets(SFiles).Range(CelPathAir)
    
    intFile = FreeFile
    
    
    AirRemoveDatabase = False
    
    
    '-- Ouverture de la base
    Open strFic For Input As intFile
    
    i = 0
    While Not EOF(intFile)
        i = i + 1
        ReDim Preserve Text(i)
        Line Input #intFile, Text(i)
        
        If Text(i) = "*" + Item Then
            Line Input #intFile, Text(i)
            
            While (InStr(1, Text(i), "*", vbTextCompare) = 0) And (Not EOF(intFile))
                Line Input #intFile, Text(i)
            Wend
            If EOF(intFile) Then i = i - 2
            AirRemoveDatabase = True
        End If
        
    Wend
    
    Close #intFile
    
    '-- Ouverture de la base en écriture (écrasement)
    Open strFic For Output As intFile
    
        For j = 1 To i
            Print #intFile, Text(j)
        Next j
        
    Close #intFile
    
    If AirRemoveDatabase = True Then MsgBox "The comburant : " & Item & " has been removed from the database.", vbInformation
    
End Function

'----------------------------------------
' Steel Grades Database Functions
'----------------------------------------


Public Function Steel_GetDensity(Bisra As String)

    '----------------------------------------------------------------
    ' Look-up for Density of steel Bisra density in Database sheet
    '----------------------------------------------------------------
    ' L. Ferrand, 01/2009
    '----------------------------------------------------------------
    
    Dim i As Integer
    
    With Worksheets(SBisraRo)
    
        i = 1
        While Bisra <> .Range(CelDataOrig).Offset(0, i) And i <= .Range(CelDataOrig)
            i = i + 1
        Wend
        If i > .Range(CelDataOrig) Then
            Call MsgBox("The steel " & Bisra & " was not found in the database. Density = 0", vbExclamation, Fox)
            Steel_GetDensity = 0
            
            Exit Function
        End If
        Steel_GetDensity = .Range(CelDataOrig).Offset(1, i)
        
    End With

End Function





Sub Steel_GetCpTable(Bisra As String, nval As Integer, vecT() As Double, vecCp() As Double)
    '----------------------------------------------------------------
    ' Look-up for Specific Heat of steel Bisra in Database sheet
    '----------------------------------------------------------------
    ' L. Ferrand, 01/2009
    '----------------------------------------------------------------
    
    Dim i As Integer
    
    With Worksheets(SBisraCp)
    
        i = 1
        While Bisra <> .Range(CelDataOrig).Offset(0, i) And i <= .Range(CelDataOrig)
            i = i + 1
        Wend
        If i > .Range(CelDataOrig) Then
            Call MsgBox("The steel " & Bisra & " was not found in the database.", vbExclamation, Fox)
            
            Exit Sub
        End If
        
        nval = .Range(CelDataOrig).Offset(-1, 0)
        
        ReDim vecT(nval)
        ReDim vecCp(nval)
        
        For j = 1 To nval
            vecT(j) = .Range(CelDataOrig).Offset(j, 0)
            vecCp(j) = .Range(CelDataOrig).Offset(j, i)
            
        Next j
        
    End With

End Sub




Sub Steel_GetCoTable(Bisra As String, nval As Integer, vecT() As Double, vecCo() As Double)
    '----------------------------------------------------------------
    ' Look-up for Conductivity of steel Bisra in Database sheet
    '----------------------------------------------------------------
    ' L. Ferrand, 01/2009
    '----------------------------------------------------------------
    
    Dim i As Integer
    
    With Worksheets(SBisraCo)
    
        i = 1
        While Bisra <> .Range(CelDataOrig).Offset(0, i) And i <= .Range(CelDataOrig)
            i = i + 1
        Wend
        If i > .Range(CelDataOrig) Then
            Call MsgBox("The steel " & Bisra & " was not found in the database.", vbExclamation, Fox)
            
            Exit Sub
        End If
        
        nval = .Range(CelDataOrig).Offset(-1, 0)
        
        ReDim vecT(nval)
        ReDim vecCo(nval)
        
        For j = 1 To nval
            vecT(j) = .Range(CelDataOrig).Offset(j, 0)
            vecCo(j) = .Range(CelDataOrig).Offset(j, i)
            
        Next j
        
    End With

End Sub

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
