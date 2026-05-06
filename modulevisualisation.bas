Attribute VB_Name = "modulevisualisation"
Sub ModrayVisu()

    Dim strVizMod As String
    Dim strVM As String
    Call ModrayFour(300, 4, 0.8)

    
    
    strVizMod = Worksheets(SFiles).Range(CModray) + "\medit-2.2-win.exe "
    temp1 = numzone
    temp1 = " " + Worksheets(SFiles).Range(CModray) + "\fourmesh_m1.mesh"
'        If FileExists(temp1) = True Then
'             Call FileCopy(temp1, FileName)
             ShellWait (strVizMod + temp1)

End Sub
