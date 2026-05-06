Attribute VB_Name = "ModuleFichiers"

'-----------------------------------------------------------------------------
' Open the Project File
'-----------------------------------------------------------------------------

Public Sub OpenProjectFile()
        
 
    Dim frm As FrmMessage ' Message de progression
   
    Dim fileToOpen As Variant
    
    Dim i As Integer
    Dim j As Integer
    
    Dim SheetName As Variant
    Dim SheetOffset As Integer
    Dim SheetNbData As Integer
    
    Dim CelName As String
    Dim CelOffsetRow As Integer
    Dim CelOffsetCol As Integer
    Dim CelData As String
    
    
    Set WorkbookApp = ThisWorkbook

    Application.ScreenUpdating = False

   
    SetCurrentDirectory (WorkbookApp.Worksheets(SFiles).Range(CDirectory))
    
    fileToOpen = Application.GetOpenFilename("PilotFurnace case (*.plf), *.plf")
    
    If fileToOpen <> False Then ' Si le nom de fichier est correct
    
        Dim fs As Object
        Dim F As Object
                
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set F = fs.GetFile(fileToOpen)
        
        WorkbookApp.Worksheets(SFiles).Range(CDirectory).Value = F.ParentFolder.Path
        WorkbookApp.Worksheets(SFiles).Range(CFileName).Value = F.Name
                
        
        'Ouverture du fichier sélectionné
        Set WorkbookData = Workbooks.Open(fileToOpen)
        
        WorkbookApp.Activate
                
        '------------------
        ' START Load data
        '------------------
'        Set frm = New FrmMessage
'        frm.Title = "PilotFurnace"
'        frm.Show vbModeless ' Ouvre mais rend la main au code
'        frm.message = "Reading data..."
            
         Worksheets(Solve).Cells(25.1) = "Lecture resultats"
        '----------Chargement des fichiers 1.0--------------------
        
        If WorkbookData.Worksheets(1&).Cells(1, 1) = "Version 1.0" Then
            Worksheets(SGeom).Range(CelWeigthFactor) = 0
        End If
            
            
        For i = 1 To WorkbookData.Worksheets(1&).Range("A2") ' Nombre de feuilles
            
            
            SheetName = WorkbookData.Worksheets(1&).Range("A2").Offset(i)
            SheetOffset = WorkbookData.Worksheets(1&).Range("A2").Offset(i, 1)
            SheetNbData = WorkbookData.Worksheets(1&).Range("A2").Offset(i, 2)
            
            For j = 1 To SheetNbData
                
                If j Mod 10 = 0 Then frm.Progress = j / SheetNbData * 100
        
                CelName = WorkbookData.Worksheets(1&).Cells(SheetOffset + j - 1, 2)
                CelOffsetRow = WorkbookData.Worksheets(1&).Cells(SheetOffset + j - 1, 3)
                CelOffsetCol = WorkbookData.Worksheets(1&).Cells(SheetOffset + j - 1, 4)
                CelData = WorkbookData.Worksheets(1&).Cells(SheetOffset + j - 1, 5)
                
                Worksheets(SheetName).Range(CelName).Offset(CelOffsetRow, CelOffsetCol) = CelData
            Next j
        Next i
                
       WorkbookData.Close
    
    Call frm.CloseForm ' Ferme la fenêtre d'avancement
    Set frm = Nothing
             
    End If
    
    
    Application.ScreenUpdating = True
    
End Sub

Sub SaveProjectFile()

    
    Dim frm As FrmMessage ' Message de progression
   
    Dim fileToSave As Variant
    
    Dim i As Integer
    
    Set WorkbookApp = ThisWorkbook
    
    SetCurrentDirectory (WorkbookApp.Worksheets(SFiles).Range(CDirectory))
    
    fileToSave = Application.GetSaveAsFilename(WorkbookApp.Worksheets(SFiles).Range(CFileName), fileFilter:="PilotFurnace case, *.plf")
    
    If fileToSave <> False Then ' Si le nom de fichier est correct
    
        Dim xlApp As New Excel.Application
        Dim xlBook As Workbook
        Dim Nline As Integer
        
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add

        
        Nline = WorkbookApp.Worksheets(SSave).Range("A2") ' Nombre de feuilles
        Nline = WorkbookApp.Worksheets(SSave).Range("A2").Offset(Nline, 1) + _
            WorkbookApp.Worksheets(SSave).Range("A2").Offset(Nline, 2) - 1 'nb cellules après la dernière feuille
        
        Set frm = New FrmMessage
        frm.Title = "PilotFurnace"
        frm.Show vbModeless ' Ouvre mais rend la main au code
        frm.message = "Saving data..."
            
        For i = 0 To Nline
            If i Mod 20 = 0 Then frm.Progress = i / Nline * 100
            
            For j = 0 To 4
                xlBook.Worksheets(1&).Range("A1").Offset(i, j) = WorkbookApp.Worksheets(SSave).Range("A1").Offset(i, j)
            Next j
        Next i
            
'        Call frm.CloseForm ' Ferme la fenêtre d'avancement
'        Set frm = Nothing
    
        xlBook.SaveAs fileToSave
        xlBook.Close
        xlApp.Quit

        Set xlBook = Nothing
        Set xlApp = Nothing

        Dim fs As Object
        Dim F As Object
                
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set F = fs.GetFile(fileToSave)
        
        WorkbookApp.Worksheets(SFiles).Range(CDirectory).Value = F.ParentFolder.Path
        WorkbookApp.Worksheets(SFiles).Range(CFileName).Value = F.Name
        WorkbookApp.Worksheets(SFiles).Activate
    End If
        
    
   
End Sub

