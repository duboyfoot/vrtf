Attribute VB_Name = "steel"
Function EnthalpySteel(SteelGrade As String, TKelvin As Double) As Double
    
    '------------------------------------------------------------
    ' Look-up table to get the enthalpy of the BISRA [J/kg]
    '------------------------------------------------------------
    ' P. Dubois, 29/09/2015
    '------------------------------------------------------------
    
    Dim i As Integer
    
    Dim Tinf As Double
    Dim Tsup As Double
    
    Dim hinf As Double
    Dim hsup As Double
    
    Dim Col As Integer
    Dim temp As String
    
    ' Search for the steel column
    Col = 2
    
    temp = ThisWorkbook.Sheets("SteelEnthalpy").Cells(2, 2)
    
    While SteelGrade <> temp And Col < 50
        Col = Col + 1
        temp = ThisWorkbook.Sheets("SteelEnthalpy").Cells(2, Col)
    Wend
    
    If Col = 50 Then
        EnthalpySteel = -1
        Exit Function
    End If
    
    
    i = 3
    While ThisWorkbook.Sheets("SteelEnthalpy").Cells(i, 1) < TKelvin And i < 29
        i = i + 1
    Wend
    
    Tinf = ThisWorkbook.Sheets("SteelEnthalpy").Cells(i - 1, 1)
    Tsup = ThisWorkbook.Sheets("SteelEnthalpy").Cells(i, 1)
    
    hinf = ThisWorkbook.Sheets("SteelEnthalpy").Cells(i - 1, Col)
    hsup = ThisWorkbook.Sheets("SteelEnthalpy").Cells(i, Col)
    
    EnthalpySteel = hinf + (hsup - hinf) / (Tsup - Tinf) * (TKelvin - Tinf)


End Function

Function DensitySteel(SteelGrade As String) As Double
    
    '------------------------------------------------------------
    ' Look-up table to get the enthalpy of the BISRA [J/kg]
    '------------------------------------------------------------
    ' P. Dubois, 29/09/2015
    '------------------------------------------------------------
    
    Dim i As Integer
    
   
    Dim Col As Integer
    Dim temp As String
    
    ' Search for the steel column
    Col = 2
    
    temp = ThisWorkbook.Sheets("SteelEnthalpy").Cells(2, 2)
    
    While SteelGrade <> temp And Col < 50
        Col = Col + 1
        temp = ThisWorkbook.Sheets("SteelEnthalpy").Cells(2, Col)
    Wend
    
    If Col = 50 Then
        DensitySteel = -1
        Exit Function
    End If
    
    
   
    DensitySteel = ThisWorkbook.Sheets("SteelEnthalpy").Cells(35, Col)


End Function

