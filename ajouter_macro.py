"""
Injecte la macro Worksheet_Change dans la feuille Furnace design
et convertit le classeur en .xlsm (nécessaire pour les macros VBA).

Usage
-----
    py ajouter_macro.py [--excel CHEMIN_XLSX]

Le fichier .xlsm est créé à côté du fichier source.
Le fichier .xlsx original est conservé comme sauvegarde.
"""

import argparse
import time
from pathlib import Path

import win32com.client

_HERE = Path(__file__).parent
_XLSM = _HERE / "BLD VRTF 1.1_modifiable.xlsm"
_XLSX = _HERE / "BLD VRTF 1.1_modifiable.xlsx"
_DEFAULT_XLSX = _XLSM if _XLSM.exists() else _XLSX

# ---------------------------------------------------------------------------
# Code VBA injecté dans le module de la feuille "Furnace design"
# ---------------------------------------------------------------------------

_VBA = """\
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Col 14 (N) = valeurs editables : "Nombre Sections" et "Columns Nber."
    ' Col 12 (L) = label "Nombre Sections"
    ' Col 13 (M) = ID section (1-20) pour les lignes de la table sections

    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Column <> 14 Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    ' Cas 1 : "Nombre Sections" -> col 12 contient le texte "Nombre"
    If InStr(1, CStr(Me.Cells(Target.Row, 12).Value), "Nombre", vbTextCompare) > 0 Then
        If IsNumeric(Target.Value) Then
            Dim newN As Long
            newN = CLng(Target.Value)
            If newN >= 1 And newN <= 20 Then Call AjusterTableSections(newN)
        End If
        GoTo Cleanup
    End If

    ' Cas 2 : "Columns Nber." -> col 13 contient un ID section numerique 1-20
    Dim col13Val As Variant
    col13Val = Me.Cells(Target.Row, 13).Value
    If Not IsNumeric(col13Val) Then GoTo Cleanup
    Dim secId As Long
    secId = CLng(col13Val)
    If secId < 1 Or secId > 20 Then GoTo Cleanup

    If Not IsNumeric(Target.Value) Then GoTo Cleanup
    Dim newNb As Long
    newNb = CLng(Target.Value)
    If newNb < 1 Or newNb > 20 Then GoTo Cleanup

    Call AjusterLignesTubes(secId, newNb)

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Private Sub AjusterTableSections(newN As Long)
    ' Determine le nb de sections actives = max ID section dans le tableau tubes
    Dim headerRow As Long
    Dim r As Long
    headerRow = 0
    For r = 1 To Me.UsedRange.Row + Me.UsedRange.Rows.Count - 1
        If CStr(Me.Cells(r, 1).Value) = "Reelle zone" Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Sub

    Dim currentN As Long
    currentN = 0
    For r = headerRow + 1 To headerRow + 300
        Dim vb As Variant
        vb = Me.Cells(r, 2).Value
        If vb = "" Or IsEmpty(vb) Then Exit For
        If Not IsNumeric(vb) Then Exit For
        If CLng(vb) > currentN Then currentN = CLng(vb)
    Next r

    If newN = currentN Then Exit Sub

    If newN > currentN Then
        ' Ajouter les sections manquantes en lisant Columns Nber depuis la table sections
        Dim i As Long
        For i = currentN + 1 To newN
            Dim nb As Long
            nb = GetColumnsNber(i)
            Call AjusterLignesTubes(i, nb)
        Next i
    Else
        ' Supprimer les sections en trop
        Dim j As Long
        For j = currentN To newN + 1 Step -1
            Call AjusterLignesTubes(j, 0)
        Next j
    End If
End Sub


Private Function GetColumnsNber(secId As Long) As Long
    ' Lit Columns Nber dans la table sections (col M=13, col N=14)
    Dim r As Long
    For r = 1 To Me.UsedRange.Row + Me.UsedRange.Rows.Count - 1
        Dim mv As Variant
        mv = Me.Cells(r, 13).Value
        If IsNumeric(mv) And Not IsEmpty(mv) Then
            If CLng(mv) = secId Then
                Dim nv As Variant
                nv = Me.Cells(r, 14).Value
                If IsNumeric(nv) And Not IsEmpty(nv) Then
                    GetColumnsNber = CLng(nv)
                    Exit Function
                End If
            End If
        End If
    Next r
    GetColumnsNber = 4  ' valeur par defaut
End Function


Private Sub AjusterLignesTubes(secId As Long, newNb As Long)
    ' Trouver la ligne d'en-tete "Reelle zone" en col A
    Dim headerRow As Long
    Dim r As Long
    headerRow = 0
    For r = 1 To Me.UsedRange.Row + Me.UsedRange.Rows.Count - 1
        If CStr(Me.Cells(r, 1).Value) = "Reelle zone" Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then Exit Sub

    ' Trouver la derniere ligne du tableau tubes
    Dim tableEnd As Long
    tableEnd = headerRow
    For r = headerRow + 1 To headerRow + 300
        Dim v0 As Variant
        v0 = Me.Cells(r, 1).Value
        If v0 = "" Or IsEmpty(v0) Then Exit For
        If Not IsNumeric(v0) Then Exit For
        tableEnd = r
    Next r

    ' Trouver les lignes de la section concernee (col B = secId)
    Dim firstRow As Long, lastRow As Long
    firstRow = 0
    lastRow = 0
    For r = headerRow + 1 To tableEnd
        Dim vb As Variant
        vb = Me.Cells(r, 2).Value
        If IsNumeric(vb) Then
            If CLng(vb) = secId Then
                If firstRow = 0 Then firstRow = r
                lastRow = r
            End If
        End If
    Next r

    Dim currentNb As Long
    currentNb = 0
    If firstRow > 0 Then currentNb = lastRow - firstRow + 1

    If newNb = currentNb Then Exit Sub

    If newNb > currentNb Then
        Dim insertAt As Long
        If lastRow = 0 Then
            insertAt = headerRow
            Dim s As Long
            For s = secId - 1 To 1 Step -1
                For r = headerRow + 1 To tableEnd
                    If IsNumeric(Me.Cells(r, 2).Value) Then
                        If CLng(Me.Cells(r, 2).Value) = s Then insertAt = r
                    End If
                Next r
                If insertAt > headerRow Then Exit For
            Next s
        Else
            insertAt = lastRow
        End If

        Dim i As Long
        For i = 1 To newNb - currentNb
            Me.Rows(insertAt + 1).Insert Shift:=xlDown
            Me.Rows(insertAt).Copy
            Me.Rows(insertAt + 1).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
            Me.Rows(insertAt + 1).ClearContents
            Me.Cells(insertAt + 1, 1).Value = secId
            Me.Cells(insertAt + 1, 2).Value = secId
            Me.Cells(insertAt + 1, 3).Value = currentNb + i
            insertAt = insertAt + 1
        Next i
    Else
        Dim j As Long
        For j = 1 To currentNb - newNb
            Me.Rows(lastRow).Delete Shift:=xlUp
            lastRow = lastRow - 1
        Next j
    End If
End Sub
"""


def main():
    parser = argparse.ArgumentParser(
        description="Injecte la macro VBA et génère le .xlsm"
    )
    parser.add_argument("--excel", default=str(_DEFAULT_XLSX),
                        help="Classeur source .xlsx")
    args = parser.parse_args()

    src = Path(args.excel).resolve()
    out = src.with_suffix(".xlsm")

    if not src.exists():
        print(f"ERREUR : fichier introuvable : {src}")
        return

    print(f"Ouverture : {src.name}")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(src))
        ws = wb.Worksheets("Furnace design")

        # Utiliser le CodeName de la feuille pour trouver son module VBA
        vbc = wb.VBProject.VBComponents(ws.CodeName)
        module = vbc.CodeModule
        print(f"  Module VBA : {vbc.Name} -> '{ws.Name}'")

        # Effacer le module existant s'il y a du code
        if module.CountOfLines > 0:
            module.DeleteLines(1, module.CountOfLines)

        # Injecter le code VBA
        module.AddFromString(_VBA)
        print(f"  Macro injectée dans '{ws.Name}'")

        # Sauvegarder en .xlsm (FileFormat 52 = xlOpenXMLWorkbookMacroEnabled)
        wb.SaveAs(str(out), FileFormat=52)
        wb.Close(SaveChanges=False)
        print(f"  Enregistré : {out.name}")

    finally:
        excel.Quit()

    print("Terminé.")


if __name__ == "__main__":
    main()
