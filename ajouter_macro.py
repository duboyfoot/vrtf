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
_DEFAULT_XLSX = _HERE / "BLD VRTF 1.1_modifiable.xlsx"

# ---------------------------------------------------------------------------
# Code VBA injecté dans le module de la feuille "Furnace design"
# ---------------------------------------------------------------------------

_VBA = """\
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Ajuste le tableau tubes quand "Columns Nber." change pour une section.
    ' Sections table : col M(13) = ID section, col N(14) = Columns Nber.

    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Column <> 14 Then Exit Sub   ' col N uniquement

    ' Verifier que la cellule modifiee est bien dans la table sections
    Dim secId As Variant
    secId = Me.Cells(Target.Row, 13).Value
    If Not IsNumeric(secId) Then Exit Sub
    secId = CLng(secId)
    If secId < 1 Or secId > 20 Then Exit Sub

    Dim newNb As Variant
    newNb = Target.Value
    If Not IsNumeric(newNb) Then Exit Sub
    newNb = CLng(newNb)
    If newNb < 1 Or newNb > 20 Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False

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
    If headerRow = 0 Then GoTo Cleanup

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

    If newNb = currentNb Then GoTo Cleanup

    If newNb > currentNb Then
        ' --- Ajouter des lignes apres lastRow ---
        Dim insertAt As Long
        If lastRow = 0 Then
            ' Pas encore de lignes : inserer apres la section precedente
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
            Me.Cells(insertAt + 1, 1).Value = secId           ' Reelle zone
            Me.Cells(insertAt + 1, 2).Value = secId           ' Section
            Me.Cells(insertAt + 1, 3).Value = currentNb + i   ' Rangee
            insertAt = insertAt + 1
        Next i

    Else
        ' --- Supprimer les dernieres lignes de cette section ---
        Dim j As Long
        For j = 1 To currentNb - newNb
            Me.Rows(lastRow).Delete Shift:=xlUp
            lastRow = lastRow - 1
        Next j
    End If

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
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

        # Trouver le module VBA correspondant à cette feuille :
        # VBComponents : index 1 = ThisWorkbook, index N+1 = feuille N
        fd_idx = ws.Index
        vbc = wb.VBProject.VBComponents.Item(fd_idx + 1)
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
