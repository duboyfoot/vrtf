"""
Injecte les macros VBA dans le classeur VRTF :
  - Feuille "Furnace design" : Worksheet_Change (ajustement tableau sections/tubes)
  - Module standard "ModVRTF" : LancerThermette / LancerModrayThermette
  - Feuille "Solver" : deux boutons liés à ces macros

Usage
-----
    py ajouter_macro.py [--excel CHEMIN_XLSX]

Le fichier .xlsm est créé à côté du fichier source.
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
# Module standard ModVRTF : boutons Solver → calculer.py
# ---------------------------------------------------------------------------

_VBA_CALCUL = """\
Option Explicit

Private Sub RunCalcul(useModray As Boolean)
    Dim xlsmPath As String
    Dim outPath As String
    Dim scriptPath As String
    Dim cmd As String
    Dim wsh As Object
    Dim ret As Long

    xlsmPath = ThisWorkbook.FullName
    scriptPath = ThisWorkbook.Path & "\\calculer.py"
    outPath = Left(xlsmPath, InStrRev(xlsmPath, ".") - 1) & "_résultats.xlsm"

    ThisWorkbook.Save

    Dim q As String
    q = Chr(34)
    cmd = "py " & q & scriptPath & q & " --excel " & q & xlsmPath & q & " --out " & q & outPath & q
    If useModray Then cmd = cmd & " --modray"

    Application.StatusBar = IIf(useModray, "Modray + Thermette en cours...", "Thermette en cours...")

    Set wsh = CreateObject("WScript.Shell")
    ret = wsh.Run("cmd /c " & cmd, 1, True)

    Application.StatusBar = False

    If ret = 0 Then
        If Dir(outPath) <> "" Then Workbooks.Open outPath
        MsgBox "Calcul terminé avec succès.", vbInformation, "VRTF"
    Else
        MsgBox "Erreur lors du calcul (code " & ret & ").", vbCritical, "VRTF"
    End If
End Sub

Sub LancerThermette()
    Call RunCalcul(False)
End Sub

Sub LancerModrayThermette()
    Call RunCalcul(True)
End Sub
"""

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

        # ── 1. Feuille "Furnace design" : Worksheet_Change ────────────────────
        ws_fd = wb.Worksheets("Furnace design")
        vbc_fd = wb.VBProject.VBComponents(ws_fd.CodeName)
        mod_fd = vbc_fd.CodeModule
        print(f"  Module VBA : {vbc_fd.Name} -> '{ws_fd.Name}'")
        if mod_fd.CountOfLines > 0:
            mod_fd.DeleteLines(1, mod_fd.CountOfLines)
        mod_fd.AddFromString(_VBA)
        print(f"  Macro injectée dans '{ws_fd.Name}'")

        # ── 2. Module standard ModVRTF : macros de calcul ─────────────────────
        # Supprimer l'ancien module s'il existe
        for comp in wb.VBProject.VBComponents:
            if comp.Name == "ModVRTF":
                wb.VBProject.VBComponents.Remove(comp)
                break
        # Créer un nouveau module standard (Type 1)
        vbext_ct_StdModule = 1
        mod_vrtf = wb.VBProject.VBComponents.Add(vbext_ct_StdModule)
        mod_vrtf.Name = "ModVRTF"
        mod_vrtf.CodeModule.AddFromString(_VBA_CALCUL)
        print("  Module ModVRTF injecté")

        # ── 3. Feuille "Solver" : deux boutons de calcul ──────────────────────
        ws_sol = wb.Worksheets("Solver")

        # Supprimer les anciens boutons VRTF s'ils existent
        btns = ws_sol.Buttons()
        to_delete = []
        for i in range(1, btns.Count + 1):
            b = btns.Item(i)
            if b.Name in ("BtnThermette", "BtnModray"):
                to_delete.append(b)
        for b in to_delete:
            b.Delete()

        def btn_rect(row1, col1, row2, col2):
            c1 = ws_sol.Cells(row1, col1)
            c2 = ws_sol.Cells(row2, col2)
            left   = c1.Left
            top    = c1.Top
            width  = c2.Left + c2.Width - c1.Left
            height = c2.Top  + c2.Height - c1.Top
            return left, top, width, height

        # Bouton "THERMETTE SEUL"  (D16:H17)
        l, t, w, h = btn_rect(16, 4, 17, 8)
        btn1 = ws_sol.Buttons().Add(l, t, w, h)
        btn1.Text = "THERMETTE SEUL"
        btn1.OnAction = "LancerThermette"
        btn1.Name = "BtnThermette"
        print("  Bouton 'THERMETTE SEUL' ajouté")

        # Bouton "MODRAY + THERMETTE"  (J16:N17)
        l, t, w, h = btn_rect(16, 10, 17, 14)
        btn2 = ws_sol.Buttons().Add(l, t, w, h)
        btn2.Text = "MODRAY + THERMETTE"
        btn2.OnAction = "LancerModrayThermette"
        btn2.Name = "BtnModray"
        print("  Bouton 'MODRAY + THERMETTE' ajouté")

        # ── 4. Sauvegarde ─────────────────────────────────────────────────────
        wb.SaveAs(str(out), FileFormat=52)
        wb.Close(SaveChanges=False)
        print(f"  Enregistré : {out.name}")

    finally:
        excel.Quit()

    print("Terminé.")


if __name__ == "__main__":
    main()
