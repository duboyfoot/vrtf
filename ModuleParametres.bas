Attribute VB_Name = "ModuleParametres"
Public FlagWorksheet_Change As Boolean

'----------------------------------------------------
'     Déclaration des fonctions API utilisées
'----------------------------------------------------

Public Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" ( _
                ByVal lpPathName As String) As Long
                
Public Declare PtrSafe Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" ( _
                ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'---------------------------------------------------------------
'   Déclaration d'API pour le lancement d'applications Windows
'---------------------------------------------------------------


Private Declare PtrSafe Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ACTIVE = &H103
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'----------- Constante de S. Boltzmann -----------
Public Const sigma As Double = 0.0000000567
Public Const Pi As Double = 3.1415926535




'-----------------------------------'
'     Gestion des versions          '
'-----------------------------------'
Public Const Version As String = "Version 1.0"


'-----------------------------------'
'         Nom des feuilles          '
'-----------------------------------'

Public Const SSave As String = "Save"

Public Const SFiles As String = "Files"
Public Const SGeom As String = "Furnace design"

Public Const SResult As String = "Results"
Public Const SOption As String = "Options"

Public Const SMesh As String = "Mesh"

Public Const SRefracCo As String = "RefracCo"
Public Const SRefracCp As String = "RefracCp"
Public Const SRefracRo As String = "RefracRo"
Public Const SRefracPoro As String = "RefracPoro"
Public Const SRefracEmiss As String = "RefracEps"

Public Const SFuel As String = "Fuels"
Public Const SComburant As String = "Comburants"

Public Const SProp As String = "Prop_comb"
Public Const SHottel As String = "Hottel"

Public Const SFuelPanel As String = "FuelPanel"
Public Const SAirPanel As String = "AirPanel"
Public Const SRefracPanel As String = "RefracPanel"

Public Const SBisraCo As String = "BisraCo"
Public Const SBisraCp As String = "BisraCp"
Public Const SBisraRo As String = "BisraRo"


'-----------------------------------'
'     Feuille fichiers              '
'-----------------------------------'

Public WorkbookApp As Workbook
Public WorkbookData As Workbook
Public ChemThermette As String
Public ChemModray As String
'-------- Chemins ----------

Public Const CDirectory As String = "C17"
Public Const CFileName As String = "C20"

Public Const CThermette As String = "C31"
Public Const CModray As String = "C34"

Public Const CMedit As String = "C37"

Public Const CFuelDatabase As String = "C40"
Public Const CAirDatabase As String = "C42"
Public Const CRefracDatabase As String = "C44"


'-------- Boutons ----------
Public Const COpen As String = "C10"
Public Const CSave As String = "F10"
Public Const CAbout As String = "G3"

Public Const CLoadFuel As String = "G39"
Public Const CLoadAir As String = "G41"

Public Const CAddFuel As String = "E39"
Public Const CRemoveFuel As String = "F39"

Public Const CAddAir As String = "E41"
Public Const CRemoveAir As String = "F41"

Public Const CAddRefrac As String = "E43"
Public Const CRemoveRefrac As String = "F43"

Public Const CLoadRefrac As String = "G43"

'-------- Cellules ----------
Public Const CelReload As String = "d3"


Public Const CelEng As String = "E23"
Public Const CelCustomer As String = "E24"
Public Const CelFurnaceName As String = "E25"
Public Const CelDate As String = "E26"
Public Const CelRevision As String = "E27"


Public Const CelPathFuel As String = "C40"
Public Const CelPathAir As String = "C42"
Public Const CelPathRefrac As String = "C44"


'-----------------------------------------
'            Feuilles Databases
'-----------------------------------------
Public Const CelDataOrig As String = "A2"

Public Const NDATA As Integer = 16



'-----------------------------------------
'            Feuille Geometry
'-----------------------------------------

'---------- Cellules ----------
Public Const CelReelZone As String = "A31"
Public Const CelSect As String = "B31"
Public Const CelRangee As String = "C31"
Public Const CelTubeType As String = "D31"

Public Const CelA As String = "E31"
Public Const CelB As String = "F31"
Public Const CelC As String = "G31"
Public Const CelD As String = "H31"

Public Const CelPreAbs As String = "I31"
Public Const CelDistT_B As String = "J31"
Public Const CelTubePitch As String = "K31"
Public Const CelNBTube As String = "L31"
Public Const CelTubeTemp As String = "M31"
Public Const CelTubeEff As String = "N31"
Public Const CelTubeEmis As String = "O31"
Public Const CelTubeSurf As String = "P31"

Public Const CelLargProd As String = "C13"
Public Const CelEpProd As String = "C12"
Public Const CelTypeProd As String = "C14"
Public Const CelProd As String = "C21"
Public Const CelTempEntree As String = "C23"


Public Const CelNbSect As String = "N12"
Public Const CelHautSect As String = "N13"
Public Const CelLargSect As String = "N14"
Public Const CelSectDesc As String = "N15"
Public Const CelDiamRoul As String = "O15"

Public Const CelMurEp As String = "I12"
Public Const CelMurCond As String = "I13"

Public Const CelVisuSect As String = "J26"
Public Const CelVisuRangee As String = "J27"

'-------- Fuel ------------
Public Const CelFuel As String = "R13"

Public Const CelWobbe As String = "S18"
Public Const CelLHV As String = "S19"
Public Const CelDensity As String = "S20"
Public Const CelVa As String = "S21"
Public Const CelVf As String = "S22"

Public Const CelFuelTop As String = "S24"

Public Const CelFuelCompo As String = "U17"
Public Const CelFuelPourc As String = "V17"

'-------- Comburant ------------

Public Const CelAir As String = "X13"
Public Const CelAirCompo As String = "X17"
Public Const CelAirPourc As String = "Y17"

Public Const CelAirTop As String = "Y24"
Public Const CelAirExces As String = "Y25"


'-----------------------------------------
'            Feuille Option
'-----------------------------------------


Public Const Celno As String = "C2"
Public Const Celyes As String = "C1"
Public Const CelHorizontal As String = "d"
Public Const CelParallele As String = "E1"
Public Const CelAvec As String = "F1"
Public Const CelType As String = "G1"
Public Const CelOptionAir As String = "H1"
Public Const CelOptionChen As String = "I1"



'-----------------------------------------
'            Feuille Solver
'-----------------------------------------


'------- Boutons -------------
Public Const BDefault As String = "F6"
Public Const BSolve As String = "D14"
Public Const BVisu As String = "K30"
'--------- Constantes --------
Public Const SIM_TIME As Double = 240
Public Const INTEGRAL As Double = 20
Public Const GAIN As Double = 10
Public Const DERIVATIVE As Double = 0


'-------- Cellules ------------
Public Const CelSimTime As String = "E8"
Public Const CelGain As String = "E9"
Public Const CelIntegral As String = "E10"
Public Const CelDeriv As String = "E11"
Public Const CelRayon As String = "F8"
Public Const CelComb As String = "J9"
Public Const CelTempAmbiante As String = "J6"
Public Const CelTempDemmarage As String = "J7"
'-----------------------------------------
'            Feuille Results
'-----------------------------------------

Public Const CelNbStrip As String = "A1"
'-------- Tableau GLOBAL THERMAL BALANCE -- INPUTS ------
Public Const CelResProducts_IN As String = "E50"

Public Const CelResFuelTOP_IN As String = "E52"
Public Const CelResComburantTOP_IN As String = "E53"

Public Const CelResComburantRecuExit_IN As String = "E55"
Public Const CelResComburantBurner_IN As String = "E56"

Public Const CelResFuelRege_IN As String = "E58"
Public Const CelResComburantRege_IN As String = "E59"
'
Public Const CelResCombustion_IN As String = "E60"

Public Const CelResScale_IN As String = "E61"

'-------- Tableau GLOBAL THERMAL BALANCE -- OUTPUTS ------

Public Const CelResProducts_OUT As String = "J50"

Public Const CelResWGbeforeRecup_OUT As String = "J51"
Public Const CelResWGafterRecup_OUT As String = "J52"
Public Const CelResWGRecuperated_OUT As String = "J53"
Public Const CelResRecupLoss_OUT As String = "J54"
Public Const CelResPipingLoss_OUT As String = "J55"

Public Const CelResWGBeforeRegeAir_OUT As String = "J56"
Public Const CelResWGBeforeRegeFuel_OUT As String = "J57"

Public Const CelResWGAfterRegeAir_OUT As String = "J58"
Public Const CelResWGAfterRegeFuel_OUT As String = "J59"


Public Const CelResWGExit_OUT As String = "J60"
Public Const CelResWall_OUT As String = "J61"
Public Const CelResDoor_OUT As String = "J62"
Public Const CelResSkidPost_OUT As String = "J63"
Public Const CelResLosses_OUT As String = "J64"


'-----------------------------------------
'            Feuille Volume
'-----------------------------------------
Public Const CelVolume As String = "A2"
Public Const CelSurface As String = "A3"
'-----------------------------------------
'            Feuille Mesh
'-----------------------------------------



Public Const CelNbEchanges As String = "A1"
Public Const CelTempRad As String = "B1"
Public Const CelTableRad As String = "A1"

'-----------------------------------'
'     Feuille Combustion            '
'-----------------------------------'

Public Const CelCombuOrig As String = "A1"

Public Const CelCNbBurners As String = "B1"


Public Const CelCPuis As String = "A3"
Public Const CelCRatio As String = "A4"

Public Const CelCFuel As String = "A5"
Public Const CelCFuelNb As String = "A6"
Public Const CelCFuelNames As String = "A7"
Public Const CelCFuelCompo As String = "A31"

Public Const CelCAir As String = "A55"
Public Const CelCAirNb As String = "A56"
Public Const CelCAirNames As String = "A57"
Public Const CelCAirCompo As String = "A62"

Public Const CelCPolyTairchaud As String = "A67"
Public Const CelCPolyTfuelchaud As String = "A71"
Public Const CelCPolyTfumeesfroidesAir As String = "A75"
Public Const CelCPolyTfumeesfroidesFuel As String = "A79"

Public Const CelCPolyTfumeesEqAir As String = "A83"
Public Const CelCPolyTfumeesEqFuel As String = "A87"

Public Const CelCTfuel As String = "A91"
Public Const CelCTair As String = "A92"

Public Const CelCTairTakeOverP As String = "A136"
Public Const CelCTfuelTakeOverP As String = "A137"

Public Const CelCDebitFuel As String = "A93"
Public Const CelCDebitAir As String = "A94"
Public Const CelCDebitFumees As String = "A95"

Public Const CelCEffRegeFuel As String = "A96"
Public Const CelCAspiRegeFuel As String = "A97"
Public Const CelCEffRegeAir As String = "A98"
Public Const CelCAspiRegeAir As String = "A99"

Public Const CelCExcesAir As String = "A100"

Public Const CelCQAirRecu As String = "A101"
Public Const CelCmAirRecu As String = "A102"

Public Const CelCmFuelRecu As String = "A103"
Public Const CelCmAirRege As String = "A104"
Public Const CelCmFuelRege As String = "A105"

Public Const CelCmFumRecu As String = "A106"
Public Const CelCQFumRecu As String = "A107"
Public Const CelCmFumRegeAir As String = "A108"
Public Const CelCmFumRegeFuel As String = "A109"

Public Const CelCChaleurMassique As String = "A110"
Public Const CelCMasseVolumique As String = "A114"

Public Const CelCFumNb As String = "A115"
Public Const CelCFumNames As String = "A116"
Public Const CelCFumCompo As String = "A121"

Public Const CelCEnthalpieAir As String = "A126"
Public Const CelCEnthalpieFuel As String = "A130"
Public Const CelCEnthalpieFum As String = "A138"

Public Const CelCmAir As String = "A134"
Public Const CelCmFuel As String = "A135"

'----------------------------------------'
'     Feuille Refractory Database panel        '
'----------------------------------------'

'-------- Boutons -----
Public Const CRefracPanelAdd As String = "C29"
Public Const CRefracPanelRemove As String = "R7"

'-------- Cellules ----
Public Const CelRefracPanelName As String = "C5"
Public Const CelRefracPanelRo As String = "C7"
Public Const CelRefracPanelCp As String = "C8"
Public Const CelRefracPanelCoTable As String = "C11"
Public Const CelRefracPanelEmissTable As String = "d1"

Public Const CelRefracPanelRemoveSelect As String = "R5"
Public Const CelRefracPanelRemoveList As String = "R10"



'----------------------------------------'
'     Feuille Fuel Database panel        '
'----------------------------------------'

'-------- Boutons -----
Public Const CFuelPanelAdd As String = "C34"
Public Const CFuelPanelRemove As String = "R7"

'-------- Cellules ----
Public Const CelFuelPanelName As String = "C5"
Public Const CelFuelPanelCompoTable As String = "C7"
Public Const CelFuelPanelSum100 As String = "C32"

Public Const CelFuelPanelRemoveSelect As String = "R5"
Public Const CelFuelPanelRemoveList As String = "R10"


'----------------------------------------'
'     Feuille Air Database panel        '
'----------------------------------------'

'-------- Boutons -----
Public Const CAirPanelAdd As String = "C15"
Public Const CAirPanelRemove As String = "R7"

'-------- Cellules ----
Public Const CelAirPanelName As String = "C5"
Public Const CelAirPanelCompoTable As String = "C7"
Public Const CelAirPanelSum100 As String = "C13"

Public Const CelAirPanelRemoveSelect As String = "R5"
Public Const CelAirPanelRemoveList As String = "R10"




Function FileExists(sFileName As String) As Boolean

    On Error Resume Next
    FileExists = ((GetAttr(sFileName) And vbDirectory) = 0)

End Function



Public Sub ShellWait(ByVal JobToDo As String)

    Dim hProcess As Long, RetVal As Long

    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, vbNormalFocus))
    Do
        GetExitCodeProcess hProcess, RetVal
        DoEvents
        Sleep 100
    Loop While RetVal = STILL_ACTIVE
End Sub


