Attribute VB_Name = "ProcessusSysteme"
'Private Type PROCESSENTRY32
'dwSize As Long
'cntUsage As Long
'th32ProcessID As Long
'th32DefaultHeapID As Long
'th32ModuleID As Long
'cntThreads As Long
'th32ParentProcessID As Long
'pcPriClassBase As Long
'dwFlags As Long
'szExeFile As String * 260
' End Type
'
''Déclarations d'API
''Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
''Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
''Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
''Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
''Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
''Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
''Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
''Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Private Const PROCESS_QUERY_INFORMATION = &H400
'Private Const STILL_ACTIVE = &H103
'
'Public Sub ShellWait(ByVal JobToDo As String)
'
'Dim hProcess As Long, RetVal As Long
'
'     hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(JobToDo, vbNormalFocus))
'     Do
'         GetExitCodeProcess hProcess, RetVal
'         DoEvents
'         Sleep 100
'     Loop While RetVal = STILL_ACTIVE
' End Sub
'
'Public Sub Form_Load()
'    Dim Processus As PROCESSENTRY32
'    Capture = CreateToolhelp32Snapshot(2, 0)
'     'Capture permete de parcourir la liste des processus du système
'     Processus.dwSize = Len(Processus)
'
'    courant = Process32First(Capture, Processus)
'    Do While courant
'    If Left$(Processus.szExeFile, IIf(InStr(1, Processus.szExeFile, Chr$(0)) > 0, InStr(1, Processus.szExeFile, Chr$(0)) - 1, 0)) = "VizNOF.exe" Then
'    'Si "VizNof.exe" est trouvé dans les processus du système, le parcours des processus s'arrete là
'     courant = False
'     Else
'     'Processus suivant
'     courant = Process32Next(Capture, Processus)
'      End If
'      Loop
'
'    CloseHandle Capture
'
'    'Si "VizNOF.exe" a été trouvé, courant=False puisqu'on a manuellement définit cette valeur pour arreter la boucle ; dans ce cas, TypeName(courant)="Boolean"
'     'Si "VizNOF.exe" n'a pas été trouvé, la boucle est allée jusqu'au dernier processus du système ; dans ce cas, TypeName(courant)="Long" car courant=0
'
'       If TypeName(courant) = "Boolean" Then
'       Identifiant = OpenProcess(1, 0, Processus.th32ProcessID)
'     TerminateProcess Identifiant, 0
'      CloseHandle Identifiant
'      End If
'
'    '   Unload Me
'   End Sub
'
