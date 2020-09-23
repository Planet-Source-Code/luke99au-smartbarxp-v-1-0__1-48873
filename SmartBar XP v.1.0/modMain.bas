Attribute VB_Name = "modMain"
Private Declare Function LockWorkStation Lib "user32.dll" () As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
Public Const EWX_LOGOFF = 0
Public SH As New Shell
Public Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Function LockWS() As Boolean
  Dim i As Integer
  i = LockWorkStation
  LockWS = (i > 0)
End Function

Public Sub ShellAndWait(ByVal JobToDo As String, Optional ExecMode = vbMinimizedNoFocus)
    On Error GoTo ProcedureError
    Const kstrProcedureName = "ShellAndWait"
    Dim ProcessID As Long
    Dim hProcess As Long
    Dim hwnd As Long
    Dim ret As Long

    ProcessID = Shell(JobToDo, CLng(ExecMode))

    hProcess = OpenProcess(&H100000, False, ProcessID)
    ret = WaitForSingleObject(hProcess, -1&)
  
    CloseHandle hProcess

ProcedureExit:
    On Error Resume Next
    Exit Sub
    
ProcedureError:
    Select Case Err.Number
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    Resume ProcedureExit
    Resume
End Sub

Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function


