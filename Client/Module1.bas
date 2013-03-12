Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal Filename$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal Filename$)
Type VersionType
Alpha1 As Long
Beta1 As Long
Full1 As Long
Alpha2 As Long
Beta2 As Long
Full2 As Long
End Type

Public Version(1) As VersionType

Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = &H400

Public Function IsRunning(pid As Long) As Boolean
Dim hProcess As Long
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
CloseHandle hProcess
IsRunning = hProcess
End Function

Public Sub WriteINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Public Function ReadINI(INISection As String, INIKey As String, INIFile As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space$(5000)
  
    Call GetPrivateProfileString$(INISection, INIKey, szReturn, sSpaces, Len(sSpaces), INIFile)
  
    ReadINI = RTrim$(sSpaces)
    ReadINI = Left$(ReadINI, Len(ReadINI) - 1)
End Function
Public Function FileExiste(ByVal Filename As String, Optional RAW As Boolean = False) As Boolean
    FileExiste = True
    If Not RAW Then
        If LenB(Dir$(App.Path & "\" & Filename)) = 0 Then FileExiste = False
    Else
        If LenB(Dir$(Filename)) = 0 Then FileExiste = False
    End If
End Function
