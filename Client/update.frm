VERSION 5.00
Begin VB.Form update 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If FileExiste("\Client2.exe") And FileExiste("\client.exe") Then
    Kill (App.Path & "\client.Exe")
    Name App.Path & "\Client2.exe" As App.Path & "\Client.exe"
End If
    End
End Sub
Public Function FileExiste(ByVal Filename As String, Optional RAW As Boolean = False) As Boolean
    FileExiste = True
    If Not RAW Then
        If LenB(Dir$(App.Path & "\" & Filename)) = 0 Then FileExiste = False
    Else
        If LenB(Dir$(Filename)) = 0 Then FileExiste = False
    End If
End Function
