VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1200
      Left            =   360
      Pattern         =   "*.ini*"
      TabIndex        =   2
      Top             =   720
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AJOUTER 70 FLAGS AUX COMPTES JOUEURS"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i, j, k As Long

File1.ListIndex = 1

For i = 1 To File1.ListCount - 1
    File1.ListIndex = i
    For j = 1 To 3

        Call PutVar(App.Path & "\" & File1.List(i), "CHAR" & j, "ID", Int(Rnd(100) * 100) & Int(Rnd(100) * 100))
            For k = 1 To 70
                
                Call PutVar(App.Path & "\" & File1.List(i), "CHAR" & j, "flag" & k, 0)
    Next k
    Next j

   
Next i



End Sub

