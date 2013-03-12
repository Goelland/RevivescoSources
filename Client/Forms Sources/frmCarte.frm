VERSION 5.00
Begin VB.Form frmCarte 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carte"
   ClientHeight    =   4725
   ClientLeft      =   -15
   ClientTop       =   210
   ClientWidth     =   5790
   Icon            =   "frmCarte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3785
      Left            =   480
      Picture         =   "frmCarte.frx":17D2A
      ScaleHeight     =   3780
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   4805
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4320
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmCarte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
        
        If FileExiste("images\Carte" & Ending) Then imgCarte.Picture = LoadPicture(App.Path & "\images\Carte" & Ending)
    Next i
End Sub


