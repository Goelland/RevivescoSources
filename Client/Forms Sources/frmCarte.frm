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
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
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
      ScaleHeight     =   252
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   360
      Width           =   4805
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   30
         Left            =   2400
         Top             =   1890
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   47
         Left            =   2400
         Top             =   2850
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   46
         Left            =   2880
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   44
         Left            =   2880
         Top             =   480
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   10
         Left            =   2880
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   3
         Left            =   2880
         Top             =   1440
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   4
         Left            =   2880
         Top             =   1890
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   7
         Left            =   2880
         Top             =   2370
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   24
         Left            =   2880
         Top             =   2850
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   45
         Left            =   3360
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   43
         Left            =   3360
         Top             =   480
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   27
         Left            =   3360
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   2
         Left            =   3360
         Top             =   1440
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   5
         Left            =   3360
         Top             =   1890
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   8
         Left            =   3360
         Top             =   2370
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   23
         Left            =   3360
         Top             =   2850
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   40
         Left            =   3840
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   58
         Left            =   3840
         Top             =   480
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   32
         Left            =   3840
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   18
         Left            =   3840
         Top             =   1440
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   6
         Left            =   3840
         Top             =   1890
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   9
         Left            =   3840
         Top             =   2370
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   22
         Left            =   3840
         Top             =   2850
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   39
         Left            =   4320
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   37
         Left            =   4320
         Top             =   480
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   54
         Left            =   4320
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   19
         Left            =   4320
         Top             =   1410
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   14
         Left            =   4320
         Top             =   1890
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   15
         Left            =   4320
         Top             =   2370
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   16
         Left            =   4320
         Top             =   2850
         Width           =   495
      End
      Begin VB.Shape ShapePos 
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
Call UpdateCarte
End Sub

