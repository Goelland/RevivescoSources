VERSION 5.00
Begin VB.Form frmGuild 
   BackColor       =   &H00789298&
   BorderStyle     =   0  'None
   Caption         =   "Création de Guilde"
   ClientHeight    =   5280
   ClientLeft      =   30
   ClientTop       =   -60
   ClientWidth     =   5100
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuild.frx":0000
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picGuildAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   0
      ScaleHeight     =   4875
      ScaleWidth      =   5070
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   5100
      Begin VB.CommandButton cmdMember 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Recruter (comme recruteur)"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox TextName2 
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton cmdAccess 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Rétrograder"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton cmdDisown 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Faire quitter la Guilde"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CommandButton cmdAccess 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Promouvoir"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   3810
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdTrainee 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Recruter"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dissoudre la Guilde"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   17
         Top             =   4680
         Width           =   1470
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vous êtes rang :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblRank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1080
         TabIndex        =   15
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label cmdLeave 
         BackStyle       =   0  'Transparent
         Caption         =   "Quitter la Guilde"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3960
         TabIndex        =   14
         Top             =   4680
         Width           =   1110
      End
      Begin VB.Label lblGuild 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Guild"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4995
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtGuild 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Command2 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   4800
      TabIndex        =   3
      Top             =   0
      Width           =   315
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   1800
      TabIndex        =   2
      Top             =   4800
      Width           =   1485
   End
End
Attribute VB_Name = "frmGuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAccess_Click(Index As Integer)
Dim Packet As String
    List1.SetFocus
    List1.ListIndex = List1.ListIndex
    If List1.ListIndex = -1 Then Exit Sub
    If List1.List(List1.ListIndex) = vbNullString Or List1.List(List1.ListIndex) = vbNullString Or Not IsNumeric(List1.ItemData(List1.ListIndex)) Then Exit Sub
    Packet = "GUILDCHANGEACCESS" & SEP_CHAR & List1.List(List1.ListIndex) & SEP_CHAR & Index & END_CHAR
    Call SendData(Packet)

End Sub

Private Sub cmdDisown_Click()
Dim Packet As String
    If List1.List(List1.ListIndex) = vbNullString Then Exit Sub
    Packet = "GUILDDISOWN" & SEP_CHAR & List1.List(List1.ListIndex) & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdLeave_Click()
Dim Packet As String
Dim s As Boolean
    s = MsgBox("Confirmez vous vouloir quitter la guilde?", vbYesNo)
    If s = True Then
    Packet = "GUILDLEAVE" & END_CHAR
    Call SendData(Packet)
    lblGuild.Caption = vbNullString
    lblRank.Caption = 0
    frmGuild.Visible = False
    FrmMirage.SetFocus
    End If
End Sub

Private Sub cmdMember_Click()
Dim Packet As String
    If TextName2.Text = vbNullString Then Exit Sub
    Packet = "GUILDMEMBER" & SEP_CHAR & TextName2.Text & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdTrainee_Click()
Dim Packet As String
    If TextName2.Text = vbNullString Then Exit Sub
    Packet = "guildtraineevbyesno" & SEP_CHAR & TextName2.Text & END_CHAR '"GUILDTRAINEE"
    Call SendData(Packet)
End Sub


Private Sub Command1_Click()
Dim Packet As String

Packet = "GUILDMAKE" & SEP_CHAR & txtName.Text & SEP_CHAR & txtGuild.Text & END_CHAR

Call SendData(Packet)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
Call SendData("GuildCancel" & END_CHAR)
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
    
        If FileExiste(Rep_Theme & "\jeu\creeguilde" & Ending) Then frmGuild.Picture = LoadPNG(App.Path & Rep_Theme & "\jeu\creeguilde" & Ending)
    Next i

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dr = True
drx = X
dry = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If dr Then DoEvents: If dr Then Call Me.Move(Me.Left + (X - drx), Me.Top + (Y - dry))
If Me.Left > Screen.Width Or Me.Top > Screen.height Then Me.Top = Screen.height \ 2: Me.Left = Screen.Width \ 2
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
dr = False
drx = 0
dry = 0
End Sub

Private Sub Label1_Click()
Dim Packet As String
Dim s As Boolean
    s = MsgBox("Confirmez vous vouloir supprimer définitivement la guilde" & lblGuild.Caption & " ?", vbYesNo)
    If s = True Then
    Packet = "GUILDDELETE" & END_CHAR
    Call SendData(Packet)
    frmGuild.Visible = False
    End If
End Sub

Private Sub List1_Click()
Frame1.Visible = True
Frame1.Caption = List1.List(List1.ListIndex) & " - Rang : " & List1.ItemData(List1.ListIndex)
End Sub

Private Sub picGuild_Click()

End Sub

