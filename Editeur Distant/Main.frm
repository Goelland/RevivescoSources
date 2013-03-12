VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   9735
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9360
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3000
      TabIndex        =   10
      Top             =   2760
      Width           =   3375
      Begin VB.TextBox TxtName 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Text            =   "Pseudo"
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox TxtPasseword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   13
         Text            =   "Mot de Passe"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   10
         TabIndex        =   12
         Text            =   "IP SERVEUR"
         Top             =   480
         Width           =   2150
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Connection"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   6735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   7200
         TabIndex        =   15
         Text            =   "Combo2"
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Charger NPC"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   120
         Width           =   3135
      End
      Begin VB.ListBox List1 
         Height          =   5520
         Left            =   7440
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   6360
         Width           =   5415
      End
      Begin RichTextLib.RichTextBox Text1 
         Height          =   4935
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1440
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8705
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"Main.frx":08CA
      End
      Begin MSWinsockLib.Winsock Winedit 
         Left            =   9360
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   4001
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Items disponibles"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   16
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Message de bienvenue :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   6405
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   7335
      End
   End
   Begin VB.ListBox List2 
      Height          =   6300
      Left            =   9960
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Npcs 
         Caption         =   "Charger Npcs"
      End
      Begin VB.Menu Quitter 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Combo1_Click()
Dim i, j As Long
i = Combo1.ItemData(Combo1.ListIndex) - 1
j = InStr(1, Combo1.List(i), "-")
NPCEditName = Trim(Mid(Combo1.List(i), j + 1, Len(Combo1.List(i)) - j))
NPCEditNum = Val(Left(Combo1.List(i), j - 1))
Command2.Caption = "Charger NPC " & NPCEditNum
Command2.Visible = True
Text1.text = ""
Command3.Visible = False
End Sub

Private Sub Command1_Click()
Dim s As Boolean
On Error GoTo suite

If Winedit.State <> sckConnected Then
    Winedit.Close
    Call Winedit.Connect(Trim(Text2.text), "4001")
    Exit Sub
Else
    Winedit.SendData ("logination" & SEP_CHAR & Trim(TxtName.text) & SEP_CHAR & Trim(TxtPasseword.text) & END_CHAR)
    Exit Sub
End If
suite:
Main.Caption = "Erreur d'IP ou connection refusée"
End Sub

Private Sub Command2_Click()
If Winedit.State > 7 Then Winedit.Close: Command2.Visible = False: Exit Sub
If Val(Combo1.ItemData(Combo1.ListIndex)) <> 0 Then
    Winedit.SendData ("EDITNPC" & SEP_CHAR & Val(Combo1.ItemData(Combo1.ListIndex)) & END_CHAR)
End If
End Sub

Private Sub Command3_Click()
Command3.Visible = False
Winedit.SendData ("SAVENPC" & SEP_CHAR & Combo1.ItemData(Combo1.ListIndex) & SEP_CHAR & Text1.text & SEP_CHAR & Trim(Text3.text) & END_CHAR)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If TEXTFOCUS = True Then Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If TEXTFOCUS = True Then Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TEXTFOCUS = True Then Exit Sub
End Sub

Private Sub Form_Load()
    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)
    If FileExiste(App.Path & "\Account.ini", 1) Then
        TxtName.text = Trim$(ReadINI("compte", "nom", App.Path & "\EDistant.ini"))
        Text2.text = Trim$(ReadINI("compte", "IP", App.Path & "\EDistant.ini"))
        TxtPasseword.text = Trim$(ReadINI("compte", "passeword", App.Path & "\EDistant.ini"))
    End If
    If FileExiste(App.Path & "\ClsCommands.cls", 1) Then
     Call ChargerAide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SaveText
End Sub


Private Sub List1_Click()
Label1.Caption = Mid(List2.List(List1.ListIndex), 1, InStr(1, List2.List(List1.ListIndex), vbNewLine))
End Sub

Private Sub List1_DblClick()
Dim text As String
text = Replace(Label1.Caption, "ByVal", vbNullString)
text = Replace(text, "As Long", vbNullString)
text = Replace(text, "As String", vbNullString)
text = Replace(text, "As Integer", vbNullString)
text = Replace(text, "As Variant", vbNullString)
Text1.SelLength = 0
Text1.SelText = text
End Sub

Private Sub Npcs_Click()
Main.Winedit.SendData ("NPC" & END_CHAR)
Main.Winedit.SendData ("ITEM" & END_CHAR)
End Sub

Private Sub Quitter_Click()
Unload Me
End Sub

Private Sub Text1_Change()
On Error Resume Next
Command3.Caption = "Sauver NPC " & Combo1.ItemData(Combo1.ListIndex)
Command3.Visible = True
End Sub


Private Sub Text1_GotFocus()
TEXTFOCUS = True
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then
Text1.SelText = vbTab
KeyCode = 0
End If
End Sub

Private Sub Text1_LostFocus()
TEXTFOCUS = False
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(TxtName.text)
End Sub


Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call Command1_Click: Exit Sub
End Sub

Private Sub Timer1_Timer()

If Winedit.State = sckConnected Then Main.Caption = "Connecté" Else Main.Caption = "Déconnecté": Frame1.Visible = False: Frame2.Visible = True
If Winedit.State = sckConnecting Then Main.Caption = "Connection en cours..."
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub TxtName_GotFocus()
TxtName.SelStart = 0
TxtName.SelLength = Len(TxtName.text)
End Sub

Private Sub TxtName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call Command1_Click: Exit Sub
End Sub

Private Sub TxtPasseword_GotFocus()
TxtPasseword.SelStart = 0
TxtPasseword.SelLength = Len(TxtPasseword.text)
End Sub

Private Sub TxtPasseword_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call Command1_Click: Exit Sub
End Sub

Private Sub Winedit_Close()
Frame1.Visible = False
Call SaveText
Frame2.Visible = True
End Sub


Private Sub Winedit_Connect()
Winedit.SendData ("logination" & SEP_CHAR & Trim(TxtName.text) & SEP_CHAR & Trim(TxtPasseword.text) & END_CHAR)
End Sub

Private Sub Winedit_DataArrival(ByVal bytesTotal As Long)
Call IncomingData(bytesTotal)
End Sub
