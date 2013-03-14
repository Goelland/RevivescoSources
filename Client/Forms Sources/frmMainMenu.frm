VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMainMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   7590
   ClientLeft      =   150
   ClientTop       =   30
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ForeColor       =   &H000000FF&
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMainMenu.frx":17D2A
   Palette         =   "frmMainMenu.frx":189F4
   Picture         =   "frmMainMenu.frx":1D4B5
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox NEWCOMPTE 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   3600
      ScaleHeight     =   1575
      ScaleWidth      =   3660
      TabIndex        =   18
      Top             =   5640
      Width           =   3660
      Begin VB.TextBox txtPassword2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   720
         Width           =   1845
      End
      Begin VB.TextBox txtpassword22 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   360
         Width           =   1845
      End
      Begin VB.TextBox txtname2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   19
         Top             =   0
         Width           =   1845
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ANNULER"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mot de Passe      "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Verif. mot de passe"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CREER LE COMPTE                "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.PictureBox PERSONNAGES 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   480
      ScaleHeight     =   1575
      ScaleWidth      =   3975
      TabIndex        =   9
      Top             =   5640
      Width           =   3975
      Begin VB.PictureBox PicChar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   2400
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   11
         Top             =   0
         Width           =   960
      End
      Begin VB.ListBox lstChars 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   930
         ItemData        =   "frmMainMenu.frx":FE4F7
         Left            =   0
         List            =   "frmMainMenu.frx":FE4F9
         TabIndex        =   10
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label PicCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANNULER"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   17
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label picUseChar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SELECTIONNER"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1200
         TabIndex        =   16
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label picDelChar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EFFACER"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label picNewChar 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOUVEAU"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblCharNom 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblCharClasse 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.PictureBox LOGIN 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   3975
      TabIndex        =   4
      Top             =   6240
      Width           =   3975
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Mémoriser"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   0
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Login"
         Top             =   0
         Width           =   1635
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   0
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         ToolTipText     =   "Mot de passe"
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label lbl_creer 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Creer un Compte"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label PicConnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Connection"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrPicChar 
      Interval        =   500
      Left            =   5760
      Top             =   0
   End
   Begin MSComctlLib.ImageList imgl 
      Left            =   9000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   65280
      _Version        =   393216
   End
   Begin VB.Timer splash 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7200
      Top             =   0
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   0
   End
   Begin VB.Timer Tmrmusic 
      Interval        =   1000
      Left            =   6720
      Top             =   0
   End
   Begin VB.Label versionlbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2E2E2&
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   690
   End
   Begin VB.Label Blague 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hhhhhhh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   7200
      Width           =   1260
   End
   Begin VB.Label picQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quitter"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   0
      Top             =   6960
      Width           =   1305
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public animi As Long
Public DragImg As Long
Public DragX As Long
Public DragY As Long
Private twippx As Long
Private twippy As Long
Private Texte As String
Private Tailletext As Long
Private TailleText2 As Long
Private Tours As Byte


Public Function getreselotionX()
    getreselotionX = Screen.Width \ Screen.TwipsPerPixelX
End Function

Public Function getreselotionY()
    getreselotionY = Screen.height \ Screen.TwipsPerPixelY
End Function


Private Sub chk_fullscreen_Click()

End Sub

Private Sub Form_GotFocus()
If frmNewChar.Visible Then Call frmNewChar.SetFocus
End Sub

'Private Sub chk_fullscreen_Click()
'    If chk_fullscreen.value = "1" Then
'        Call WriteINI("PLEIN_ECRAN", "actif", "1", App.Path & "\Data.ini")
'        frmMirage.Height = Screen.Height / Screen.TwipsPerPixelY
'        frmMirage.Width = Screen.Width / Screen.TwipsPerPixelX
'        frmMirage.picScreen.Height = Screen.Height / Screen.TwipsPerPixelY
'        frmMirage.picScreen.Width = Screen.Width / Screen.TwipsPerPixelX
'        'ChangeScreenSettings 640, 480, 16
'        'Me.WindowState = "2"
'        'frmMainMenu.BorderStyle = "0"
'    Else
'        Call WriteINI("PLEIN_ECRAN", "actif", "0", App.Path & "\Data.ini")
'        frmMirage.Height = 599 * Screen.TwipsPerPixelY
'        frmMirage.Width = 804 * Screen.TwipsPerPixelX
'        frmMirage.picScreen.Height = 599 * Screen.TwipsPerPixelY
'        frmMirage.picScreen.Width = 804 * Screen.TwipsPerPixelX
'        'Dim Pathy As String
'Pathy = App.Path & "\config.ini"
'ChangeScreenSettings ReadINI("CONFIG", "X", Pathy), ReadINI("CONFIG", "Y", Pathy), 32
'        frmMirage.WindowState = "0"
'    End If
'End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
On Error Resume Next
Dim temptext As String
Dim j As Long
    'SetWindowLong blague.hwnd, -20, WS_EX_TRANSPARENT
PERSONNAGES.Left = 0
NEWCOMPTE.Left = 0
   ' If FileExiste("\config\blague.txt") Then
   '     Open App.Path & "\config\blague.txt" For Input As #1
   '     Do While Not EOF(1)
   '     Input #1, temptext
   '     Texte = Texte & " " & temptext
   '     Loop
   '     Close 1
   '     Texte = Texte + "   "
   ' End If
Call conseil
    
    dragAndDrop = 0
    charSelectNum = 1
    'Check1.value = Val(ReadINI("CONFIG", "Music", App.Path & "\Config\Client.ini"))

    If getreselotionY < 768 Then
    netbook = True
    End If
    
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 1 Then Ending = ".bmp"
        
 
 If FileExiste(Rep_Theme & "\Login\fond" & Ending) Then Me.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\fond" & Ending): TransRegion Me, 255, vbBlue

        'If FileExiste(Rep_Theme & "\Login\connexion" & Ending) Then imgLogin.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\connexion" & Ending)
        'If FileExiste(Rep_Theme & "\Login\nouveau" & Ending) Then imgNouveau.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\nouveau" & Ending)
       ' If FileExiste(Rep_Theme & "\Login\personnage" & Ending) Then PERSONNAGES.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\personnage" & Ending)
         If FileExiste("GFX/Sprites/Sprites0" & Ending) Then
            PicChar.Picture = LoadPNG(App.Path & "/GFX/Sprites/Sprites0" & Ending)
        End If
    Next i

    'If Check1.value = 1 Then If FileExiste("Music\mainmenu.mid") Then Call PlayMidi("mainmenu.mid") Else Call PlayMidi("mainmenu.mp3")
            
    'Picsprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png", True)
    
    NEWCOMPTE.Visible = False
    PERSONNAGES.Visible = False
    txtName.Text = Trim$(ReadINI("INFO", "Account", App.Path & "\Config\Account.ini"))
    txtPassword.Text = Trim$(ReadINI("INFO", "Password", App.Path & "\Config\Account.ini"))
    
    If Trim$(txtPassword.Text) <> vbNullString Then Check2.value = Checked Else Check2.value = Unchecked
    
    LOGIN.Visible = True
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
        
'    If Val(ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini")) = 0 Then
'        chk_fullscreen.value = "0"
'    Else
'        chk_fullscreen.value = "1"
'    End If
    
    'If ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini") = 1 Then
    'ChangeScreenSettings 640, 480, 16
    'End If
    twippy = Screen.TwipsPerPixelY
    twippx = Screen.TwipsPerPixelX
    
    versionlbl.Caption = "Version: " & ReadINI("CONFIG", "Version", App.Path & "\Config\Client.ini")
    
    Me.Icon = FrmMirage.Icon

    LOGIN.Visible = True
    
    Call netbook_change

        
End Sub
Private Sub conseil()
On Error Resume Next
Dim i, j As Long
    Texte = ""
   For j = 1 To 30
   i = Rand(1, Val(LoadResString(0)))
   Next j
   
   Texte = LoadResString(i)
    Tailletext = Len(Texte)
    Blague.Left = 700
    Blague.Caption = Texte
    Timer1.Interval = 100
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picQuit.ForeColor = &HFFFFFF
lbl_creer.ForeColor = &HFFFFFF
PicConnect.ForeColor = &HFFFFFF
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
picCancel.ForeColor = &HFFFFFF
picUseChar.ForeColor = &HFFFFFF
picDelChar.ForeColor = &HFFFFFF
picNewChar.ForeColor = &HFFFFFF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
    End
End Sub

Private Sub imgLogin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 1
DragX = X
DragY = Y
End Sub

Private Sub imgLogin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub



Private Sub imgNouveau_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 2
DragX = X
DragY = Y
End Sub


Private Sub imgNouveau_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub imgPers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 3
    DragX = X
    DragY = Y
End Sub



Private Sub imgPers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragImg = 0
    DragX = 0
    DragY = 0
End Sub

Private Sub Label1_Click()
Dim Msg As String
Dim i As Long
    
    If Trim$(txtname2.Text) <> vbNullString And Trim$(txtpassword22.Text) <> vbNullString And Trim$(txtPassword2.Text) <> vbNullString Then
        Msg = Trim$(txtname2.Text)
        
        If Trim$(txtpassword22.Text) <> Trim$(txtPassword2.Text) Then MsgBox "Le mot de passe ne correspond pas.": Exit Sub
        
        If Len(Trim$(txtname2.Text)) < 3 Or Len(Trim$(txtpassword22.Text)) < 3 Then MsgBox "Votre nom et mot de passe doit contenir plus de 3 caractères.": Exit Sub
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then Call MsgBox("Vous ne pouvez pas utiliser d'accents dans votre nom.", vbOKOnly, GAME_NAME): txtName.Text = vbNullString: Exit Sub
        Next i
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &H808080
End Sub

Private Sub Label2_Click()
    LOGIN.Visible = True
    NEWCOMPTE.Visible = False
End Sub

Private Sub Label6_Click()
 Call GameDestroy
 End
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H808080
End Sub

Private Sub lbl_creer_Click()
    NEWCOMPTE.Visible = True
    LOGIN.Visible = False
End Sub



Private Sub lbl_creer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_creer.ForeColor = &H808080

End Sub

Private Sub LOGIN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picQuit.ForeColor = &HFFFFFF
lbl_creer.ForeColor = &HFFFFFF
PicConnect.ForeColor = &HFFFFFF
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
picCancel.ForeColor = &HFFFFFF
picUseChar.ForeColor = &HFFFFFF
picDelChar.ForeColor = &HFFFFFF
picNewChar.ForeColor = &HFFFFFF
End Sub

Private Sub lstChars_Click()
Dim i As Byte
Dim Ending As String
    charSelectNum = lstChars.ListIndex + 1
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
        
        If FileExiste("/GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending) Then
            PicChar.Picture = LoadPNG(App.Path & "/GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending)
        End If
    Next i
    PicChar.height = PicChar.height / 4
    PicChar.Width = PicChar.Width / 4
    If PicChar.Width > 960 Then
        PicChar.Width = 960
    End If
    If PicChar.height > 960 Then
        PicChar.height = 960
    End If
   ' If PicChar.Width > 480 Then
   '     PicChar.Left = 840 - PicChar.Width + 480
   ' Else
   '     PicChar.Left = 840
   ' End If

    If charSelect(charSelectNum).name <> "" Then
        lblCharNom.Caption = charSelect(charSelectNum).name
        lblCharClasse.Caption = charSelect(charSelectNum).classe
    Else
        lblCharNom.Caption = "Slot Libre"
        lblCharClasse.Caption = ""
    End If
End Sub

Private Sub lstChars_DblClick()
    Call picUseChar_Click
End Sub

Private Sub lstChars_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call picUseChar_Click: KeyAscii = 0
End Sub

Private Sub NEWCOMPTE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picQuit.ForeColor = &HFFFFFF
lbl_creer.ForeColor = &HFFFFFF
PicConnect.ForeColor = &HFFFFFF
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
picCancel.ForeColor = &HFFFFFF
picUseChar.ForeColor = &HFFFFFF
picDelChar.ForeColor = &HFFFFFF
picNewChar.ForeColor = &HFFFFFF
End Sub

Private Sub PERSONNAGES_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picQuit.ForeColor = &HFFFFFF
lbl_creer.ForeColor = &HFFFFFF
PicConnect.ForeColor = &HFFFFFF
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
picCancel.ForeColor = &HFFFFFF
picUseChar.ForeColor = &HFFFFFF
picDelChar.ForeColor = &HFFFFFF
picNewChar.ForeColor = &HFFFFFF
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy(1)
    Sleep (2000)
    LOGIN.Visible = True
    PERSONNAGES.Visible = False

End Sub


Private Sub PicCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picCancel.ForeColor = &H808080
End Sub

Private Sub PicConnect_Click()
    If Trim$(txtName.Text) <> vbNullString And Trim$(txtPassword.Text) <> vbNullString Then
        If Len(Trim$(txtName.Text)) < 3 Or Len(Trim$(txtPassword.Text)) < 3 Then MsgBox "Votre nom et votre mot de passe doivent contenir plus de 3 caractéres": Exit Sub
        Call MenuState(MENU_STATE_LOGIN)
        Call WriteINI("INFO", "Account", txtName.Text, (App.Path & "\Config\Account.ini"))
        If Check2.value = Checked Then Call WriteINI("INFO", "Password", txtPassword.Text, (App.Path & "\Config\Account.ini")) Else Call WriteINI("INFO", "Password", "", (App.Path & "\Config\Account.ini"))
    End If
End Sub

Private Sub PicConnect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicConnect.ForeColor = &H808080
End Sub

Private Sub picDelChar_Click()
Dim value As Long

    If lstChars.List(lstChars.ListIndex) = "Emplacement libre" Then MsgBox "Il n'y a pas de personnage à cette emplacement.": Exit Sub

    value = MsgBox("Es-tu certains de vouloir éffacer ce personnage?", vbYesNo, GAME_NAME)
    
    If value = vbYes Then Call MenuState(MENU_STATE_DELCHAR)
End Sub

Private Sub picDelChar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picDelChar.ForeColor = &H808080
End Sub

Private Sub picNewChar_Click()
    If lstChars.List(lstChars.ListIndex) <> "Emplacement libre" Then MsgBox "Il y a déjà un personnage à cette emplacement.": Exit Sub
    Call SendData("PICVALUE" & END_CHAR)
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picNewChar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picNewChar.ForeColor = &H808080
End Sub

Private Sub picQuit_Click()
'Dim Pathy As String
'Pathy = App.Path & "\config.ini"
'ChangeScreenSettings ReadINI("CONFIG", "X", Pathy), ReadINI("CONFIG", "Y", Pathy), 32
    Call GameDestroy
    End
End Sub

Private Sub picQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picQuit.ForeColor = &H808080
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Emplacement libre" Then MsgBox "Il n'y a pas de personnage à cette emplacement.": Exit Sub
    Call SendData("PICVALUE" & END_CHAR)
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Public Sub ChangeScreenSettings(lWidth As Integer, lHeight As Integer, lColors As Integer)
Dim tDevMode As DEVMODE, lTemp As Long, lIndex As Long

lIndex = 0

Do
    lTemp = EnumDisplaySettings(0&, lIndex, tDevMode)
    If lTemp = 0 Then Exit Do
    lIndex = lIndex + 1
    With tDevMode
        If .dmPelsWidth = lWidth And .dmPelsHeight = lHeight And .dmBitsPerPel = lColors Then lTemp = ChangeDisplaySettings(tDevMode, CDS_UPDATEREGISTRY): Exit Do
    End With
Loop

End Sub

Private Sub picUseChar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picUseChar.ForeColor = &H808080
End Sub

Private Sub splash_Timer()
frmsplash.Visible = False
splash.Enabled = False
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 100
Blague.Left = Blague.Left - 5
If (Blague.Width + Blague.Left) < 0 Then Blague.Left = 700: Tours = Tours + 1
'If TailleText2 = 0 Then blague.Text = ""
  '  TailleText2 = TailleText2 + 1
  '  blague.Text = Mid(Texte, 1, TailleText2)
  '  blague.SelStart = Len(blague.Text)
 '   If Tailletext = TailleText2 Then Timer1.Enabled = False: Timer1.Interval = 5000: TailleText2 = 0: Timer1.Enabled = True
If Tours = 2 Then Tours = 0: Timer1.Enabled = False: Call conseil

End Sub

Private Sub tmr2_Timer()
If Val(ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini")) = 0 Then
    FrmMirage.BorderStyle = 3
    FrmMirage.WindowState = 0
    'frmMirage.StartUpPosition = 1
End If
If Val(ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini")) = 1 Then
    FrmMirage.BorderStyle = 0
    FrmMirage.WindowState = 2
    'frmMirage.StartUpPosition = 2
End If
End Sub

Private Sub Tmrmusic_Timer()
If FrmMirage.Mediaplayer.Controls.currentPosition = 200 Then
    If FileExiste("Music\mainmenu.mid") Then Call PlayMidi("mainmenu.mid") Else Call PlayMidi("mainmenu.mp3")
End If
If Me.Visible = False Then Tmrmusic.Enabled = False Else Tmrmusic.Enabled = True
End Sub

Private Sub tmrPicChar_Timer()
Dim i As Byte
Dim Ending As String
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
        
        If FileExiste("/GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending) Then
            PicChar.Picture = LoadPNG(App.Path & "/GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending)
        End If
    Next i
    PicChar.height = PicChar.height / 4
    PicChar.Width = PicChar.Width / 4
    If PicChar.Width > 960 Then
        PicChar.Width = 960
    End If
    If PicChar.height > 960 Then
        PicChar.height = 960
    End If
  '  If PicChar.Width > 480 Then
  '      PicChar.Left = 840 - PicChar.Width + 480
  '  Else
  '      PicChar.Left = 840
  '  End If
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call PicConnect_Click
End Sub

Private Sub txtname2_GotFocus()
txtname2.SelStart = 0
txtname2.SelLength = Len(txtname2)
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call PicConnect_Click
End Sub

Private Sub txtPassword2_GotFocus()
txtPassword2.SelStart = 0
txtPassword2.SelLength = Len(txtPassword2)
End Sub

Private Sub txtpassword22_GotFocus()
txtpassword22.SelStart = 0
txtpassword22.SelLength = Len(txtpassword22)
End Sub

