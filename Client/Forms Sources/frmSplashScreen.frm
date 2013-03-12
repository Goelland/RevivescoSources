VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmSplashScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6390
   ClientLeft      =   4920
   ClientTop       =   4020
   ClientWidth     =   8370
   Icon            =   "frmSplashScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSplashScreen.frx":08EE
   Picture         =   "frmSplashScreen.frx":15B8
   ScaleHeight     =   6390
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicRestart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5160
      Picture         =   "frmSplashScreen.frx":E25FA
      ScaleHeight     =   40
      ScaleMode       =   0  'User
      ScaleWidth      =   40
      TabIndex        =   4
      Top             =   4920
      Width           =   615
   End
   Begin RichTextLib.RichTextBox News 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmSplashScreen.frx":E38FC
   End
   Begin InetCtlsObjects.Inet InetDownload 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox status 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   4695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000C0&
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   7140
      Top             =   5910
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lancer"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape ProgressShape 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   180
      Left            =   120
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   180
      Left            =   120
      Top             =   6120
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H000000C0&
      Height          =   375
      Left            =   4995
      Top             =   5865
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ANNULER"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'| Eclipse Origins - Autoupdater  |
'| Created by: Robin Perris       |
'| Website: freemmorpgmaker.com   |
'|--------------------------------|

' file host
Private UpdateURL As String

' stores the variables for the version downloaders
Private VersionCount As String
Private CurVersion As String
Private TempVersion As String
'Public Ping(1 To 10) As Long
Public PingIndex As Long
Public lngBytesReceived As Long
Public lngFileLength As Long
Public lngBytesReceived2 As Long
Private Downloading As Boolean



Public Sub DestroyUpdater()
    ' kill temp files
    If FileExiste("\tmpClient.ini") Then Kill App.Path & "\tmpClient.ini"
    ' end updater
    InetDownload.Cancel
    
End Sub

Public Sub Update()

Dim Filename As String
Dim i As Long
Dim tableau() As String

    Shape1.Visible = True
    Label1.Visible = True
    
    status.Text = ""
    If Not FileExiste("\Config\Client.ini") Then DestroyUpdater
    UpdateURL = ReadINI("UPDATER", "updateURL", App.Path & "\Config\Client.ini")
    AddProgress "Connection au serveur de mises à jour."

    ' get the file which contains the info of updated files
    DownloadFile UpdateURL & "/Client.ini", App.Path & "\tmpClient.ini"
    DownloadFile UpdateURL & "/News.rtf", App.Path & "\Config\News.rtf"
    
    If FileExiste("\Config\News.rtf") Then
       News.LoadFile (App.Path & "\Config\News.rtf")
    Else
        News.Text = "Aucunes Nouvelles"
    End If
    
    AddProgress "Connection établie!"
    AddProgress "Comparaison des versions."
    
    ' read the version count
    VersionCount = ReadINI("Config", "Version", App.Path & "\tmpClient.ini")
    
    ' check if we've got a current client version saved
    If FileExiste("\Config\Client.ini") Then
        CurVersion = ReadINI("CONFIG", "Version", App.Path & "\Config\Client.ini")
    Else
        CurVersion = "00.00.00"
    End If

    
    tableau() = Split(CurVersion, ".")
    Version(0).Alpha1 = tableau(4)
    Version(0).Beta1 = tableau(2)
    Version(0).Full1 = tableau(0)
    Version(0).Alpha2 = tableau(5)
    Version(0).Beta2 = tableau(3)
    Version(0).Full2 = tableau(1)
    CurVersion = Version(0).Full1 & Version(0).Full2 & Version(0).Beta1 & Version(0).Beta2 & Version(0).Alpha1 & Version(0).Alpha2
    
    Erase tableau
    tableau() = Split(VersionCount, ".")
    Version(1).Alpha1 = tableau(4)
    Version(1).Beta1 = tableau(5)
    Version(1).Full1 = tableau(0)
    Version(1).Alpha2 = tableau(5)
    Version(1).Beta2 = tableau(3)
    Version(1).Full2 = tableau(1)
    VersionCount = Version(1).Full1 & Version(1).Full2 & Version(1).Beta1 & Version(1).Beta2 & Version(1).Alpha1 & Version(1).Alpha2
    Me.Caption = "Version installée: " & Version(0).Full1 & Version(0).Full2 & "." & Version(0).Beta1 & Version(0).Beta2 & "." & Version(0).Alpha1 & Version(0).Alpha2 & " - Version disponible: " & Version(1).Full1 & Version(1).Full2 & "." & Version(1).Beta1 & Version(1).Beta2 & "." & Version(1).Alpha1 & Version(1).Alpha2
    
On Error GoTo Suite
    Downloading = True
    ' are we up to date?
    If CurVersion < VersionCount Then
        ' make sure it's not 0!
        If CurVersion = 0 Then CurVersion = 1
        ' loop around, download and unrar each update
        For i = CurVersion + 1 To VersionCount
            ' let them know!
           
            AddProgress "Téléchargement de la Version " & i & " en cours."
            Filename = "Version" & i & ".rar"
            ' set the download going through inet
           ProgressShape.Width = 0
            DownloadFile UpdateURL & "/" & Filename, App.Path & "\" & Filename
            'Call URLDownloadToFile(0, UpdateURL & "/" & Filename, App.Path & "\" & Filename, 0, 0)
            ' us the unrar.dll to extract data
            RARExecute OP_EXTRACT, Filename
            ' kill the temp update file
            Kill App.Path & "\" & Filename
            ' update the current version
            If Downloading = False Then Call DestroyUpdater: status.Text = "Mise à jour annulée": Exit Sub
            TempVersion = i
            Version(0).Full = Int(TempVersion / 100)
            Version(0).Beta = Int((TempVersion - Int(TempVersion / 100)) / 10)
            Version(0).Alpha = TempVersion - Int((TempVersion - Int(TempVersion / 100)) / 10)
            Me.Caption = "Version installée: " & Version(0).Full & "." & Version(0).Beta & "." & Version(0).Alpha & " - Version disponible: " & Version(1).Full & "." & Version(1).Beta & "." & Version(1).Alpha
            WriteINI "CONFIG", "Version", Version(0).Full & "." & Version(0).Beta & "." & Version(0).Alpha, App.Path & "\Config\Client.ini"
            ' let them know!
            AddProgress "Version " & i & " installée."
            
        Next
        ' let them know the update has finished
        AddProgress ""
        AddProgress "Mise à jour terminée!!"
        Label1.Visible = False
        Label2.Visible = True
        Shape1.Visible = False
        Shape3.Visible = True
        DestroyUpdater
    Else
        ' they're at the correct version, or perhaps higher!
        AddProgress ""
        AddProgress "Votre jeu est à jour."
        Label1.Visible = False
        Label2.Visible = True
        Shape1.Visible = False
        Shape3.Visible = True
        DestroyUpdater
    End If
    Exit Sub
Suite:
    MsgBox "Erreur de mise à jour"
    
End Sub

Public Sub AddProgress(ByVal sProgress As String, Optional ByVal newline As Boolean = True)
    ' add a string to the textbox on the form
    status.Text = status.Text & sProgress
    If newline = True Then status.Text = status.Text & vbNewLine
    status.SelStart = Len(status.Text)
End Sub

Sub DownloadProgress(intPercent As String)
    ProgressShape.Width = (Shape2.Width * intPercent) / 100
    Label3.Caption = intPercent & "%"
End Sub

'Public Function DownloadFile(strURL As String, strDestination As String) As Boolean
Public Sub DownloadFile(strURL As String, strDestination As String) 'As Boolean
Const CHUNK_SIZE As Long = 1024
Dim intFile As Integer
Dim lngBytesReceived As Long
Dim lngFileLength As Long
Dim strHeader As String
Dim b() As Byte
Dim i As Integer

DoEvents
    
With InetDownload
    
.URL = strURL
.Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
        
While .StillExecuting
DoEvents
Wend

strHeader = .GetHeader
End With
    
    
strHeader = InetDownload.GetHeader("Content-Length")
lngFileLength = Val(strHeader)

DoEvents
    
lngBytesReceived = 0

intFile = FreeFile()

Open strDestination For Binary Access Write As #intFile

Do
b = InetDownload.GetChunk(CHUNK_SIZE, icByteArray)
Put #intFile, , b
lngBytesReceived = lngBytesReceived + UBound(b, 1) + 1

DownloadProgress (Round((lngBytesReceived / lngFileLength) * 100))
DoEvents
Loop While UBound(b, 1) > 0

Close #intFile
 
End Sub




Private Sub Form_Load()
SetWindowLong News.hwnd, -20, WS_EX_TRANSPARENT
Me.Visible = True
Me.Show
Update
'Call ChangeCursor(Me, App.Path & "\Mouse.cur")
End Sub



Private Sub splashtimer_Timer()
    'frmSplashScreen.Visible = False
    'splashtimer.Enabled = False
    'Call Main
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.BorderStyle = 0
Shape3.BorderStyle = 0
PicRestart.Picture = LoadPicture(App.Path & "\images\Restart1.bmp")
End Sub

Private Sub Label1_Click()
InetDownload.Cancel
Downloading = False
Shape1.Visible = False
Label1.Visible = False
PicRestart.Visible = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.BorderStyle = 1
End Sub


Private Sub Label2_Click()
Dim mpid As Long
    mpid = Shell(App.Path & "\Client.ch", vbNormalFocus)
    Do While Not IsRunning(mpid)
     DoEvents
    Loop
    
    End
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BorderStyle = 1
End Sub

Private Sub picrestart_Click()
Update
End Sub

Private Sub picrestart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicRestart.Picture = LoadPicture(App.Path & "\images\Restart2.bmp")
End Sub
