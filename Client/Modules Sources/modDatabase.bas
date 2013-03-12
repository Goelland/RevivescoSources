Attribute VB_Name = "modDatabase"
Option Explicit
' This code to be inserted into a module

Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function CopyIcon Lib "user32" (ByVal hcur As Long) As Long

Private Declare Function SetClassLongPtr Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Const GCW_HCURSOR = (-12)

Private Const OCR_NORMAL = 32512

Public SysCursHandle As Long, Curs2Handle As Long
Public lngOldCursor As Long, lngNewCursor As Long
Public Raccourcit(0 To 23) As Long

'les canneaux du discution et onglets de chat

Private Type RTBType 'enregistrement des canneaux par Richtextbox
Canal(0 To 5) As Boolean
End Type
Public RTB(0 To 3) As RTBType 'variable correspondante a chaque RTB cree
Public OngletActif As Integer
Public PopupOK As Boolean




Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then StripTerminator = Left$(strString, intZeroPos - 1) Else StripTerminator = strString
End Function

Public Function FileExiste(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
    FileExiste = True
    If Not RAW Then
        If LenB(Dir$(App.Path & "\" & filename)) = 0 Then FileExiste = False
    Else
        If LenB(Dir$(filename)) = 0 Then FileExiste = False
    End If
End Function
Sub SaveLocalMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\maps\map" & MapNum & ".fcc"
                            
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub LoadMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\maps\map" & MapNum & ".fcc"
        
    If Not FileExiste("maps\map" & MapNum & ".fcc") Then Exit Sub
    f = FreeFile
    Open filename For Binary As #f
        Get #f, , Map(MapNum)
    Close #f
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
    GetMapRevision = Map(MapNum).Revision
End Function

Sub MoveFrame(PB As Frame, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a, B As Long

a = PB.Top + ((Y / twippy) - (DragY / twippy))
B = PB.Left + ((X / twippx) - (DragX / twippx))

If B < 50 - PB.Width Then B = 50 - PB.Width
If B > FrmMirage.picScreen.Width - 50 Then B = FrmMirage.picScreen.Width - 50
If a < 50 - PB.height Then a = 50 - PB.height
If a > FrmMirage.picScreen.height - 50 Then a = FrmMirage.picScreen.height - 50

If Button = 1 Then
    PB.Left = B: PB.Top = a
End If
End Sub
Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a, B As Long

a = PB.Top + Y - DragY
B = PB.Left + X - DragX

If B < 50 - PB.Width Then B = 50 - PB.Width
If B > FrmMirage.picScreen.Width - 50 Then B = FrmMirage.picScreen.Width - 50
If a < 50 - PB.height Then a = 50 - PB.height
If a > FrmMirage.picScreen.height - 50 Then a = FrmMirage.picScreen.height - 50

If Button = 1 Then
    PB.Left = B: PB.Top = a
End If
End Sub

Sub MoveForm(f As Form, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim GlobalX As Integer
Dim GlobalY As Integer

GlobalX = f.Left
GlobalY = f.Top

If Button = 1 Then f.Left = GlobalX + X - DragX: f.Top = GlobalY + Y - DragY
End Sub


Public Sub StartCursor(AniFilePath As String)

    ' Create a copy of the current cursor,
    ' for Windows NT compatibility
    
    lngOldCursor = CopyIcon(GetCursor())
    
    ' Check the passed string, if it contains
    ' a solid file path, then load the cursor
    ' from file. If not, add the App.Path,
    ' *then* load cursor...
    
    If InStr(1, AniFilePath, "\") Then
        lngNewCursor = LoadCursorFromFile(AniFilePath)
    Else
        lngNewCursor = LoadCursorFromFile(App.Path & _
            "\" & AniFilePath)
    End If
    
    ' Activate the cursor
        
    SetSystemCursor lngNewCursor, OCR_NORMAL
    
End Sub
Public Sub ResetCursor(Cancel As Integer)
        DestroyCursor lngNewCursor
        SetSystemCursor lngOldCursor, OCR_NORMAL
End Sub
Public Sub ChangerTouche(ByVal KeyCode As Long, ByVal KeyChange As Long)
Select Case KeyChange 'On choisit quel touche ("haut bas gauche droite") va etre modifiée
 Case 1
    FrmMirage.key1(0).Caption = "Touche " & NomTouche(KeyCode)
    FrmMirage.key1(0).DataField = KeyCode
    FrmMirage.key1(0).ForeColor = vbRed
    Exit Sub
 Case 2
    FrmMirage.key1(1).Caption = "Touche " & NomTouche(Int(KeyCode))
    FrmMirage.key1(1).DataField = KeyCode
    FrmMirage.key1(1).ForeColor = vbRed
    Exit Sub
 Case 3
    FrmMirage.key1(2).Caption = "Touche " & NomTouche(Int(KeyCode))
    FrmMirage.key1(2).DataField = KeyCode
    FrmMirage.key1(2).ForeColor = vbRed
    Exit Sub
 Case 4
    FrmMirage.key1(3).Caption = "Touche " & NomTouche(Int(KeyCode))
    FrmMirage.key1(3).DataField = KeyCode
    FrmMirage.key1(3).ForeColor = vbRed
    Exit Sub
 Case 5
    FrmMirage.key1(4).Caption = "Touche " & NomTouche(Int(KeyCode))
    FrmMirage.key1(4).DataField = KeyCode
    FrmMirage.key1(4).ForeColor = vbRed
    Exit Sub
 Case 6
    FrmMirage.key1(5).Caption = "Touche " & NomTouche(Int(KeyCode))
    FrmMirage.key1(5).DataField = KeyCode
    FrmMirage.key1(5).ForeColor = vbRed
    Exit Sub
 Case 7
    FrmMirage.key1(6).Caption = "Touche " & NomTouche(Int(KeyCode))
    FrmMirage.key1(6).DataField = KeyCode
    FrmMirage.key1(6).ForeColor = vbRed
    Exit Sub
 Case 8
    FrmMirage.key1(7).Caption = "Touche " & NomTouche(Int(KeyCode))
    FrmMirage.key1(7).DataField = KeyCode
    FrmMirage.key1(7).ForeColor = vbRed
    Exit Sub
 End Select
        
 If KeyChange >= 10 And KeyChange <= 24 Then
    FrmMirage.key2(KeyChange - 10).Caption = "Touche " & NomTouche(Int(KeyCode))
    FrmMirage.key2(KeyChange - 10).DataField = KeyCode
    FrmMirage.key2(KeyChange - 10).ForeColor = vbRed
 End If

End Sub

Public Sub LoadTouches(Index As Byte)



Select Case Index

    Case 1
        FrmMirage.key1(0).Caption = NomTouche(Val(ReadINI("TJEU", "haut", App.Path & "\Config\Option.ini")))
        FrmMirage.key1(0).DataField = Val(ReadINI("TJEU", "haut", App.Path & "\Config\Option.ini"))
        FrmMirage.key1(0).ForeColor = vbBlack
        Raccourcit(Index - 1) = FrmMirage.key1(Index - 1).DataField
    Case 2
        FrmMirage.key1(1).Caption = NomTouche(Val(ReadINI("TJEU", "bas", App.Path & "\Config\Option.ini")))
        FrmMirage.key1(1).DataField = Val(ReadINI("TJEU", "bas", App.Path & "\Config\Option.ini"))
        FrmMirage.key1(1).ForeColor = vbBlack
        Raccourcit(Index - 1) = FrmMirage.key1(Index - 1).DataField
    Case 3
        FrmMirage.key1(2).Caption = NomTouche(Val(ReadINI("TJEU", "gauche", App.Path & "\Config\Option.ini")))
        FrmMirage.key1(2).DataField = Val(ReadINI("TJEU", "gauche", App.Path & "\Config\Option.ini"))
        FrmMirage.key1(2).ForeColor = vbBlack
        Raccourcit(Index - 1) = FrmMirage.key1(Index - 1).DataField
    Case 4
        FrmMirage.key1(3).Caption = NomTouche(Val(ReadINI("TJEU", "droite", App.Path & "\Config\Option.ini")))
        FrmMirage.key1(3).DataField = Val(ReadINI("TJEU", "droite", App.Path & "\Config\Option.ini"))
        FrmMirage.key1(3).ForeColor = vbBlack
        Raccourcit(Index - 1) = FrmMirage.key1(Index - 1).DataField
    Case 5
        FrmMirage.key1(4).Caption = NomTouche(Val(ReadINI("TJEU", "attaque", App.Path & "\Config\Option.ini")))
        FrmMirage.key1(4).DataField = Val(ReadINI("TJEU", "attaque", App.Path & "\Config\Option.ini"))
        FrmMirage.key1(4).ForeColor = vbBlack
        Raccourcit(Index - 1) = FrmMirage.key1(Index - 1).DataField
    Case 6
        FrmMirage.key1(5).Caption = NomTouche(Val(ReadINI("TJEU", "courir", App.Path & "\Config\Option.ini")))
        FrmMirage.key1(5).DataField = Val(ReadINI("TJEU", "courir", App.Path & "\Config\Option.ini"))
        FrmMirage.key1(5).ForeColor = vbBlack
        Raccourcit(Index - 1) = FrmMirage.key1(Index - 1).DataField
    Case 7
        FrmMirage.key1(6).Caption = NomTouche(Val(ReadINI("TJEU", "ramasser", App.Path & "\Config\Option.ini")))
        FrmMirage.key1(6).DataField = Val(ReadINI("TJEU", "ramasser", App.Path & "\Config\Option.ini"))
        FrmMirage.key1(6).ForeColor = vbBlack
        Raccourcit(Index - 1) = FrmMirage.key1(Index - 1).DataField
    Case 8
        FrmMirage.key1(7).Caption = NomTouche(Val(ReadINI("TJEU", "action", App.Path & "\Config\Option.ini")))
        FrmMirage.key1(7).DataField = Val(ReadINI("TJEU", "action", App.Path & "\Config\Option.ini"))
        FrmMirage.key1(7).ForeColor = vbBlack
        Raccourcit(Index - 1) = FrmMirage.key1(Index - 1).DataField

End Select
    
   ' For i = 0 To 13
   '     cbtr(i).ListIndex = CByte(Val(ReadINI("TRAC", "rac" & (i + 1), App.Path & "\Config\Option.ini")))
   '     cbtr(i).Text = optTouche(CByte(Val(ReadINI("TRAC", "rac" & (i + 1), App.Path & "\Config\Option.ini")))).nom
   ' Next i
 If Index >= 10 And Index <= 23 Then
        FrmMirage.key2(Index - 10).Caption = NomTouche(Val(ReadINI("TRAC", "rac" & (Index - 9), App.Path & "\Config\Option.ini")))
        FrmMirage.key2(Index - 10).DataField = Val(ReadINI("TRAC", "rac" & (Index - 9), App.Path & "\Config\Option.ini"))
        FrmMirage.key2(Index - 10).ForeColor = vbBlack
        Raccourcit(Index) = FrmMirage.key2(Index - 10).DataField
End If
End Sub
Public Sub NewRTBChat(ByVal Index As Long)
With FrmMirage
.RTBChat(Index).Left = .RTBChat(0).Left
.RTBChat(Index).Top = .RTBChat(0).Top
.RTBChat(Index).Width = .RTBChat(0).Width
.RTBChat(Index).height = .RTBChat(0).height
.RTBChat(Index).Text = ""
End With
End Sub
Public Sub LoadChat()
Dim i, j, k As Long

j = Val(ReadINI("ONGLETTOTAL", "total", App.Path & "\Config\Ecriture.ini"))

    For i = 1 To j
        FrmMirage.Onglet(i).Caption = ReadINI("ONGLET" & i, "NOM", App.Path & "\Config\Ecriture.ini")
    Next i
    
    For i = 1 To j
        Load FrmMirage.RTBChat(i)
        Call NewRTBChat(i)
        For k = 0 To FrmMirage.Check1.count - 1
            RTB(i).Canal(k) = Val(ReadINI("ONGLET" & j, "Canal" & k, App.Path & "\Config\Ecriture.ini"))
        Next k
    Next i
End Sub
Public Sub ShowChat(ByVal Index As Long)
Dim i As Long
For i = 0 To Val(ReadINI("ONGLETTOTAL", "total", App.Path & "\Config\Ecriture.ini"))
    If i <> Index Then FrmMirage.RTBChat(i).Visible = False
Next i
FrmMirage.RTBChat(Index).Visible = True
FrmMirage.RTBChat(Index).SelStart = Len(FrmMirage.RTBChat(Index).Text)
End Sub
Public Sub GuildUpdate(ByVal data As String)
Dim Parse() As String
Dim i As Long
Dim j As Long
Parse = Split(data, SEP_CHAR)
frmGuild.List1.Clear
j = 0
For i = 1 To (UBound(Parse) - 1)

    Call frmGuild.List1.AddItem(Parse(i))
    frmGuild.List1.ItemData(j) = Parse(i + 1)
    i = i + 1
    j = j + 1
Next i
frmGuild.lblGuild.Caption = GetPlayerGuild(MyIndex)
frmGuild.lblRank.Caption = GetPlayerGuildAccess(MyIndex)
End Sub
