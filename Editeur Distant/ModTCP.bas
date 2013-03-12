Attribute VB_Name = "ModTCP"
Option Explicit
Public SEP_CHAR As String * 1
Public END_CHAR  As String * 1
Public Buffer As String
Public NPCEditName As String
Public NPCEditNum As String
Public TEXTFOCUS As Boolean


Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal filename$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal filename$)

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

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer2 As String
Dim packet As String
Dim Top As String * 3
Dim Start As Long

 
        Main.Winedit.GetData Buffer2, vbString, DataLength
        Buffer = Buffer & Buffer2
        
        Start = InStr(Buffer, END_CHAR)
        Do While Start > 0
            packet = Mid$(Buffer, 1, Start - 1)
            Buffer = Mid$(Buffer, Start + 1, Len(Buffer))
            Start = InStr(Buffer, END_CHAR)
            If Len(packet) > 0 Then
                    Call HandleEditData(packet)
            End If
        Loop
End Sub

Public Function FileExiste(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
    FileExiste = True
    If Not RAW Then
        If LenB(Dir$(App.Path & "\" & filename)) = 0 Then FileExiste = False
    Else
        If LenB(Dir$(filename)) = 0 Then FileExiste = False
    End If
End Function
Sub HandleEditData(ByVal data As String)
Dim parse() As String
Dim f As Long
Dim Msg As String
Dim tableau() As String
Dim i, z As Long
Dim s As String
On Error GoTo er


parse = Split(data, SEP_CHAR)

Select Case LCase(parse(0))

        Case "npc"
            Main.Combo1.Clear
            f = 2

            For i = 1 To Val(parse(1))
            If i = 85 Then
            i = i
            End If
                If Trim(parse(f + 1)) <> "" Then
                    Call Main.Combo1.AddItem(Val(parse(f)) & "- " & Trim(parse(f + 1)))
                    Main.Combo1.ItemData(i - 1) = Val(parse(f))
                Else
                    Call Main.Combo1.AddItem(Val(parse(f)) & "- " & "VIDE")
                    Main.Combo1.ItemData(i - 1) = Val(parse(f))
                End If
            f = f + 2
            Next i
            Main.Combo1.ListIndex = 0
            Main.Frame1.Visible = True
            Exit Sub
            
            
        Case "item"
            Main.Combo2.Clear
            f = 2
            For i = 1 To Val(parse(1))
                If Trim(parse(f + 1)) <> "" Then
                    Call Main.Combo2.AddItem(Val(parse(f)) & "- " & Trim(parse(f + 1)))
                Else
                    Call Main.Combo2.AddItem(Val(parse(f)) & "- " & "VIDE")
                End If
            f = f + 2
            Next i
            Main.Combo2.ListIndex = 0
            Exit Sub
            
        Case "editnpc"
            Main.Text1.text = ""
            Main.Text1.text = parse(1)
            Main.Text3.text = parse(2)
            If Trim(Main.Text1.text) = "vide" Then Main.Text1.text = "Public Sub NPC" & NPCEditNum & "(Byval Index, Byval Texte)" & vbNewLine & vbTab & "Select Case texte" & vbNewLine & vbTab & vbTab & "Case ""bonjour""" & vbNewLine & vbTab & vbTab & vbTab & "Call MapPlayerMsg(index, npc(" & NPCEditNum & ").Name &  "" : Salutations!"", 0)" & vbNewLine & vbTab & vbTab & vbTab & "Exit sub" & vbNewLine & vbTab & "End Select" & vbNewLine & "Exit Sub"
            Exit Sub
            
        Case "loginfail"
            MsgBox parse(1)
            Exit Sub
            
        Case "loginok"
            Main.Frame2.Visible = False
            Call WriteINI("compte", "nom", Trim(Main.TxtName.text), App.Path & "\EDistant.ini")
            Call WriteINI("compte", "IP", Main.Winedit.RemoteHostIP, App.Path & "\EDistant.ini")
            Call WriteINI("compte", "passeword", Main.TxtPasseword.text, App.Path & "\EDistant.ini")
            Exit Sub
        End Select
Exit Sub
er:
MsgBox "Donnée recues incorrectes" & i
Main.Combo1.ListIndex = 0
End Sub

Public Sub SaveText()
Dim f As Long
Call Main.Text1.SaveFile("SaveText.txt", 1)
f = FreeFile



End Sub
