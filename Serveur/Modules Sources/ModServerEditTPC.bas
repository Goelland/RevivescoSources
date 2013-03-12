Attribute VB_Name = "ModServerEditTPC"
Option Explicit
Public Editeur() As EditeurRec
Sub AcceptEditConnection(ByVal Index As Long, ByVal SocketId As Long)
Dim i As Long
On Error Resume Next
        i = FindSlot
        If i <> 0 Then
            frmServer.WinEdit(i).Close
            frmServer.WinEdit(i).Accept SocketId
            Call TextAdd(frmServer.txtText(0), "Editeur Connecté : " & frmServer.WinEdit(i).RemoteHostIP, True)
        End If
End Sub
Private Function FindSlot() As Byte
Dim i As Long
For i = 1 To frmServer.WinEdit.Count - 1
    If frmServer.WinEdit(i).State <> sckConnected Then FindSlot = i: Exit Function
Next i
FindSlot = 0
End Function
Sub IncomingEditData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer As String
Dim packet As String
Dim Top As String * 3
Dim Start As Long

    If Index > 0 Then
        frmServer.WinEdit(Index).GetData Buffer, vbString, DataLength
        Editeur(Index).Buffer = Editeur(Index).Buffer & Buffer
        
        Start = InStr(Editeur(Index).Buffer, END_CHAR)
        Do While Start > 0
            packet = Mid$(Editeur(Index).Buffer, 1, Start - 1)
            Editeur(Index).Buffer = Mid$(Editeur(Index).Buffer, Start + 1, Len(Editeur(Index).Buffer))
            Start = InStr(Editeur(Index).Buffer, END_CHAR)
            If Len(packet) > 0 Then
                    Call HandleEditData(Index, packet)
            End If
        Loop
     End If
End Sub
Sub SendEditDataTo(ByVal Index As Long, ByVal Data As String)
On Error Resume Next
    If frmServer.WinEdit(Index).State = sckConnected Then frmServer.WinEdit(Index).SendData Data: NewDoEvents
End Sub
Sub HandleEditData(ByVal Index As Long, ByVal Data As String)
Dim Parse() As String
Dim packet As String
Dim f As Long
Dim Msg As String
Dim i As Long
Dim s As String
Dim Name As String
Dim Password As String
Dim n As String

On Error GoTo er:
Parse = Split(Data, SEP_CHAR)
'le serveur recoit des infos d'edition mais n'authorise pas les connection : on déconnecte l'éditeur.
If frmServer.Check2.value = 0 Then Call SendEditDataTo(Index, "LoginFail" & SEP_CHAR & "Le serveur refuse la connection de l'éditeur distant." & END_CHAR): CloseWinEdit (Index): Exit Sub
Select Case LCase(Parse(0))

        Case "logination"
            If Editeur(Index).Logged = False Then
                Name = Parse(1)
                Password = Parse(2)
                
                If Len(Name) < 3 Then
                    packet = "LoginFail" & SEP_CHAR & "Nom trop court" & END_CHAR
                 GoTo trap:
                End If
                
                For i = 1 To Len(Name)
                    n = Asc(Mid$(Name, i, 1))
                    If (n <= 65 And n >= 90) Or (n <= 97 And n >= 122) Or (n = 95) Or (n = 32) Or (n <= 48 And n >= 57) Then
                        packet = "LoginFail" & SEP_CHAR & "Caractère spéciaux interdit dans le Nom" & END_CHAR
                 GoTo trap:
                    End If
                Next i
        
                For i = 1 To 3
                    If Val(GetVar(App.Path & "\accounts\" & Trim$(Name) & ".ini", "CHAR" & i, "access")) > 3 Then GoTo suite:
                Next i
                packet = "LoginFail" & SEP_CHAR & "Droits insuffisants" & END_CHAR: GoTo trap
suite:
                If Not AccountExist(Name) Then packet = "LoginFail" & SEP_CHAR & "Compte introuvable" & END_CHAR: GoTo trap
            
                If Not PasswordOK(Name, Password) Then packet = "LoginFail" & SEP_CHAR & "Mot de passe incorrecte" & END_CHAR: GoTo trap
            
                              
                If frmServer.Closed.value = Checked Then packet = "LoginFail" & SEP_CHAR & "Le Serveur est en train fermer" & END_CHAR: GoTo trap
                packet = "loginok" & END_CHAR
                Editeur(Index).Logged = True
                            
trap:
                Call SendEditDataTo(Index, packet)
        
                End If
            Exit Sub
            
            
            Exit Sub
        Case "npc"
            If Editeur(Index).Logged = False Then Exit Sub
            packet = "NPC" & SEP_CHAR & MAX_NPCS & SEP_CHAR
            For i = 1 To MAX_NPCS
                packet = packet & i & SEP_CHAR & " " & Trim(Npc(i).Name) & SEP_CHAR
            Next i
                packet = packet & END_CHAR
                Call SendEditDataTo(Index, packet)
                Exit Sub
                
        Case "item"
            If Editeur(Index).Logged = False Then Exit Sub
            packet = "ITEM" & SEP_CHAR & MAX_ITEMS & SEP_CHAR
            For i = 1 To MAX_ITEMS
                packet = packet & i & SEP_CHAR & " " & Trim(item(i).Name) & SEP_CHAR
            Next i
                packet = packet & END_CHAR
                Call SendEditDataTo(Index, packet)
                Exit Sub
                
               
                
        Case "editnpc"
        If Editeur(Index).Logged = False Then Exit Sub
            packet = "EDITNPC" & SEP_CHAR
            If FileExist("\Npcs\" & "NPC" & Val(Parse(1)) & ".txt") Then
                packet = packet & FileText(App.Path & "\Npcs\" & "NPC" & Val(Parse(1)) & ".txt") & SEP_CHAR
            Else
                 packet = packet & "vide" & SEP_CHAR
            End If
                packet = packet & Npc(Val(Parse(1))).AttackSay & END_CHAR
                Call SendEditDataTo(Index, packet)
                Exit Sub
                
                
        Case "savenpc"
        If Editeur(Index).Logged = False Then Exit Sub
            Call SaveFileText(App.Path & "\Npcs\" & "NPC" & Val(Parse(1)) & ".txt", Parse(2))
            Call TextAdd(frmServer.txtText(0), "Editeur " & frmServer.WinEdit(i).RemoteHostIP & ": a édité le NPC" & Val(Parse(1)), True)
            Npc(Val(Parse(1))).AttackSay = Trim(Parse(3))
            SaveNpc (Val(Parse(1)))
            Call InitNpcScript
            Exit Sub
End Select
Exit Sub
er:
Call TextAdd(frmServer.txtText(0), "Données de l'editeur : " & frmServer.WinEdit(Index).RemoteHostIP & " incorrectes", True)
End Sub
Function GetEditNpc(ByVal npcnum As Long) As String
Dim packet As String
Dim i As Long

 packet = Trim$(Npc(npcnum).Name) & SEP_CHAR & Trim$(Npc(npcnum).AttackSay) & SEP_CHAR & Npc(npcnum).sprite & SEP_CHAR & Npc(npcnum).SpawnSecs & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & Npc(npcnum).Range & SEP_CHAR & Npc(npcnum).STR & SEP_CHAR & Npc(npcnum).def & SEP_CHAR & Npc(npcnum).Speed & SEP_CHAR & Npc(npcnum).magi & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).Exp & SEP_CHAR & Npc(npcnum).SpawnTime & SEP_CHAR & Npc(npcnum).QueteNum & SEP_CHAR & CLng(Npc(npcnum).Inv) & SEP_CHAR & CLng(Npc(npcnum).Vol) & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        packet = packet & Npc(npcnum).ItemNPC(i).chance
        packet = packet & SEP_CHAR & Npc(npcnum).ItemNPC(i).ItemNum
        packet = packet & SEP_CHAR & Npc(npcnum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    For i = 1 To MAX_NPC_SPELLS
        packet = packet & Npc(npcnum).Spell(i) & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    GetEditNpc = packet
End Function
Sub CloseWinEdit(ByVal Index As Long)
Call TextAdd(frmServer.txtText(0), "Editeur Déconnecté : " & frmServer.WinEdit(Index).RemoteHostIP, True)
frmServer.WinEdit(Index).Close
Editeur(Index).Buffer = vbNullString
Editeur(Index).Logged = False
End Sub
