Attribute VB_Name = "ModPNJ"
Option Explicit
Public Sub SetPlayerFlag(ByVal Index As Long, ByVal Flag As Long, ByVal value As Long)
Player(Index).Char(Player(Index).CharNum).Flag(Flag) = value
End Sub
Public Function GetPlayerFlag(ByVal Index As Long, ByVal Flag As Long)
GetPlayerFlag = Player(Index).Char(Player(Index).CharNum).Flag(Flag)
End Function
Public Sub PlayerTalk(ByVal Index As Long, ByVal Msg As String, ByVal Targetnum As Long, ByVal targetype As Long)
Dim Parse() As String
Dim i As Long
Dim npcnum As Long

npcnum = MapNpc(Player(Index).Char(Player(Index).CharNum).Map, Targetnum).Num


'découpe du msg en mots
Parse = Split(Msg, " ")
'étude de chaque mot
On Error GoTo suite ' si le script ne trouve rien, on change de mot
For i = 0 To UBound(Parse) 'pour chaque mot
    Parse(i) = LCase(Trim(Parse(i))) 'on le met en minuscules
    Call frmServer.NPCScript.Run("NPC" & npcnum, Index, Parse(i))
    'MsgBox frmServer.NPCScript.Modules.Count
    'frmServer.NPCScript.Run "NPC" & npcnum, Index & "," & Parse(i)
    Exit Sub
suite:
Next i
Call MapPlayerMsg(Index, Err.Description, 0)
End Sub
