Attribute VB_Name = "ModType"
Option Explicit
Public Const MAX_NPC_DROPS = 10
Public Const MAX_NPC_SPELLS = 10

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    chance As Long
End Type


Type NpcRec
    Name As String * 20
    AttackSay As String
    
    sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    STR  As Long
    def As Long
    Speed As Long
    magi As Long
    MaxHp As Long
    Exp As Long
    SpawnTime As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    QueteNum As Long
    Inv As Long
    Vol As Long
    Spell(1 To MAX_NPC_SPELLS) As Integer
End Type

Public NPC As NpcRec
Public Sub ChargerAide()
Dim f, i, j, k As Long
Dim Subs() As String

Dim Fonctions() As String

Dim texte, TempTexte As String


f = FreeFile


Open App.Path & "\ClsCommands.cls" For Input As #f
        Do Until EOF(f) = True
            Line Input #f, TempTexte
                If Mid(TempTexte, 1, 3) = "Sub" Or Mid(TempTexte, 8, 3) = "Sub" Then texte = texte & Trim(TempTexte) & vbNewLine
                If Mid(TempTexte, 1, 8) = "Function" Or Mid(TempTexte, 8, 8) = "Function" Then
                    texte = texte & Trim(TempTexte) & vbNewLine
                End If
        Loop
        
Subs = Split(texte, "Sub")

Main.List1.Clear
k = 0
For i = 0 To UBound(Subs)
j = InStr(1, Subs(i), "(")
    If j > 0 Then
        Call Main.List1.AddItem(Mid(Subs(i), 1, j - 1), k)
        Call Main.List2.AddItem(Subs(i), k)
        Main.List1.ItemData(k) = 1
        k = k + 1
    End If
Next i

Call Main.List1.AddItem("==============FONCTIONS==============", k)
Call Main.List2.AddItem("==", k)
k = k + 1
Fonctions = Split(texte, "Function")
For i = 1 To UBound(Fonctions)

j = InStr(1, Fonctions(i), "(")
    If j > 0 Then
        Call Main.List1.AddItem(Mid(Fonctions(i), 1, j - 1), k)
        Call Main.List2.AddItem(Fonctions(i), k)
        k = k + 1
    End If
Next i
Close #f
End Sub
