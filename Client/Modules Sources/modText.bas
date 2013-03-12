Attribute VB_Name = "modText"
Option Explicit

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Const Quote As String = """"

Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15

Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = Green
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = White
Public Const EmoteColor As Byte = White
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = White
Public Const WhoColor As Byte = Grey
Public Const JoinLeftColor As Byte = Grey
Public Const NpcColor As Byte = White
Public Const AlertColor As Byte = White
Public Const NewMapColor As Byte = Grey

Type RGB
    r As Byte
    g As Byte
    B As Byte
End Type

Public MsgRgb() As RGB
Public MsgRgb2() As RGB ' 0=Playermsg 1=adminmsg 2=serveurmsg 3=privatemsg 4=guilde
Public MaxColor As Byte

Public TexthDC As Long
Public GameFont As Long

Public Sub SetFont(ByVal Font As String, ByVal size As Byte)
    GameFont = CreateFont(size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, X + 1, Y + 0, Text, Len(Text))
    Call TextOut(hDC, X + 0, Y + 1, Text, Len(Text))
    Call TextOut(hDC, X - 1, Y - 0, Text, Len(Text))
    Call TextOut(hDC, X - 0, Y - 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub
Public Sub DrawPlayerNameText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, X + 1, Y + 0, Text, Len(Text))
    Call TextOut(hDC, X + 0, Y + 1, Text, Len(Text))
    Call TextOut(hDC, X - 1, Y - 0, Text, Len(Text))
    Call TextOut(hDC, X - 0, Y - 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub

Public Sub DrawTextInter(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, vbBlack)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Long, Optional ByVal Genre As Byte)
Dim s As String
Dim C As Long
Dim t As Long
Dim i As Long
Dim z As Long
Dim r, g, B As Long
On Error Resume Next
t = 0
'frmMirage.PicScreen2.Picture = frmMirage.picScreen.Image
With FrmMirage
If OngletActif <> 0 Then .Onglet(0).ForeColor = vbRed
    With .RTBChat(0)
    
Select Case Genre

' 0=Playermsg 1=adminmsg 2=serveurmsg 3=privatemsg 4=guilde

Case 0 'messages joueurs/map
    r = MsgRgb(0).r
    g = MsgRgb(0).g
    B = MsgRgb(0).B

    t = Len(.Text)
    C = InStr(Msg, ":")
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = C + 1
    .SelColor = RGB(0, 0, 0)
    .SelBold = True
    .SelStart = t + C + 1
    .SelLength = Len(Msg) - C + 2
    .SelColor = RGB(r, g, B)
    
Case 1 'Messages Interne
    r = MsgRgb(2).r
    g = MsgRgb(2).g
    B = MsgRgb(2).B
    
    t = Len(.Text)
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = Len(Msg) + 2
     .SelItalic = False
    .SelColor = RGB(r, g, B)
    

    
Case 2 'messages Admin

    r = MsgRgb(1).r
    g = MsgRgb(1).g
    B = MsgRgb(1).B
    
    
    t = Len(.Text)
    C = InStr(Msg, ")")
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = C + 2
     .SelItalic = False
    .SelColor = RGB(205, 0, 0)
    .SelBold = True
    .SelStart = Len(.Text) - (Len(Msg) - C)
    .SelLength = Len(Msg) - C
    .SelColor = RGB(r, g, B)
    
Case 3 ' Message priv�
    r = MsgRgb(3).r
    g = MsgRgb(3).g
    B = MsgRgb(3).B

    t = Len(.Text)
    C = InStr(Msg, ")")
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = C + 2
    .SelItalic = False
    .SelColor = RGB(0, 0, 0)
    .SelBold = True
    .SelStart = Len(.Text) - (Len(Msg) - C)
    .SelLength = Len(Msg) - C
    .SelColor = RGB(r, g, B)
    
Case 4 ' Message Guilde
    r = MsgRgb(4).r
    g = MsgRgb(4).g
    B = MsgRgb(4).B

    t = Len(.Text)
    C = InStr(Msg, ")")
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = C + 2
     .SelItalic = False
    .SelColor = RGB(0, 0, 0)
    .SelBold = True
    .SelStart = Len(.Text) - (Len(Msg) - C)
    .SelLength = Len(Msg) - C
    .SelColor = RGB(r, g, B)
    
Case 5 'Emotes
    r = MsgRgb(5).r
    g = MsgRgb(5).g
    B = MsgRgb(5).B
    
    t = Len(.Text)
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = Len(Msg) + 2
    .SelColor = RGB(r, g, B)
    .SelItalic = True
End Select
End With




For i = 1 To .RTBChat.count - 1
With .RTBChat(i)
If RTB(i).Canal(Genre) = True Then 'PUTAIN DE CHIERIE
If i <> OngletActif Then FrmMirage.Onglet(i).ForeColor = vbRed
Select Case Genre

' 0=Playermsg 1=adminmsg 2=serveurmsg 3=privatemsg 4=guilde

Case 0 'messages joueurs/map
    r = MsgRgb(0).r
    g = MsgRgb(0).g
    B = MsgRgb(0).B

    t = Len(.Text)
    C = InStr(Msg, ":")
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = C + 1
    .SelColor = RGB(0, 0, 0)
    .SelBold = True
    .SelStart = t + C + 1
    .SelLength = Len(Msg) - C + 2
    .SelColor = RGB(r, g, B)
    
Case 1 'Messages Interne
    r = MsgRgb(2).r
    g = MsgRgb(2).g
    B = MsgRgb(2).B
    
    t = Len(.Text)
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = Len(Msg) + 2
     .SelItalic = False
    .SelColor = RGB(r, g, B)
    

    
Case 2 'messages Admin

    r = MsgRgb(1).r
    g = MsgRgb(1).g
    B = MsgRgb(1).B
    
    
    t = Len(.Text)
    C = InStr(Msg, ")")
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = C + 2
     .SelItalic = False
    .SelColor = RGB(205, 0, 0)
    .SelBold = True
    .SelStart = Len(.Text) - (Len(Msg) - C)
    .SelLength = Len(Msg) - C
    .SelColor = RGB(r, g, B)
    
Case 3 ' Message priv�
    r = MsgRgb(3).r
    g = MsgRgb(3).g
    B = MsgRgb(3).B

    t = Len(.Text)
    C = InStr(Msg, ")")
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = C + 2
    .SelItalic = False
    .SelColor = RGB(0, 0, 0)
    .SelBold = True
    .SelStart = Len(.Text) - (Len(Msg) - C)
    .SelLength = Len(Msg) - C
    .SelColor = RGB(r, g, B)
    
Case 4 ' Message Guilde
    r = MsgRgb(4).r
    g = MsgRgb(4).g
    B = MsgRgb(4).B

    t = Len(.Text)
    C = InStr(Msg, ")")
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = C + 2
     .SelItalic = False
    .SelColor = RGB(0, 0, 0)
    .SelBold = True
    .SelStart = Len(.Text) - (Len(Msg) - C)
    .SelLength = Len(Msg) - C
    .SelColor = RGB(r, g, B)
    
Case 5 'Emotes
    r = MsgRgb(5).r
    g = MsgRgb(5).g
    B = MsgRgb(5).B
    
    t = Len(.Text)
    .SelStart = t
    .SelText = vbNewLine & Msg
    .SelStart = t
    .SelLength = Len(Msg) + 2
    .SelColor = RGB(r, g, B)
    .SelItalic = True
    


End Select
End If
End With
Next i

End With
Exit Sub

 '      For i = 1 To MAX_BLT_LINE
 '           If t = 0 Then
 '               If BattlePMsg(i).Index <= 0 Then
 '                   BattlePMsg(i).Index = 1
 '                   BattlePMsg(i).Msg = Msg
 '                   BattlePMsg(i).Color = Color
 '                   BattlePMsg(i).Time = GetTickCount
 '                   BattlePMsg(i).Done = 1
 '                   BattlePMsg(i).y = 0
 '                   Exit Sub
 '               Else
 '                   BattlePMsg(i).y = BattlePMsg(i).y - 15
 '               End If
 '           Else
 '               If BattleMMsg(i).Index <= 0 Then
 '                   BattleMMsg(i).Index = 1
 '                   BattleMMsg(i).Msg = Msg
 '                   BattleMMsg(i).Color = Color
  '                  BattleMMsg(i).Time = GetTickCount
 '                   BattleMMsg(i).Done = 1
 '                   BattleMMsg(i).y = 0
 '                   Exit Sub
 '               Else
 '                   BattleMMsg(i).y = BattleMMsg(i).y - 15
 '               End If
 '           End If
 '       Next i
 '
 '       z = 1
 '       If t = 0 Then
 '           For i = 1 To MAX_BLT_LINE
 '               If i < MAX_BLT_LINE Then If BattlePMsg(i).y < BattlePMsg(i + 1).y Then z = i Else If BattlePMsg(i).y < BattlePMsg(1).y Then z = i
 '           Next i
 '
 '           BattlePMsg(z).Index = 1
 '           BattlePMsg(z).Msg = Msg
 '           BattlePMsg(z).Color = Color
 '           BattlePMsg(z).Time = GetTickCount
 '           BattlePMsg(z).Done = 1
 '           BattlePMsg(z).y = 0
 '       Else
 '           For i = 1 To MAX_BLT_LINE
 '               If i < MAX_BLT_LINE Then If BattleMMsg(i).y < BattleMMsg(i + 1).y Then z = i Else If BattleMMsg(i).y < BattleMMsg(1).y Then z = i
 '           Next i
 '
 '           BattleMMsg(z).Index = 1
 '           BattleMMsg(z).Msg = Msg
 '           BattleMMsg(z).Color = Color
 '           BattleMMsg(z).Time = GetTickCount
 '           BattleMMsg(z).Done = 1
 '           BattleMMsg(z).y = 0
 '       End If
 '       Exit Sub
End Sub


Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then Txt.Text = Txt.Text + Msg + vbCrLf Else Txt.Text = Txt.Text + Msg
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Function Parse(ByVal num As Long, ByVal data As String)
Dim i As Long
Dim n As Long
Dim sChar As Long
Dim eChar As Long

    n = 0
    sChar = 1
    
    For i = 1 To Len(data)
        If Mid$(data, i, 1) = SEP_CHAR Then
            If n = num Then
                eChar = i
                Parse = Mid$(data, sChar, eChar - sChar)
                Exit For
            End If
            
            sChar = i + 1
            n = n + 1
        End If
    Next i
    
End Function
