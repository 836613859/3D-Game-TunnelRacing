Attribute VB_Name = "�̵���ֿ�"
Sub ���زֿ��뱳������()

If Dir(App.Path & "\tln.dat") = "" Then '           �ֿ�
Open App.Path & "\tln.dat" For Append As #1
For a = 1 To 4
Print #1, Chr(-a - 3000)
Tools(a) = 0
Next
 SpeedUpNum = 0
 SpeedDownNum = 0
 HPNum = 0
 ProtectNum = 0
Close #1
End If

If Dir(App.Path & "\m.dat") = "" Then '          Ǯ
Open App.Path & "\m.dat" For Append As #2
Print #2, Asc(1) & Asc(0) & Asc(0) & Asc(0)
MoneyOwn = 1000
Close #2
Form2.����XO��.Caption = "����XO�ң�" & MoneyOwn

GoTo loadForm:
End If

If Dir(App.Path & "\hs.dat") = "" Then '��Զ·��
Open App.Path & "\hs.dat" For Append As #3
Print #3, Asc(0)
Close #3
End If

'�����������������������������������òֿ��MONEY����Ϊֹ����������������������

Open App.Path & "\tln.dat" For Input As #4
For m = 1 To 4
Line Input #4, temp
Tools(m) = Asc(temp) + 3000 + m
HPNum = Tools(1)
SpeedDownNum = Tools(2)
SpeedUpNum = Tools(3)
ProtectNum = Tools(4)
Next
Close #4

Open App.Path & "\m.dat" For Input As #5
Line Input #5, mon
Close #5

Open App.Path & "\hs.dat" For Input As #6
Line Input #6, hs
Close #6

For i1 = 1 To Len(mon) Step 2
MoneyOwn = MoneyOwn & Chr(Mid(mon, i1, 2))
Next

For i2 = 1 To Len(hs) Step 2
HighScore = HighScore & Chr(Mid(hs, i2, 2))
Next

loadForm:
Set soundE = New TVSoundEngine
Set sound = New TVSoundMP3
Set bgm = New TVSoundMP3
GPath = App.Path & "\data\"

soundE.Init Form2.hwnd
sound.Load GPath & "snd/choice.mp3"
bgm.Load GPath & "snd/bgm1.mp3"
bgm.Play

Form2.Text1 = "�����빺����� "
Form2.Text2 = "�����빺����� "
Form2.Text3 = "�����빺����� "
Form2.Text4 = "�����빺����� "
Form2.HP������ = "��������" & HPNum
Form2.SDown������ = "��������" & SpeedDownNum
Form2.Sup������ = "��������" & SpeedUpNum
Form2.Prt������ = "��������" & ProtectNum

Form2.����XO��.Caption = "����XO�ң�" & MoneyOwn

End Sub

Sub ��¼ʣ��Ǯ�����()

If Dir(App.Path & "\m.dat") <> "" Then Kill App.Path & "\m.dat"
If Dir(App.Path & "\tln.dat") <> "" Then Kill App.Path & "\tln.dat"
    
    Dim mon_ey As String
    For m = 1 To Len(MoneyOwn)
    mon_ey = mon_ey & Asc(Mid(MoneyOwn, m, 1))
    Next
    
    HP_tem = Chr(HPNum - 3001)
    SPdown_tem = Chr(SpeedDownNum - 3002)
    SPup_tem = Chr(SpeedUpNum - 3003)
    Prot_tem = Chr(ProtectNum - 3004)

    Open App.Path & "\m.dat" For Append As #1
    Print #1, mon_ey
    Close #1
    
    Open App.Path & "\tln.dat" For Append As #2
    Print #2, HP_tem
    Print #2, SPdown_tem
    Print #2, SPup_tem
    Print #2, Prot_tem
    Close #2
    
    If STcam / 20 > HighScore Then
    HighScore = Int(STcam / 20)
     If Dir(App.Path & "\hs.dat") <> "" Then Kill App.Path & "\hs.dat"
     Dim high_s As String
     For h = 1 To Len(HighScore)
     high_s = high_s & Asc(Mid(HighScore, h, 1))
     Next
     Open App.Path & "\hs.dat" For Append As #3
     Print #3, high_s
     Close #3
    End If
    
End Sub
