Dim lw, a, b, ANS, arc

Do

lw = InputBox("輸入面積(cm x cm)","面積換算")

If lw<>"" Then

    lw = Split(lw,"*")

    a = CInt(lw(0))
    b = CInt(lw(1))
    ANS = CStr(a*b/10000*0.3025)+"坪"

    arc = MsgBox(ANS,0,"換算成功！")

Else
    Exit Do

End If

Loop