Dim lw, a, b, ANS, arc

Do

lw = InputBox("��J���n(cm x cm)","���n����")

If lw<>"" Then

    lw = Split(lw,"*")

    a = CInt(lw(0))
    b = CInt(lw(1))
    ANS = CStr(a*b/10000*0.3025)+"�W"

    arc = MsgBox(ANS,0,"���⦨�\�I")

Else
    Exit Do

End If

Loop