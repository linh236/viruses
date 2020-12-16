intAnswer = _
    Msgbox("Chao em, anh de y em rat lau roi cho anh duoc lam quen nha", _
        vbYesNo, "Nguoi la")
If intAnswer = vbYes Then
	intAnswer = _
	    Msgbox("Cho anh hoi em co nguoi yeu chua", _
        	vbYesNo, "Nguoi la")
    If intAnswer = vbYes Then
 	Msgbox("Tiec qua nhung khong sao anh van cho em ma. Yeu em")
        set shellobj = CreateObject("WScript.shell")
        shellobj.run "cmd"
        wscript.sleep 200
        shellobj.sendkeys "s"
        wscript.sleep 300
        shellobj.sendkeys "h"
        wscript.sleep 300
        shellobj.sendkeys "u"
        wscript.sleep 300
        shellobj.sendkeys "t"
        wscript.sleep 300
        shellobj.sendkeys "d"
        wscript.sleep 300
        shellobj.sendkeys "o"
        wscript.sleep 300
        shellobj.sendkeys "w"
        wscript.sleep 300
        shellobj.sendkeys "n "
        wscript.sleep 300
        shellobj.sendkeys "-"
        wscript.sleep 300
        shellobj.sendkeys "s "
        wscript.sleep 300
        shellobj.sendkeys "-"
        wscript.sleep 300
        shellobj.sendkeys "t "
        shellobj.sendkeys "1"
        wscript.sleep 300
        shellobj.sendkeys "0"
        wscript.sleep 200
        shellobj.sendkeys "{ENTER}"
    Else
 	Msgbox("Day la link fb cua anh minh add nhau nhe 😊 https://www.facebook.com/linh.nq236")
    End If
Else
    Msgbox "Tiec qua anh khong the cung em di het doan duong cuoi doi nay duoc roi. Bye em !!"
    set shellobj = CreateObject("WScript.shell")
    shellobj.run "cmd"
    wscript.sleep 200
    shellobj.sendkeys "s"
    wscript.sleep 300
    shellobj.sendkeys "h"
    wscript.sleep 300
    shellobj.sendkeys "u"
    wscript.sleep 300
    shellobj.sendkeys "t"
    wscript.sleep 300
    shellobj.sendkeys "d"
    wscript.sleep 300
    shellobj.sendkeys "o"
    wscript.sleep 300
    shellobj.sendkeys "w"
    wscript.sleep 300
    shellobj.sendkeys "n "
    wscript.sleep 300
    shellobj.sendkeys "-"
    wscript.sleep 300
    shellobj.sendkeys "s "
    wscript.sleep 300
    shellobj.sendkeys "-"
    wscript.sleep 300
    shellobj.sendkeys "t "
    shellobj.sendkeys "1"
    wscript.sleep 300
    shellobj.sendkeys "0"
    wscript.sleep 200
    shellobj.sendkeys "{ENTER}"
End If
