Sub ShowVoiceList()
    Set sv = CreateObject("sapi.spvoice")  '获得语音引擎集合

    Dim s

		Dim sn
    sn=1
		
    For Each v In sv.GetVoices
        s=s & "(" & CStr(sn) & ") " & v.GetDescription & vbcrlf 
        s=s & v.id & vbcrlf
        
        sn = sn+1
    Next

    MsgBox s

End Sub


ShowVoiceList()


'Dim s
'set s = "hi"



'MsgBox s