set sp = createobject("SAPI.SpVoice")


set oArgs = WScript.Arguments  '�����в���

For Each text In oArgs
        sp.speak(text)
Next