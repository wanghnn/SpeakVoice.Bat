set sp = createobject("SAPI.SpVoice")


set oArgs = WScript.Arguments  'ÃüÁîĞĞ²ÎÊı

For Each text In oArgs
        sp.speak(text)
Next