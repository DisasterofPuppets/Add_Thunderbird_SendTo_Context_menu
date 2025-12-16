Set objShell = CreateObject("WScript.Shell")
Set objArgs  = WScript.Arguments

If objArgs.Count = 0 Then WScript.Quit

tbird = "C:\Program Files\Mozilla Thunderbird\thunderbird.exe"

' Build a comma-separated attachment list: attachment='file:///path1,file:///path2'
attachments = ""
For i = 0 To objArgs.Count - 1
    filePath = objArgs(i)
    ' Convert to file:/// URL and percent-encode spaces
    fileUrl = "file:///" & Replace(filePath, "\", "/")
    fileUrl = Replace(fileUrl, " ", "%20")
    attachments = attachments & fileUrl
    If i < objArgs.Count - 1 Then attachments = attachments & ","
Next

command = """" & tbird & """ -compose " & """attachment='" & attachments & "'"""
' 0 = hidden window; False = do not wait
objShell.Run command, 0, False