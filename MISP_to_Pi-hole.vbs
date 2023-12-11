Option Explicit

Dim fso, inputFile, outputFile, line, regex, matches, match
Dim blockIP

blockIP = "0.0.0.0" ' IP-адреса для блокування

Set fso = CreateObject("Scripting.FileSystemObject")
Set regex = New RegExp

regex.Pattern = "((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))|((?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z0-9][a-z0-9-]{0,61}[a-z0-9])"
regex.Global = True
regex.IgnoreCase = True

Set inputFile = fso.OpenTextFile("input.txt", 1)
Set outputFile = fso.OpenTextFile("pi-hole.txt", 2, True)

Do Until inputFile.AtEndOfStream
    line = inputFile.ReadLine
    line = Replace(line, "[", "")
    line = Replace(line, "]", "")
    Set matches = regex.Execute(line)
    If matches.Count > 0 Then
        For Each match In matches
            outputFile.WriteLine blockIP & " " & match.Value
        Next
    End If
Loop

inputFile.Close
outputFile.Close
