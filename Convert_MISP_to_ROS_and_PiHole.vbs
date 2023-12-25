Option Explicit

Dim fso, inputFile, line, regex, matches, match
Dim mikroTikFile, piholeFile
Dim blockIP

blockIP = "0.0.0.0" ' IP-адреса для блокування у Pi-hole

Set fso = CreateObject("Scripting.FileSystemObject")
Set regex = New RegExp

' Регулярний вираз для виявлення IP-адрес та доменних імен
regex.Pattern = "((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))|((?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z0-9][a-z0-9-]{0,61}[a-z0-9])"
regex.Global = True
regex.IgnoreCase = True

' Відкриття файлів для читання та запису
Set inputFile = fso.OpenTextFile("input.txt", 1)
Set mikroTikFile = fso.OpenTextFile("For_MikroTik.txt", 2, True)
Set piholeFile = fso.OpenTextFile("For_PiHole.txt", 2, True)

' Запис заголовка для файлу ROS 7
mikroTikFile.WriteLine "/ip firewall address-list"

Do Until inputFile.AtEndOfStream
    line = inputFile.ReadLine
    line = Replace(line, "[", "")
    line = Replace(line, "]", "")
    Set matches = regex.Execute(line)

    If matches.Count > 0 Then
        For Each match In matches
            If IsIP(match.Value) Then
                ' Запис IP-адреси у форматі ROS 7
                mikroTikFile.WriteLine "add list=""Site Block"" address=" & match.Value
            Else
                ' Запис доменного імені у форматі Pi-hole
                piholeFile.WriteLine blockIP & " " & match.Value
            End If
        Next
    End If
Loop

inputFile.Close
mikroTikFile.Close
piholeFile.Close

' Функція для перевірки, чи є рядок IP-адресою
Function IsIP(str)
    Dim ipRegex
    Set ipRegex = New RegExp
    ipRegex.Pattern = "^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    IsIP = ipRegex.Test(str)
End Function
