Set o = CreateObject("WScript.Shell")
Set env = o.Environment("Process")

home = o.ExpandEnvironmentStrings("%USERPROFILE%") & "\.lark-channel"
env("LARK_CHANNEL_HOME") = home

' ???????? PATH ?????? daemon / agent ?????
p = env("PATH")
nodeDir = o.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\nodejs"
codexBin = o.ExpandEnvironmentStrings("%USERPROFILE%") & "\.codex\bin"
gitCmd  = "C:\Program Files\Git\cmd"
If InStr(1, p, nodeDir, 1) = 0 Then p = nodeDir & ";" & p
If InStr(1, p, codexBin, 1) = 0 Then p = codexBin & ";" & p
If InStr(1, p, gitCmd, 1) = 0 Then p = gitCmd & ";" & p
env("PATH") = p

node    = nodeDir & "\node.exe"
bridge  = nodeDir & "\node_modules\lark-channel-bridge\bin\lark-channel-bridge.mjs"
logFile = home & "\bridge-autostart.log"

'-----------------------------------------------------------
' safe-restart: kill-old + wait + clear-stale-locks + start
' Handles all scenarios safely:
'   A) no old process, no stale locks (first time / after reboot)
'   B) dead old process with orphaned lock files (crash / power loss)
'   C) live old process (re-run while already running)
'-----------------------------------------------------------

' Step 1: kill any live bridge node process (taskkill /F /T = kill tree).
' No-op if none are running.
KillBridgeProcesses

' Step 2: wait up to 15s (30 x 500ms) for the process tree to fully exit.
Dim waitTries
waitTries = 0
Do While BridgeProcessRunning And waitTries < 30
    WScript.Sleep 500
    waitTries = waitTries + 1
Loop
WScript.Sleep 500

' Step 3: once process is confirmed dead, clear orphaned runtime lock files.
' Retry deletion up to 10x500ms because the OS may briefly hold file locks.
If Not BridgeProcessRunning Then
    On Error Resume Next
    ClearStaleRuntimeLocksWithRetry home
    On Error GoTo 0
End If

' Step 4: start the bridge hidden. style 0 = hidden window; False = do not wait.
cmd = "cmd /c " & node & " " & bridge & " run --profile codex --skip-check-lark-cli >> " & logFile & " 2>&1"
o.Run cmd, 0, False

Sub KillBridgeProcesses
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='node.exe' AND CommandLine LIKE '%lark-channel%'")
    For Each pr in procs
        pid = pr.ProcessId
        On Error Resume Next
        o.Run "taskkill /F /T /PID " & pid, 0, True
        On Error GoTo 0
    Next
End Sub

Function BridgeProcessRunning
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='node.exe' AND CommandLine LIKE '%lark-channel%'")
    BridgeProcessRunning = (procs.Count > 0)
End Function

Sub ClearStaleRuntimeLocksWithRetry(homeDir)
    Set fso = CreateObject("Scripting.FileSystemObject")
    locks = fso.BuildPath(homeDir, "registry\locks")
    If Not fso.FolderExists(locks) Then Exit Sub
    Dim retries, done, pf, af
    retries = 0
    done = False
    Do While Not done And retries < 10
        pf = fso.BuildPath(locks, "profile")
        af = fso.BuildPath(locks, "app")
        On Error Resume Next
        ' Delete all files + subfolders in profile/
        If fso.FolderExists(pf) Then
            fso.DeleteFile fso.BuildPath(pf, "*.*"), True
            For Each sf In fso.GetFolder(pf).SubFolders
                sf.Delete True
            Next
        End If
        ' Delete all files + subfolders in app/
        If fso.FolderExists(af) Then
            fso.DeleteFile fso.BuildPath(af, "*.*"), True
            For Each sf In fso.GetFolder(af).SubFolders
                sf.Delete True
            Next
        End If
        On Error GoTo 0
        ' Verify both files AND subfolders are gone
        If (Not fso.FolderExists(pf) Or _
            (fso.GetFolder(pf).Files.Count = 0 And fso.GetFolder(pf).SubFolders.Count = 0)) _
           And (Not fso.FolderExists(af) Or _
            (fso.GetFolder(af).Files.Count = 0 And fso.GetFolder(af).SubFolders.Count = 0)) Then
            done = True
        Else
            WScript.Sleep 500
            retries = retries + 1
        End If
    Loop
End Sub
