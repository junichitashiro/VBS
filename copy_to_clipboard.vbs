Option Explicit

Dim ID
Dim PW

ID = "abcde"
PW = "password"

CopyToClip ID
Call MsgBox("ID copied to clipboard.",vbSystemModal)
CopyToClip PW

Public Sub CopyToClip(ByVal str)
    Dim cmd
    cmd = "cmd /c ""echo " & str & "| clip"""
    CreateObject("WScript.Shell").Run cmd, 0
End Sub
