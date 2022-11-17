Option Explicit

Sub Import(importFile)
    Dim fso, libFile
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set libFile = fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) &"\"& importFile, 1)
    ExecuteGlobal libFile.ReadAll
    If Err.number <> 0 Then
        WScript.Echo "Error importing library """& importFile &"""("& Err.Number &"): "& Err.Description
        WScript.Quit 1
    End If
    libFile.Close
End Sub

Import "libs\pyenv-lib.vbs"
Import "libs\pyenv-install-lib.vbs"

Sub ShowHelp()
    ' WScript.echo "kkotari: pyenv-install.vbs..!"
    WScript.Echo "Usage: pyenv list"
    WScript.Echo ""
    WScript.Echo "  --python2              List available versions of Python 2. Default is Python 3"
    WScript.Echo "  --help                 Help, list of options allowed on pyenv list"
    WScript.Echo ""
    WScript.Quit 0
End Sub

Sub main(arg)
    ' WScript.echo "kkotari: pyenv-list.vbs Main..!"

    Dim idx
    Dim optListPython2

    optListPython2 = False

    For idx = 0 To arg.Count - 1
        Select Case arg(idx)
            Case "--help"
                ShowHelp
            Case "--python2"
                optListPython2 = True
            Case Else
        End Select
    Next

    Dim versions
    Dim version
    Set versions = LoadVersionsXML(strDBFile)
    If versions.Count = 0 Then
        WScript.Echo "pyenv-install: no definitions in local database"
        WScript.Echo
        WScript.Echo "Please update the local database cache with `pyenv update'."
        WScript.Quit 1
    End If

    Dim isPython2
    For Each version In versions.Keys
        isPython2 = Left(version, 2) = "2."
        If isPython2 = optListPython2 Then
            WScript.Echo version
        End If
    Next
End Sub

main(WScript.Arguments)
