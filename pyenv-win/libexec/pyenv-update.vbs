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

WScript.Echo ":: [Info] ::  Mirror: " & mirror

Sub ShowHelp()
    WScript.Echo "Usage: pyenv update [--ignore]"
    WScript.Echo
    WScript.Echo "  --ignore  Ignores any HTTP/VBScript errors that occur during downloads."
    WScript.Echo
    WScript.Echo "Updates the internal database of python installer URL's."
    WScript.Echo
    WScript.Quit 0
End Sub

Sub EnsureBaseURL(ByRef html, ByVal URL)
    Dim head
    Dim base

    Set head = html.getElementsByTagName("head")(0)
    If head Is Nothing Then
        Set head = html.createElement("head")
        html.insertBefore html.body, head
    End If

    Set base = head.getElementsByTagName("base")(0)
    If base Is Nothing Then
        If Len(URL) And Right(URL, 1) <> "/" Then URL = URL &"/"
        Set base = html.createElement("base")
        base.href = URL
        head.appendChild base
    End If
End Sub

Function CollectionToArray(collection) _
    Dim i
    Dim arr()
    ReDim arr(collection.Count-1)
    For i = 0 To collection.Count-1
        If IsObject(collection.Item(i)) Then
            Set arr(i) = collection.Item(i)
        Else
            arr(i) = collection.Item(i)
        End If
    Next
    CollectionToArray = arr
End Function

Function CopyDictionary(dict)
    Dim key
    Set CopyDictionary = CreateObject("Scripting.Dictionary")
    For Each key In dict.Keys
        CopyDictionary.Add key, dict(key)
    Next
End Function

Sub UpdateDictionary(dict1, dict2)
    Dim key
    For Each key In dict2.Keys
        If IsObject(dict2(key)) Then
            Set dict1(key) = dict2(key)
        Else
            dict1(key) = dict2(key)
        End If
    Next
End Sub

Function ScanForVersions(URL, optIgnore, ByRef pageCount)
    Dim objHTML
    Set objHTML = CreateObject("htmlfile")
    Set ScanForVersions = CreateObject("Scripting.Dictionary")

    With objweb
        .open "GET", URL, False
        On Error Resume Next
        .send
        If Err.number <> 0 Then
            WScript.Echo "HTTP Error downloading from mirror page """& URL &""""& vbCrLf &"Error(0x"& Hex(Err.Number) &"): "& Err.Description
            If optIgnore Then Exit Function
            WScript.Quit 1
        End If
        On Error GoTo 0
        If .status <> 200 Then
            WScript.Echo "HTTP Error downloading from mirror page """& URL &""""& vbCrLf &"Error("& .status &"): "& .statusText
            If optIgnore Then Exit Function
            WScript.Quit 1
        End If

        objHTML.write .responseText
        pageCount = pageCount + 1
    End With
    EnsureBaseURL objHTML, URL

    Dim link
    Dim fileName
    Dim matches
    Dim match
    Dim major, minor, patch, rel
    For Each link In objHTML.links
        fileName = Trim(link.innerText)
        Set matches = regexFile.Execute(fileName)
        If matches.Count = 1 Then
            ' WScript.Echo "FileName " &fileName
            match = CollectionToArray(matches(0).SubMatches)
            ' Save as a dictionary entry with Key/Value as:
            '  -Key: [filename]
            '  -Value: Array([filename], [url], Array([regex submatches]))
            ScanForVersions.Add fileName, Array(fileName, link.href, match)
        End If
    Next
End Function

' Test if ver1 < ver2
Function SymanticCompare(ver1, ver2)
    Dim comp1, comp2

    ' Major
    comp1 = ver1(VRX_Major)
    comp2 = ver2(VRX_Major)
    If Len(comp1) = 0 Then comp1 = 0: Else comp1 = CLng(comp1)
    If Len(comp2) = 0 Then comp2 = 0: Else comp2 = CLng(comp2)
    SymanticCompare = comp1 < comp2
    If comp1 <> comp2 Then Exit Function

    ' Minor
    comp1 = ver1(VRX_Minor)
    comp2 = ver2(VRX_Minor)
    If Len(comp1) = 0 Then comp1 = 0: Else comp1 = CLng(comp1)
    If Len(comp2) = 0 Then comp2 = 0: Else comp2 = CLng(comp2)
    SymanticCompare = comp1 < comp2
    If comp1 <> comp2 Then Exit Function

    ' Patch
    comp1 = ver1(VRX_Patch)
    comp2 = ver2(VRX_Patch)
    If Len(comp1) = 0 Then comp1 = 0: Else comp1 = CLng(comp1)
    If Len(comp2) = 0 Then comp2 = 0: Else comp2 = CLng(comp2)
    SymanticCompare = comp1 < comp2
    If comp1 <> comp2 Then Exit Function

    ' Release
    comp1 = ver1(VRX_Release)
    comp2 = ver2(VRX_Release)
    If Len(comp1) = 0 And Len(comp2) Then
        SymanticCompare = False
        Exit Function
    ElseIf Len(comp1) And Len(comp2) = 0 Then
        SymanticCompare = True
        Exit Function
    Else
        SymanticCompare = comp1 < comp2
    End If
    If comp1 <> comp2 Then Exit Function

    ' Release Number
    comp1 = ver1(VRX_RelNumber)
    comp2 = ver2(VRX_RelNumber)
    If Len(comp1) = 0 Then comp1 = 0: Else comp1 = CLng(comp1)
    If Len(comp2) = 0 Then comp2 = 0: Else comp2 = CLng(comp2)
    SymanticCompare = comp1 < comp2
    If comp1 <> comp2 Then Exit Function

    ' embeded or not
    comp1 = ver1(VRX_Embed)
    comp2 = ver2(VRX_Embed)
    SymanticCompare = comp1 < comp2
    If comp1 <> comp2 Then Exit Function

    ' target: amd64 arm64 win32
    comp1 = ver1(VRX_Target)
    comp2 = ver2(VRX_Target)
    SymanticCompare = comp1 < comp2
    If comp1 <> comp2 Then Exit Function

    ' webinstall
    comp1 = ver1(VRX_Web)
    comp2 = ver2(VRX_Web)
    SymanticCompare = comp1 < comp2
    If comp1 <> comp2 Then Exit Function

    ' ext: exe msi
    comp1 = ver1(VRX_Ext)
    comp2 = ver2(VRX_Ext)
    SymanticCompare = comp1 < comp2
    If comp1 <> comp2 Then Exit Function
End Function

' Modified from code by "Reverend Jim" at:
' https://www.daniweb.com/programming/code/515601/vbscript-implementation-of-quicksort
Sub SymanticQuickSort(arr, arrMin, arrMax)
    Dim middle  ' value of the element in the middle of the range
    Dim swap    ' temporary item for the swapping of two elements
    Dim arrFrst ' index of the first element in the range to check
    Dim arrLast ' index of the last element in the range to check
    Dim arrMid  ' index of the element in the middle of the range
    If arrMax <= arrMin Then Exit Sub

    ' Start the checks at the lower and upper limits of the Array
    arrFrst = arrMin
    arrLast = arrMax

    ' Find the midpoint of the region to sort and the value of that element
    arrMid = (arrMin + arrMax) \ 2
    middle = arr(arrMid)
    Do While (arrFrst <= arrLast)
        ' Find the first element > the element at the midpoint
        Do While SymanticCompare(arr(arrFrst)(SFV_Version), middle(SFV_Version))
            arrFrst = arrFrst + 1
            If arrFrst = arrMax Then Exit Do
        Loop

        ' Find the last element < the element at the midpoint
        Do While SymanticCompare(middle(SFV_Version), arr(arrLast)(SFV_Version))
            arrLast = arrLast - 1
            If arrLast = arrMin Then Exit Do
        Loop

        ' Pivot the two elements around the midpoint if they are out of order
        If (arrFrst <= arrLast) Then
            swap = arr(arrFrst)
            arr(arrFrst) = arr(arrLast)
            arr(arrLast) = swap
            arrFrst = arrFrst + 1
            arrLast = arrLast - 1
        End If
    Loop

    ' Sort sub-regions (recurse) if necessary
    If arrMin  < arrLast Then SymanticQuickSort arr, arrMin,  arrLast
    If arrFrst < arrMax  Then SymanticQuickSort arr, arrFrst, arrMax
End Sub

Sub main(arg)
    Dim idx
    Dim optIgnore
    Dim optPython2
    Dim optPython3
    Dim optEmbed
    Dim optTargets
    Dim optWebInstall

    optIgnore = False
    optPython2 = False
    optPython3 = ""
    optEmbed = False
    optTargets = Array(HostTarget)
    optWebInstall = True

    For idx = 0 To arg.Count - 1
        Select Case arg(idx)
            Case "--help"
                ShowHelp
            Case "--ignore"
                optIgnore = True
            Case "--python2"
                optPython2 = True
            Case "--python3"
                idx = idx + 1
                optPython3 = arg(idx)
            Case "--embed"
                optEmbed = True
            Case "--target"
                idx = idx + 1
                optTargets = Split(arg(idx), ",")
            Case "--no-web-install"
                optWebInstall = False
            Case Else
                WScript.Echo "Unknown option "& arg(idx)
                WScript.Quit 1
        End Select
    Next

    Dim objHTML
    Dim pageCount
    Set objHTML = CreateObject("htmlfile")
    pageCount = 0

    With objweb
        On Error Resume Next
        .Open "GET", mirror, False
        If Err.number <> 0 Then
            WScript.Echo "HTTP Error downloading from mirror """& mirror &""""& vbCrLf &"Error(0x"& Hex(Err.number) &"): "& Err.Description
            If optIgnore Then Exit Sub
            WScript.Quit 1
        End If

        .Send
        If Err.number <> 0 Then
            WScript.Echo "HTTP Error downloading from mirror """& mirror &""""& vbCrLf &"Error(0x"& Hex(Err.number) &"): "& Err.Description
            If optIgnore Then Exit Sub
            WScript.Quit 1
        End If
        On Error GoTo 0

        If .Status <> 200 Then
            WScript.Echo "HTTP Error downloading from mirror """& mirror &""""& vbCrLf &"Error("& .Status &"): "& .StatusText
            If optIgnore Then Exit Sub
            WScript.Quit 1
        End If

        objHTML.write .responseText
        pageCount = pageCount + 1
    End With
    EnsureBaseURL objHTML, mirror

    Dim link
    Dim version
    Dim matches
    Dim match
    Dim installers1
    Set installers1 = CreateObject("Scripting.Dictionary")
    For Each link In objHTML.links
    Do
        version = objfs.GetFileName(link.pathname)
        Set matches = regexVer.Execute(version)
        If matches.Count = 1 Then
            ' WScript.Echo "Link " &link.href
            match = CollectionToArray(matches(0).SubMatches)

            If optPython2 Then
                ' Ignore Python < 2.4, Wise Installer's command line is unusable.
                If match(0) = "2" And CLng(match(1)) < CLng(4) Then
                    WScript.Echo "Skip " &version
                    Exit Do
                End If
            ElseIf match(0) = "2" Then
                WScript.Echo "Skip " &version
                Exit Do
            End If

            If match(0) = "3" And optPython3 <> "" Then
                If CLng(match(1)) < CLng(optPython3) Then
                    WScript.Echo "Skip " &version
                    Exit Do
                End If
            End If

            WScript.Echo "Add " &version
            UpdateDictionary installers1, ScanForVersions(link.href, optIgnore, pageCount)
        End If
    Loop While False ' Workaround for "Continue"
    Next

    Dim fileName
    Dim versPieces
    Dim installers2
    Dim target
    Dim targetMatch
    Dim fileAlternative
    Set installers2 = CopyDictionary(installers1) ' Use a copy because "For Each" and .Remove don't play nice together.

    For Each fileName In installers1.Keys()
    Do
        ' Array(
        '     [filename],
        '     [url],
        '     Array(
        '         [major], [minor], [path], [rel], [rel_num],
        '         [embed], [target], [webinstall], [ext]
        '     )
        ' )
        versPieces = installers1(fileName)(SFV_Version)

        targetMatch = False
        For Each target In optTargets
            If target = versPieces(VRX_Target) Then
                targetMatch = True
                Exit For
            End If
        Next
        If Not targetMatch Then
            WScript.Echo "Ignore " &fileName
            installers2.Remove fileName
            Exit Do
        End If

        ' Remove any duplicate versions that have the web installer (it's prefered)
        If Len(versPieces(VRX_Web)) Then
            fileAlternative = JoinFileNameString(Array( _
                versPieces(VRX_Major), _
                versPieces(VRX_Minor), _
                versPieces(VRX_Patch), _
                versPieces(VRX_Release), _
                versPieces(VRX_RelNumber), _
                versPieces(VRX_Embed), _
                versPieces(VRX_Target), _
                Empty, _
                versPieces(VRX_Ext) _
            ))
            If installers2.Exists(fileAlternative) Then
                If optWebInstall Then
                    WScript.Echo "Ignore " &fileName
                    installers2.Remove fileName
                Else
                    WScript.Echo "Ignore " &fileAlternative
                    installers2.Remove fileAlternative
                End If
                Exit Do
            End If
        End If
    Loop While False ' Workaround for "Continue"
    Next

    ' Now sort by semantic version and save
    Dim installArr
    installArr = installers2.Items
    SymanticQuickSort installArr, LBound(installArr), UBound(installArr)
    SaveVersionsXML strDBFile, installArr
    WScript.Echo ":: [Info] ::  Scanned "& pageCount &" pages and found "& installers2.Count &" installers."

End Sub

main(WScript.Arguments)
