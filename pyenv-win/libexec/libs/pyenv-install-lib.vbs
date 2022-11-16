Option Explicit

' Make sure to Import "pyenv-lib.vbs" before this file in a command. (for objfs/objweb variables)
' WScript.echo "kkotari: pyenv-install-lib.vbs..!"

Dim mirror
mirror = objws.Environment("Process")("PYTHON_BUILD_MIRROR_URL")
If mirror = "" Then mirror = "https://www.python.org/ftp/python"

Const SFV_FileName = 0
Const SFV_URL = 1
Const SFV_Version = 2

Const VRX_Major = 0
Const VRX_Minor = 1
Const VRX_Patch = 2
Const VRX_Release = 3
Const VRX_RelNumber = 4
Const VRX_Embed = 5
Const VRX_Target = 6
Const VRX_Web = 7
Const VRX_Ext = 8

' Version definition array from LoadVersionsXML.
Const LV_Code = 0
Const LV_FileName = 1
Const LV_URL = 2
Const LV_Embed = 3
Const LV_Target = 4
Const LV_Web = 5
Const LV_Ext = 6
Const LV_ZipRootDir = 7

' Installation parameters used for clear/extract, extension of LV.
Const IP_InstallPath = 7
Const IP_InstallFile = 8
Const IP_Quiet = 9
Const IP_Dev = 10

Dim regexVer
Dim regexFile
Set regexVer = New RegExp
Set regexFile = New RegExp
With regexVer
    .Pattern = "^(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:([a-z]+)(\d*))?$"
    .Global = True
    .IgnoreCase = True
End With
With regexFile
    .Pattern = "^python-(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:([a-z]+)(\d*))?(-embed)?(?:-(amd64|arm64|win32))?(-webinstall)?\.(exe|msi|zip)$"
    .Global = True
    .IgnoreCase = True
End With

' Normal code name
'   Fix missing "-win32", e.g. python-3.11.0.exe python-3.11.0-webinstall.exe
Function JoinCodeString(pieces)
    ' WScript.echo "kkotari: pyenv-install-lib.vbs JoinCodeString..!"
    JoinCodeString = ""
    If Len(pieces(VRX_Major))       Then JoinCodeString = JoinCodeString & pieces(VRX_Major)
    If Len(pieces(VRX_Minor))       Then JoinCodeString = JoinCodeString &"."& pieces(VRX_Minor)
    If Len(pieces(VRX_Patch))       Then JoinCodeString = JoinCodeString &"."& pieces(VRX_Patch)
    If Len(pieces(VRX_Release))     Then JoinCodeString = JoinCodeString & pieces(VRX_Release)
    If Len(pieces(VRX_RelNumber))   Then JoinCodeString = JoinCodeString & pieces(VRX_RelNumber)
    If Len(pieces(VRX_Embed))       Then JoinCodeString = JoinCodeString &"-"& pieces(VRX_Embed)
    If Len(pieces(VRX_Target))      Then JoinCodeString = JoinCodeString &"-"& pieces(VRX_Target)
    If Len(pieces(VRX_Target)) = 0  Then JoinCodeString = JoinCodeString &"-win32"
End Function

' Normal file name
Function JoinFileNameString(pieces)
    ' WScript.echo "kkotari: pyenv-install-lib.vbs JoinFileNameString..!"
    JoinFileNameString = "python-"
    If Len(pieces(VRX_Major))       Then JoinFileNameString = JoinFileNameString & pieces(VRX_Major)
    If Len(pieces(VRX_Minor))       Then JoinFileNameString = JoinFileNameString &"."& pieces(VRX_Minor)
    If Len(pieces(VRX_Patch))       Then JoinFileNameString = JoinFileNameString &"."& pieces(VRX_Patch)
    If Len(pieces(VRX_Release))     Then JoinFileNameString = JoinFileNameString & pieces(VRX_Release)
    If Len(pieces(VRX_RelNumber))   Then JoinFileNameString = JoinFileNameString & pieces(VRX_RelNumber)
    If Len(pieces(VRX_Embed))       Then JoinFileNameString = JoinFileNameString & pieces(VRX_Embed)
    If Len(pieces(VRX_Target))      Then JoinFileNameString = JoinFileNameString & pieces(VRX_Target)
    If Len(pieces(VRX_Web))         Then JoinFileNameString = JoinFileNameString & pieces(VRX_Web)
    If Len(pieces(VRX_Ext))         Then JoinFileNameString = JoinFileNameString &"."& pieces(VRX_Ext)
End Function

' Download exe file
Function DownloadFile(strUrl, strFile)
    ' WScript.echo "kkotari: pyenv-install-lib.vbs DownloadFile..!"
    On Error Resume Next

    objweb.Open "GET", strUrl, False
    If Err.Number <> 0 Then
        WScript.Echo ":: [ERROR] :: "& Err.Description
        WScript.Quit 1
    End If

    objweb.Send
    If Err.Number <> 0 Then
        WScript.Echo ":: [ERROR] :: "& Err.Description
        WScript.Quit 1
    End If
    On Error GoTo 0

    If objweb.Status <> 200 Then
        WScript.Echo ":: [ERROR] :: "& objweb.Status &" :: "& objweb.StatusText
        WScript.Quit 1
    End If

    With CreateObject("ADODB.Stream")
        .Open
        .Type = 1
        .Write objweb.responseBody
        .SaveToFile strFile, 2
        .Close
    End With
End Function

Sub clear(params)
    ' WScript.echo "kkotari: pyenv-install-lib.vbs clear..!"
    If objfs.FolderExists(params(IP_InstallPath)) Then _
        objfs.DeleteFolder params(IP_InstallPath), True

    If objfs.FileExists(params(IP_InstallFile)) Then _
        objfs.DeleteFile params(IP_InstallFile), True
End Sub

' pyenv python versions DB scheme
Dim strDBSchema
' WScript.echo "kkotari: pyenv-install-lib.vbs DBSchema..!"
strDBSchema = _
"<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">"& _
  "<xs:element name=""versions"">"& _
    "<xs:complexType>"& _
      "<xs:sequence>"& _
        "<xs:element name=""version"" maxOccurs=""unbounded"" minOccurs=""0"">"& _
          "<xs:complexType>"& _
            "<xs:sequence>"& _
              "<xs:element name=""code"" type=""xs:string""/>"& _
              "<xs:element name=""file"" type=""xs:string""/>"& _
              "<xs:element name=""URL"" type=""xs:anyURI""/>"& _
              "<xs:element name=""zipRootDir"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1""/>"& _
            "</xs:sequence>"& _
            "<xs:attribute name=""x64"" type=""xs:boolean"" default=""false""/>"& _
            "<xs:attribute name=""embed"" type=""xs:boolean"" default=""false""/>"& _
            "<xs:attribute name=""target"" type=""xs:string""/>"& _
            "<xs:attribute name=""webInstall"" type=""xs:boolean"" default=""false""/>"& _
            "<xs:attribute name=""ext"" type=""xs:string""/>"& _
            "<xs:attribute name=""msi"" type=""xs:boolean"" default=""true""/>"& _
          "</xs:complexType>"& _
        "</xs:element>"& _
      "</xs:sequence>"& _
    "</xs:complexType>"& _
  "</xs:element>"& _
"</xs:schema>"

' Load versions xml to pyenv
Function LoadVersionsXML(xmlPath)
    ' WScript.echo "kkotari: pyenv-install-lib.vbs LoadVersionsXML..!"
    Dim dbSchema
    Dim doc
    Dim schemaError
    Set LoadVersionsXML = CreateObject("Scripting.Dictionary")
    Set dbSchema = CreateObject("Msxml2.DOMDocument.6.0")
    Set doc = CreateObject("Msxml2.DOMDocument.6.0")

    If Not objfs.FileExists(xmlPath) Then Exit Function

    With dbSchema
        .validateOnParse = False
        .resolveExternals = False
        .loadXML strDBSchema
    End With

    With doc
        Set .schemas = CreateObject("Msxml2.XMLSchemaCache.6.0")
        .schemas.add "", dbSchema
        .validateOnParse = False
        .load xmlPath
        Set schemaError = .validate
    End With

    With schemaError
        If .errorCode <> 0 Then
            WScript.Echo "Parsing "& xmlPath
            WScript.Echo "Validation error in DB cache(0x"& Hex(.errorCode) & _
            ") on line "& .line &", pos "& .linepos &":"& vbCrLf & .reason
            WScript.Quit 1
        End If
    End With

    Dim versDict
    Dim version
    Dim code
    Dim zipRootDirElement, zipRootDir
    For Each version In doc.documentElement.childNodes
        code = version.getElementsByTagName("code")(0).text
        Set zipRootDirElement = version.getElementsByTagName("zipRootDir")
        If zipRootDirElement.length = 1 Then
            zipRootDir = zipRootDirElement(0).text
        Else
            zipRootDir = ""
        End If
        LoadVersionsXML.Item(code) = Array( _
            code, _
            version.getElementsByTagName("file")(0).text, _
            version.getElementsByTagName("URL")(0).text, _
            CBool(version.getAttribute("embed")), _
            version.getAttribute("target"), _
            CBool(version.getAttribute("webInstall")), _
            version.getAttribute("ext"), _
            zipRootDir _
        )
    Next
End Function

' Append xml element
Sub AppendElement(doc, parent, tag, text)
    ' WScript.echo "kkotari: pyenv-install-lib.vbs AppendElement..!"
    Dim elem
    Set elem = doc.createElement(tag)
    elem.text = text
    parent.appendChild elem
End Sub

Function LocaleIndependantCStr(booleanVal)
    If booleanVal Then
        LocaleIndependantCStr = "true"
    Else
        LocaleIndependantCStr = "false"
    End If
End Function

' Append new version to DB
Sub SaveVersionsXML(xmlPath, versArray)
    ' WScript.echo "kkotari: pyenv-install-lib.vbs SaveVersionsXML..!"
    Dim doc
    Set doc = CreateObject("Msxml2.DOMDocument.6.0")
    Set doc.documentElement = doc.createElement("versions")

    Dim versRow
    Dim versElem
    For Each versRow In versArray
        If Len(versRow(SFV_Version)(VRX_Target)) = 0 Then
            versRow(SFV_Version)(VRX_Target) = "win32"
        End If

        Set versElem = doc.createElement("version")
        doc.documentElement.appendChild versElem

        With versElem
            .setAttribute "embed",      LocaleIndependantCStr(CBool(Len(versRow(SFV_Version)(VRX_Embed))))
            .setAttribute "target",     LCase(versRow(SFV_Version)(VRX_Target))
            .setAttribute "webInstall", LocaleIndependantCStr(CBool(Len(versRow(SFV_Version)(VRX_Web))))
            .setAttribute "ext",        LCase(versRow(SFV_Version)(VRX_Ext))
        End With
        AppendElement doc, versElem, "code", JoinCodeString(versRow(SFV_Version))
        AppendElement doc, versElem, "file", versRow(0)
        AppendElement doc, versElem, "URL", versRow(1)
    Next

    ' Use SAXXMLReader/MXXMLWriter to "pretty print" the XML data.
    Dim writer
    Dim parser
    Dim outXML
    Set writer = CreateObject("Msxml2.MXXMLWriter.6.0")
    Set parser = CreateObject("Msxml2.SAXXMLReader.6.0")
    Set outXML = CreateObject("ADODB.Stream")

    With outXML
        .Open
        .Type = 1
    End With
    With writer
        .encoding = "utf-8"
        .indent = True
        .output = outXML
    End With
    With parser
        Set .contentHandler = writer
        Set .dtdHandler = writer
        Set .errorHandler = writer
        .putProperty "http://xml.org/sax/properties/declaration-handler", writer
        .putProperty "http://xml.org/sax/properties/lexical-handler", writer
        .parse doc
    End With
    With outXML
        .SaveToFile xmlpath, 2
        .Close
    End With
End Sub
