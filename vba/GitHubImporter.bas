' ============================================================
' GitHub Private Repo - VBA Module Importer
' Requires: Tools > References > "Microsoft Scripting Runtime"
'           and "Microsoft Visual Basic for Applications Extensibility 5.3"
' ============================================================

Private Const GITHUB_TOKEN  As String = "ghp_YourPersonalAccessTokenHere"
Private Const GITHUB_OWNER  As String = "your-username"
Private Const GITHUB_REPO   As String = "your-repo-name"
Private Const GITHUB_BRANCH As String = "main"

' ------------------------------------------------------------
' Entry point: define which files to import, then call this
' ------------------------------------------------------------
Public Sub ImportModulesFromGitHub()
    Dim files() As String
    files = Split("src/Module1.bas,src/Module2.bas,src/MyClass.cls,src/Sheet1.cls,src/ThisWorkbook.cls", ",")

    Dim f As Variant
    For Each f In files
        ImportSingleModule CStr(f)
    Next f

    MsgBox "Import complete!", vbInformation
End Sub

' ------------------------------------------------------------
' Fetches one file from GitHub and imports or injects it
' ------------------------------------------------------------
Private Sub ImportSingleModule(repoFilePath As String)
    Dim url As String
    url = "https://api.github.com/repos/" & GITHUB_OWNER & "/" & GITHUB_REPO & _
          "/contents/" & repoFilePath & "?ref=" & GITHUB_BRANCH

    ' --- 1. Fetch JSON from GitHub API ---
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "token " & GITHUB_TOKEN
    http.setRequestHeader "Accept", "application/vnd.github.v3+json"
    http.setRequestHeader "User-Agent", "Excel-VBA"
    http.Send

    If http.Status <> 200 Then
        MsgBox "GitHub API error " & http.Status & " for: " & repoFilePath & vbLf & http.responseText, vbCritical
        Exit Sub
    End If

    ' --- 2. Decode Base64 content ---
    Dim b64 As String
    b64 = ExtractJsonValue(http.responseText, "content")
    b64 = Join(Split(b64, "\n"), "")

    Dim tempPath As String
    tempPath = Environ("TEMP") & "\" & GetFilenameFromPath(repoFilePath)
    WriteBase64ToFile b64, tempPath

    ' --- 3. Route to correct import strategy ---
    Dim compName As String
    compName = GetModuleNameFromFile(tempPath)

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    If IsDocumentModule(vbProj, compName) Then
        InjectCodeIntoDocumentModule vbProj, compName, tempPath
    Else
        On Error Resume Next
        vbProj.VBComponents.Remove vbProj.VBComponents(compName)
        On Error GoTo 0
        vbProj.VBComponents.Import tempPath
    End If

    Kill tempPath
End Sub

' ------------------------------------------------------------
' Returns True if the named component is a document module
' (vbext_ct_Document = 100: ThisWorkbook, Sheet, UserForm code)
' ------------------------------------------------------------
Private Function IsDocumentModule(vbProj As Object, compName As String) As Boolean
    Const vbext_ct_Document As Integer = 100
    On Error Resume Next
    Dim comp As Object
    Set comp = vbProj.VBComponents(compName)
    On Error GoTo 0
    If comp Is Nothing Then
        IsDocumentModule = False
    Else
        IsDocumentModule = (comp.Type = vbext_ct_Document)
    End If
End Function

' ------------------------------------------------------------
' Reads a .cls file from disk, strips the VBComponent header,
' then replaces the target document module's code entirely
' ------------------------------------------------------------
Private Sub InjectCodeIntoDocumentModule(vbProj As Object, compName As String, filePath As String)
    ' Read file contents
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    Dim fileContents As String
    fileContents = ts.ReadAll
    ts.Close

    ' Strip .cls header (everything up to and including "Attribute VB_Name = ..." block)
    ' Headers end before the first non-Attribute line
    Dim lines() As String
    lines = Split(fileContents, vbLf)

    Dim codeStart As Long
    codeStart = 0
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim trimmed As String
        trimmed = Trim(Replace(lines(i), vbCr, ""))
        ' Skip blank lines and Attribute declarations at the top
        If Len(trimmed) > 0 And Left(trimmed, 9) <> "Attribute" Then
            codeStart = i
            Exit For
        End If
    Next i

    ' Rebuild code string from codeStart onwards
    Dim codelines() As String
    ReDim codelines(UBound(lines) - codeStart)
    For i = codeStart To UBound(lines)
        codelines(i - codeStart) = lines(i)
    Next i
    Dim cleanCode As String
    cleanCode = Join(codelines, vbLf)

    ' Replace the document module's code
    Dim comp As Object
    Set comp = vbProj.VBComponents(compName)
    With comp.CodeModule
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, cleanCode
    End With
End Sub

' ------------------------------------------------------------
' Minimal JSON value extractor for GitHub API "content" field
' ------------------------------------------------------------
Private Function ExtractJsonValue(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """:"
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then Exit Function

    pos = pos + Len(pattern)
    Do While Mid(json, pos, 1) = " " Or Mid(json, pos, 1) = Chr(10) Or Mid(json, pos, 1) = Chr(13)
        pos = pos + 1
    Loop
    If Mid(json, pos, 1) = """" Then pos = pos + 1

    Dim result As String
    Dim c As String
    Do
        c = Mid(json, pos, 1)
        If c = """" Then Exit Do
        result = result & c
        pos = pos + 1
    Loop

    ExtractJsonValue = result
End Function

' ------------------------------------------------------------
' Decodes Base64 and writes to a temp file via ADO Stream
' ------------------------------------------------------------
Private Sub WriteBase64ToFile(b64 As String, filePath As String)
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Dim node As Object
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = b64

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write node.nodeTypedValue
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

' ------------------------------------------------------------
' Helpers
' ------------------------------------------------------------
Private Function GetFilenameFromPath(p As String) As String
    GetFilenameFromPath = Mid(p, InStrRev(p, "/") + 1)
End Function

Private Function GetModuleNameFromFile(filePath As String) As String
    Dim n As String
    n = GetFilenameFromPath(filePath)
    GetModuleNameFromFile = Left(n, InStrRev(n, ".") - 1)
End Function
