Attribute VB_Name = "GitHubImporter"
' ============================================================
' GitHubImporter.bas
' Fetched from GitHub by Bootstrap. Owns all sequencing after
' the initial bootstrap — fetches config, parses manifest,
' imports all modules, and hands off to migration runner.
' ============================================================

' Module-level config — populated once from Secret Manager
Private mGithubToken    As String
Private mGithubOwner    As String
Private mGithubRepo     As String
Private mGithubBranch   As String
Private mManifestPath   As String

' ------------------------------------------------------------
' Entry point — called by Bootstrap with config JSON only
' ------------------------------------------------------------
Public Sub ImportModulesFromGitHub(configJson As String)
    ' 1. Parse config into module-level variables
    If Not ParseConfig(configJson) Then
        MsgBox "Importer: config JSON is missing one or more required values.", vbCritical
        Exit Sub
    End If

    ' 2. Fetch and parse manifest
    Dim manifestJson As String
    manifestJson = FetchRaw(ManifestUrl())
    If manifestJson = "" Then
        MsgBox "Importer: failed to fetch manifest.json from GitHub.", vbCritical
        Exit Sub
    End If

    ' 4. Import standard and class modules (identical import path)
    Dim stdModules() As String
    Dim clsModules() As String
    stdModules = ParseJsonArray(manifestJson, "modules")
    clsModules = ParseJsonArray(manifestJson, "classModules")

    Dim f As Variant
    For Each f In stdModules
        If Trim(CStr(f)) <> "" Then ImportSingleModule Trim(CStr(f)), False
    Next f
    For Each f In clsModules
        If Trim(CStr(f)) <> "" Then ImportSingleModule Trim(CStr(f)), False
    Next f

    ' 5. Inject document modules in-place
    Dim docModules() As String
    docModules = ParseJsonArray(manifestJson, "documentModules")
    For Each f In docModules
        If Trim(CStr(f)) <> "" Then ImportSingleModule Trim(CStr(f)), True
    Next f

    ' 6. Hand off migrations to runner (if loaded)
    Dim migrations() As String
    migrations = ParseJsonArray(manifestJson, "migrations")

    On Error Resume Next
    Dim runner As Object
    Set runner = ThisWorkbook.VBProject.VBComponents("GitHubMigrationRunner")
    On Error GoTo 0

    If Not runner Is Nothing Then
        Application.Run "'" & ThisWorkbook.Name & "'!GitHubMigrationRunner.RunMigrationsFromGitHub", _
                        migrations, mGithubToken, mGithubOwner, mGithubRepo, mGithubBranch
    End If
End Sub

' ------------------------------------------------------------
' Parses config JSON into module-level variables
' ------------------------------------------------------------
Private Function ParseConfig(json As String) As Boolean
    mGithubToken  = ExtractJsonValue(json, "github_token")
    mGithubOwner  = ExtractJsonValue(json, "github_owner")
    mGithubRepo   = ExtractJsonValue(json, "github_repo")
    mGithubBranch = ExtractJsonValue(json, "github_branch")
    mManifestPath = ExtractJsonValue(json, "manifest_path")

    ParseConfig = (mGithubToken <> "" And mGithubOwner <> "" And _
                   mGithubRepo <> "" And mGithubBranch <> "" And _
                   mManifestPath <> "")
End Function

' ------------------------------------------------------------
' Builds the GitHub API URL for the manifest
' ------------------------------------------------------------
Private Function ManifestUrl() As String
    ManifestUrl = "https://api.github.com/repos/" & mGithubOwner & "/" & mGithubRepo & _
                  "/contents/" & mManifestPath & "?ref=" & mGithubBranch
End Function

' ------------------------------------------------------------
' Fetches a GitHub file and returns its decoded text content
' ------------------------------------------------------------
Private Function FetchRaw(url As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    On Error GoTo Fail
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "token " & mGithubToken
    http.setRequestHeader "Accept", "application/vnd.github.v3+json"
    http.setRequestHeader "User-Agent", "Excel-VBA"
    http.Send

    If http.Status = 200 Then
        Dim b64 As String
        b64 = ExtractJsonValue(http.responseText, "content")
        b64 = Join(Split(b64, "\n"), "")
        FetchRaw = DecodeBase64ToString(b64)
    End If
    Exit Function
Fail:
End Function

' ------------------------------------------------------------
' Downloads and imports/injects a single module
' ------------------------------------------------------------
Private Sub ImportSingleModule(repoFilePath As String, isDocModule As Boolean)
    Dim url As String
    url = "https://api.github.com/repos/" & mGithubOwner & "/" & mGithubRepo & _
          "/contents/" & repoFilePath & "?ref=" & mGithubBranch

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    On Error GoTo Fail
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "token " & mGithubToken
    http.setRequestHeader "Accept", "application/vnd.github.v3+json"
    http.setRequestHeader "User-Agent", "Excel-VBA"
    http.Send

    If http.Status <> 200 Then GoTo Fail

    Dim b64 As String
    b64 = ExtractJsonValue(http.responseText, "content")
    b64 = Join(Split(b64, "\n"), "")

    Dim tempPath As String
    tempPath = Environ("TEMP") & "\" & GetFilenameFromPath(repoFilePath)
    WriteBase64ToFile b64, tempPath

    Dim compName As String
    compName = GetModuleNameFromFile(tempPath)

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    If isDocModule Then
        InjectCodeIntoDocumentModule vbProj, compName, tempPath
    Else
        On Error Resume Next
        vbProj.VBComponents.Remove vbProj.VBComponents(compName)
        On Error GoTo 0
        vbProj.VBComponents.Import tempPath
    End If

    Kill tempPath
    Exit Sub
Fail:
    MsgBox "Failed to import: " & repoFilePath, vbCritical
End Sub

' ------------------------------------------------------------
' Injects code into an existing document module in-place
' ------------------------------------------------------------
Private Sub InjectCodeIntoDocumentModule(vbProj As Object, compName As String, filePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1)
    Dim fileContents As String
    fileContents = ts.ReadAll
    ts.Close

    Dim lines() As String
    lines = Split(fileContents, vbLf)

    Dim codeStart As Long
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim trimmed As String
        trimmed = Trim(Replace(lines(i), vbCr, ""))
        If Len(trimmed) > 0 And Left(trimmed, 9) <> "Attribute" Then
            codeStart = i
            Exit For
        End If
    Next i

    Dim codeLines() As String
    ReDim codeLines(UBound(lines) - codeStart)
    For i = codeStart To UBound(lines)
        codeLines(i - codeStart) = lines(i)
    Next i

    Dim comp As Object
    Set comp = vbProj.VBComponents(compName)
    With comp.CodeModule
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, Join(codeLines, vbLf)
    End With
End Sub

' ------------------------------------------------------------
' GCP metadata server — returns short-lived access token
' ------------------------------------------------------------
Private Function GetMetadataToken() As String
    Const url As String = "http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token"
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    On Error GoTo Fail
    http.Open "GET", url, False
    http.setRequestHeader "Metadata-Flavor", "Google"
    http.Send
    If http.Status = 200 Then
        GetMetadataToken = ExtractJsonValue(http.responseText, "access_token")
    End If
    Exit Function
Fail:
End Function

' ------------------------------------------------------------
' Secret Manager — retrieves and decodes a secret version
' ------------------------------------------------------------
Private Function GetSecret(accessToken As String, project As String, _
                           secretName As String, version As String) As String
    Dim url As String
    url = "https://secretmanager.googleapis.com/v1/projects/" & project & _
          "/secrets/" & secretName & "/versions/" & version & ":access"

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    On Error GoTo Fail
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Bearer " & accessToken
    http.Send
    If http.Status = 200 Then
        Dim b64 As String
        b64 = ExtractJsonValue(ExtractJsonValue(http.responseText, "payload"), "data")
        GetSecret = DecodeBase64ToString(b64)
    End If
    Exit Function
Fail:
End Function

' ------------------------------------------------------------
' Parses a flat JSON array by key
' ------------------------------------------------------------
Private Function ParseJsonArray(json As String, key As String) As String()
    Dim emptyResult() As String
    ReDim emptyResult(0)
    emptyResult(0) = ""

    Dim pattern As String
    pattern = """" & key & """:"
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then
        ParseJsonArray = emptyResult
        Exit Function
    End If

    pos = InStr(pos, json, "[")
    If pos = 0 Then
        ParseJsonArray = emptyResult
        Exit Function
    End If

    Dim endPos As Long
    endPos = InStr(pos, json, "]")
    If endPos = 0 Then
        ParseJsonArray = emptyResult
        Exit Function
    End If

    Dim arrayContent As String
    arrayContent = Mid(json, pos + 1, endPos - pos - 1)

    Dim results() As String
    ReDim results(50)
    Dim count As Integer
    count = 0

    Dim p As Long
    p = 1
    Do
        Dim q1 As Long
        q1 = InStr(p, arrayContent, """")
        If q1 = 0 Then Exit Do
        Dim q2 As Long
        q2 = InStr(q1 + 1, arrayContent, """")
        If q2 = 0 Then Exit Do
        results(count) = Mid(arrayContent, q1 + 1, q2 - q1 - 1)
        count = count + 1
        p = q2 + 1
    Loop

    If count = 0 Then
        ParseJsonArray = emptyResult
        Exit Function
    End If

    ReDim Preserve results(count - 1)
    ParseJsonArray = results
End Function

' ------------------------------------------------------------
' Minimal JSON value extractor (strings and nested objects)
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

    Dim result As String
    Dim c As String
    If Mid(json, pos, 1) = """" Then
        pos = pos + 1
        Do
            c = Mid(json, pos, 1)
            If c = """" Then Exit Do
            result = result & c
            pos = pos + 1
        Loop
    ElseIf Mid(json, pos, 1) = "{" Then
        Dim depth As Integer
        depth = 0
        Do
            c = Mid(json, pos, 1)
            If c = "{" Then depth = depth + 1
            If c = "}" Then depth = depth - 1
            result = result & c
            pos = pos + 1
            If depth = 0 Then Exit Do
        Loop
    End If

    ExtractJsonValue = result
End Function

' ------------------------------------------------------------
' Decodes Base64 to a UTF-8 string
' ------------------------------------------------------------
Private Function DecodeBase64ToString(b64 As String) As String
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
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "UTF-8"
    DecodeBase64ToString = stream.ReadText
    stream.Close
End Function

' ------------------------------------------------------------
' Writes raw Base64-encoded bytes to a temp file
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
    Dim pos As Long
    pos = IIf(InStrRev(p, "\") > InStrRev(p, "/"), InStrRev(p, "\"), InStrRev(p, "/"))
    GetFilenameFromPath = Mid(p, pos + 1)
End Function

Private Function GetModuleNameFromFile(filePath As String) As String
    Dim n As String
    n = GetFilenameFromPath(filePath)
    GetModuleNameFromFile = Left(n, InStrRev(n, ".") - 1)
End Function
