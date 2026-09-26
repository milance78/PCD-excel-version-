Option Explicit
Const VERSION_URL = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/VERSION.json"
Const ARTIFACT_URL = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/PCD-Excel-Version-latest.xlsm"

Dim fso, sh, http, stm, folder, manifest, version, target, tmp
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
folder = fso.BuildPath(sh.ExpandEnvironmentStrings("%LOCALAPPDATA%"), "PCD-Excel")
If Not fso.FolderExists(folder) Then fso.CreateFolder(folder)
target = fso.BuildPath(folder, "PCD-Excel-Version-latest.xlsm")
tmp = target & ".download"

Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", VERSION_URL, False
http.SetRequestHeader "User-Agent", "PCD-Excel-Updater"
http.Send
If http.Status <> 200 Then
  MsgBox "PCD update: GitHub n'est pas accessible (HTTP " & http.Status & ")."
  WScript.Quit 2
End If

manifest = http.ResponseText
version = JsonString(manifest, "version")
If version = "" Then
  MsgBox "PCD update: VERSION.json est invalide."
  WScript.Quit 3
End If

http.Open "GET", ARTIFACT_URL, False
http.SetRequestHeader "User-Agent", "PCD-Excel-Updater"
http.Send
If http.Status <> 200 Then
  MsgBox "PCD update: le XLSM n'est pas disponible (HTTP " & http.Status & ")."
  WScript.Quit 4
End If

Set stm = CreateObject("ADODB.Stream")
stm.Type = 1
stm.Open
stm.Write http.ResponseBody
stm.SaveToFile tmp, 2
stm.Close

If fso.FileExists(target) Then fso.DeleteFile target, True
fso.MoveFile tmp, target
sh.Run "excel.exe """ & target & """", 1, False

Function JsonString(json, key)
  Dim re, m
  Set re = CreateObject("VBScript.RegExp")
  re.Global = False
  re.IgnoreCase = True
  re.Pattern = """" & key & """s*:s*""([^""}]*)"""
  Set m = re.Execute(json)
  If m.Count > 0 Then JsonString = m(0).SubMatches(0)
End Function
