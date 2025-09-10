' scripts\fetch_test.vbs
On Error Resume Next

Dim fso, base, dataDir, logsDir, logFile, tf
Set fso = CreateObject("Scripting.FileSystemObject")

base    = fso.GetParentFolderName(WScript.ScriptFullName)
base    = fso.GetParentFolderName(base)
dataDir = fso.BuildPath(base, "data")
logsDir = fso.BuildPath(base, "logs")

If Not fso.FolderExists(dataDir) Then fso.CreateFolder dataDir
If Not fso.FolderExists(logsDir) Then fso.CreateFolder logsDir

logFile = fso.BuildPath(logsDir, "fetch_test.log")
Set tf = fso.OpenTextFile(logFile, 8, True)
tf.WriteLine Now & " | START fetch_test.vbs"

Dim outPath, json
outPath = fso.BuildPath(dataDir, "test.json")

json = "[" & vbCrLf & _
       "  {\"ID\":1,\"Nazwa\":\"Przykład A\",\"Wartosc\":123.45,\"Status\":\"OK\"}," & vbCrLf & _
       "  {\"ID\":2,\"Nazwa\":\"Przykład B\",\"Wartosc\":67.89,\"Status\":\"PENDING\"}," & vbCrLf & _
       "  {\"ID\":3,\"Nazwa\":\"Przykład C\",\"Wartosc\":100,\"Status\":\"OK\"}" & vbCrLf & _
       "]"

Dim outf
Set outf = fso.CreateTextFile(outPath, True, True)
outf.Write json
outf.Close

If Err.Number = 0 Then
  tf.WriteLine Now & " | OK zapisano: " & outPath
  tf.Close
  WScript.Quit 0
Else
  tf.WriteLine Now & " | ERR: " & Err.Description
  tf.Close
  WScript.Quit 1
End If
