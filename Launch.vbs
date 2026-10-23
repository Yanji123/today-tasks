Set shell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
folder = fs.GetParentFolderName(WScript.ScriptFullName)
For Each file In fs.GetFolder(folder).Files
 If LCase(fs.GetExtensionName(file.Name)) = "ps1" Then
  shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File """ & file.Path & """", 0, False
  Exit For
 End If
Next