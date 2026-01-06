' Outlook Phishing Reporter - Uninstallation Script
' This VBScript provides an easy way to uninstall the add-in

Option Explicit
Dim objShell, objReg, objFSO, strRegPath, intResult

Set objShell = CreateObject("WScript.Shell")
Set objReg = CreateObject("WinReg.CoClass")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Const HKEY_CURRENT_USER = &H80000001
strRegPath = "Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter"

' Display uninstallation warning
intResult = MsgBox("Uninstall Outlook Phishing Reporter add-in?" & vbCrLf & vbCrLf & _
    "This will remove the add-in from Microsoft Outlook." & vbCrLf & vbCrLf & _
    "Click YES to continue, or NO to cancel.", _
    vbYesNo + vbQuestion, _
    "Outlook Phishing Reporter - Uninstall")

If intResult <> vbYes Then
    ' User cancelled
    MsgBox "Uninstallation cancelled.", vbInformation, "Outlook Phishing Reporter"
    WScript.Quit 0
End If

' Close Outlook if running
On Error Resume Next
objShell.Run "taskkill /IM OUTLOOK.EXE /F", 0, False
WScript.Sleep 2000

' Remove registry entry
On Error Resume Next
objShell.RegDelete "HKEY_CURRENT_USER\" & strRegPath

If Err.Number <> 0 Then
    MsgBox "Warning: Could not remove registry entries." & vbCrLf & vbCrLf & _
        "The add-in may not have been installed.", _
        vbExclamation, _
        "Outlook Phishing Reporter - Uninstall Warning"
    WScript.Quit 1
End If

' Success message
MsgBox "Outlook Phishing Reporter has been successfully uninstalled." & vbCrLf & vbCrLf & _
    "You may now restart Outlook.", _
    vbInformation, _
    "Outlook Phishing Reporter - Uninstall Complete"

WScript.Quit 0
