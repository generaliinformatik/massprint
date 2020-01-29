'-------------------------------------------------------------------------------
' massprint.vbs
'    !
'    +--- excel               <- put *.xlsx into this subfolder
'    +--- powerpoint          <- put *.pptx into this subfolder
'    +--- word                <- put *.docx into this subfolder
'
'
' Usage:
'
'     cscript massprint.vbs 10       <- print set of above documents 10 times
'
'
'
' https://github.com/generaliinformatik/massprint
' (c) 2020 Generali Informatik
'-------------------------------------------------------------------------------

Option Explicit

' get argument from commandline ------------------------------------------------

Dim printCount
If WScript.Arguments.Count > 0 Then
  printCount = WScript.Arguments.Item(0)
Else
  printCount = 0
End If

' greeting ---------------------------------------------------------------------

WScript.Echo "massprint.vbs"
WScript.Echo "-----------------------------------------------------------------"

' are you sure? ----------------------------------------------------------------

Dim result
result=MsgBox("You are going to print all the documents " + _
              CStr(printCount) + " times.", vbYesNo, "MassPrint")

If result = vbNo Then
  WScript.Echo "You clicked NO"
  WScript.Quit
End if

WScript.Echo "You clicked YES"

' init filesystem access -------------------------------------------------------

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

' setup directories ------------------------------------------------------------

Dim myDir, xlsDir, docDir, pptDir
myDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
xlsDir = myDir + "\excel"
docDir = myDir + "\word"
pptDir = myDir + "\powerpoint"

' init applications ------------------------------------------------------------

Dim objWord, objExcel, objPowerpoint
Set objWord       = CreateObject("Word.Application")
Set objExcel      = CreateObject("Excel.Application")
Set objPowerpoint = CreateObject("Powerpoint.Application")

' make applications visible
objExcel.Visible      = true
objWord.Visible       = true
objPowerpoint.Visible = true

' print documents/workbooks/presentations --------------------------------------

Dim objFile, i

For i = 1 to printCount
  WScript.Echo "  printing: " + CStr(i) + "/" + CStr(printCount)
  ' loop over Excel workbooks
  For Each objFile In objFSO.GetFolder(xlsDir).Files
    If UCase(objFSO.GetExtensionName(objFile.Name)) = "XLSX" Then
      PrintExcel objExcel, objFSO.GetAbsolutePathName(objFile)
    End If
  Next

  ' loop over Word documents
  For Each objFile In objFSO.GetFolder(docDir).Files
    If UCase(objFSO.GetExtensionName(objFile.Name)) = "DOCX" Then
      PrintWord objWord, objFSO.GetAbsolutePathName(objFile)
    End If
  Next

  ' loop over Powerpoint presentations
  For Each objFile In objFSO.GetFolder(pptDir).Files
    If UCase(objFSO.GetExtensionName(objFile.Name)) = "PPTX" Then
      PrintPowerpoint objPowerpoint, objFSO.GetAbsolutePathName(objFile)
    End If
  Next
Next

' destruct applications --------------------------------------------------------

' if we quit too fast, the last document is going to swallowed
WScript.Echo "Waiting 5 seconds to finish..."
WScript.Sleep 5000

objWord.Quit
objExcel.Quit
objPowerpoint.Quit

Set objWord       = Nothing
Set objExcel      = Nothing
Set objPowerpoint = Nothing

' destruct filesystem access
Set objFSO = Nothing

WScript.Echo "Done!"

' functions --------------------------------------------------------------------

Sub PrintExcel(objExcel, file)
  objExcel.Workbooks.Open file
  Dim objSheet
  Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
  objSheet.PrintOut()
  Set objSheet = Nothing
  objExcel.ActiveWorkbook.Close
End Sub

Sub PrintWord(objWord, file)
  objWord.Documents.Open file
  objWord.ActiveDocument.PrintOut()
  objWord.ActiveDocument.Close
End Sub

Sub PrintPowerpoint(objPowerpoint, file)
  objPowerpoint.Presentations.Open file
  objPowerpoint.ActivePresentation.PrintOut()
  objPowerpoint.ActivePresentation.Close
End Sub
