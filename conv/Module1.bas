Attribute VB_Name = "Module1"
'open web browser (internet explorer)
'******************************************
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&
Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Public Sub RunBrowser(strURL As String, iWindowStyle As Integer, fH As Long)
Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", strURL, 0&, 0&, iWindowStyle)
End Sub
'*******************************************





