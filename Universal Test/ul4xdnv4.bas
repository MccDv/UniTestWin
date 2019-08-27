Attribute VB_Name = "Version"
'This module is compatable with:
'  VB version 4.0 and up
'  Universal Library versions below 5.0 (Dighton)

Global gaOptions() As Integer, gaMemHandle()  As Integer

Dim manFormID(4) As Integer
Dim mfNoForm As Form
Dim mnIntCount As Integer

'Declare Function cbEnableEvent Lib "cbw32.dll" (ByVal BoardNum&, ByVal EventType&, ByVal ProcAddress As Any) As Long
'Declare Function cbDisableEvent Lib "cbw32.dll" (ByVal BoardNum&, ByVal EventType&) As Long

#If Win32 Then
   Declare Function cbLoadConfig& Lib "cbw32.dll" (ByVal Filename As String)
   Declare Function cbTenUSDelay& Lib "cbw32.dll" (ByVal Count%)  '800 max (to test delay calibration)
#Else
   Declare Function cbLoadConfig% Lib "cbw.dll" ()
   Declare Function cbTenUSDelay% Lib "cbw.dll" (ByVal Count%)  '800 max (to test delay calibration)
#End If

Function GetCfgFile(ByVal Filename$) As Integer

   'file name not used until Rev 5
   'Filename$ = "cb.cfg"
   GetCfgFile = cbLoadConfig(Filename$)
   If SaveFunc(mfNoForm, LoadConfig, GetCfgFile, Filename$, A2, _
      A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
   
End Function

Sub SetCfgDirec()
   
   gsConfigDirec = "\cbdevel\CBConfig\AinCfg\"
   
End Sub

