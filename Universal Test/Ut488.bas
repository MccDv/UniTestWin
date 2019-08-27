Attribute VB_Name = "CBGPIBStub"
'A dfferent module exists to handle the functions that are specific
'to CBI's library and not included in NI's library.
'see Utgpib32.bas for compatibility with CBI library.

Sub DoIBINIT()

   'CBI specific function

End Sub

Sub DoIBPTRS()

   'CBI specific function

End Sub

Sub GetGPIBBoardName(Device%, BdDevName$, UsingIni%)

   lpFileName$ = "Cfg488crh.ini"

   lpApplicationName$ = "DevName"
   lpKeyName$ = Format$(Device%, "0")
   nSize% = 16
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = ""

   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   BdDevName$ = Left$(lpReturnedString$, x%)
   UsingIni% = True

End Sub

