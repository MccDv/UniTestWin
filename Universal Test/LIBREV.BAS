Attribute VB_Name = "LIBREV"
' includes Universal Library functions up to
' and including

'             Revision 4.00


Function SetUpTrigger(BoardNum As Integer, TrigType As Integer, LowThreshold As Integer, HighThreshold As Integer) As Integer

   On Error GoTo BadFunction
   Dim NoForm As Form
   ULStat = cbSetTrigger(BoardNum, TrigType, LowThreshold, HighThreshold)
   x% = SaveFunc(NoForm, SetTrigger, ULStat, BoardNum, TrigType, LowThreshold, HighThreshold, A5, A6, A7, A8, A9, A10, A11, 0)
   SetUpTrigger = ULStat
   Exit Function

BadFunction:
   If Err = 35 Then
      SetUpTrigger = NOTWINDOWSFUNC
      MsgBox "cbSetTrigger is not supported in this version of Universal Library", , "cbSetTrigger Not Defined"
   Else
      MsgBox Error$(Err), , "Runtime Error"
   End If
   Exit Function

End Function

