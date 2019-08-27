VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private msListName As String
Private mvCSVList As Variant
Private mlListSize As Long
Private mnNamedList As Integer

Public Sub SetList(CSVList() As String)

   mvCSVList = CSVList()
   mlListSize = UBound(CSVList, 1)
   mlNumParams = UBound(CSVList, 2)
   
End Sub

Public Property Get ListSize() As Long

   ListSize = mlListSize
   
End Property

Public Function GetListItem(LineIndex As Integer, ItemOnLine As Integer) As Variant

   GetListItem = mvCSVList(LineIndex, ItemOnLine)
   
End Function

Public Property Get NamedList() As Integer

   NamedList = mnNamedList
   
End Property

Public Property Let ListName(ByVal NameOfList As String)

   msListName = NameOfList
   mnNamedList = Not (msListName = "")

End Property

Public Property Get ListName() As String

   ListName = msListName
   
End Property
