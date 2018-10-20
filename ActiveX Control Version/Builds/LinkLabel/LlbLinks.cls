VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LlbLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private WithEvents PropEnumeration As Enumeration
Attribute PropEnumeration.VB_VarHelpID = -1

Private Function ShadowLinkLabel() As LinkLabel
ComCtlsPtrToShadowObj ShadowLinkLabel, ShadowObjectPointer
End Function

Private Sub PropEnumeration_GetNextItem(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long, ByRef VNextItem As Variant, ByRef NoMoreItems As Boolean)
ShadowLinkLabel.FLinksNextItem Index, Control, Data, VNextItem, NoMoreItems
If NoMoreItems = True Then Set PropEnumeration = Nothing
End Sub

Private Sub PropEnumeration_Reset(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long)
Index = -1
Control = ShadowLinkLabel.FLinksCount()
End Sub

Private Sub PropEnumeration_Skip(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long, ByVal SkipCount As Long, ByRef SkippedAll As Boolean)
Index = Index + SkipCount
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set PropEnumeration = New Enumeration
Set NewEnum = PropEnumeration.GetEnum(ShadowLinkLabel.FLinksCount())
End Function

Friend Sub FInit(ByVal Object As LinkLabel)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Public Property Get Item(ByVal Index As Long) As LlbLink
Attribute Item.VB_Description = "Returns a link given its index."
Attribute Item.VB_UserMemId = 0
If Index > Me.Count Or Index <= 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim NewLink As New LlbLink
NewLink.FInit ShadowObjectPointer, Index
Set Item = NewLink
End Property

Public Function Exists(ByVal Index As Long) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index in the links collection exists."
Exists = CBool(Index <= Me.Count And Index > 0)
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of links in the collection."
Count = ShadowLinkLabel.FLinksCount()
End Property
