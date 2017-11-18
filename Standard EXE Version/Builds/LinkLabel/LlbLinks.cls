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
Attribute VB_Exposed = False
Option Explicit
Private ShadowObjectPointer As Long
Private PropLink As Collection

Private Function ShadowLinkLabel() As LinkLabel
ComCtlsPtrToShadowObj ShadowLinkLabel, ShadowObjectPointer
End Function

Private Sub Class_Terminate()
Set PropLink = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set PropLink = New Collection
Dim Count As Long
Count = Me.Count
If Count > 0 Then
    Dim i As Long, NewLink As LlbLink
    For i = 1 To Count
        Set NewLink = New LlbLink
        NewLink.FInit ShadowLinkLabel, i
        PropLink.Add NewLink
    Next i
End If
Set NewEnum = PropLink.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As LinkLabel)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Public Property Get Item(ByVal Index As Long) As LlbLink
Attribute Item.VB_Description = "Returns a link given its index."
Attribute Item.VB_UserMemId = 0
If Index > Me.Count Or Index <= 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim NewLink As New LlbLink
NewLink.FInit ShadowLinkLabel, Index
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
