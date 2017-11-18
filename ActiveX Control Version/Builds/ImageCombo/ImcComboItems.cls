VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImcComboItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowObjectPointer As Long
Private PropComboItem As Collection

Private Function ShadowImageCombo() As ImageCombo
ComCtlsPtrToShadowObj ShadowImageCombo, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
Set PropComboItem = New Collection
End Sub

Private Sub Class_Terminate()
Set PropComboItem = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropComboItem.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As ImageCombo)
ShadowObjectPointer = ObjPtr(Object)
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Image As Variant, Optional ByVal SelImage As Variant, Optional ByVal Indentation As Variant) As ImcComboItem
Attribute Add.VB_Description = "Adds a combo item to the combo items collection."
If Index > (PropComboItem.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim ImageIndex As Long, SelImageIndex As Long
If IsMissing(Image) Then
    Image = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowImageCombo, ShadowImageCombo.ImageList, Image, ImageIndex)
End If
If IsMissing(SelImage) Then
    SelImage = Empty
Else
    Call ComCtlsImlListImageIndex(ShadowImageCombo, ShadowImageCombo.ImageList, SelImage, SelImageIndex)
End If
ShadowImageCombo.FComboItemsAdd Index, Text, ImageIndex, SelImageIndex, Indentation
Dim NewComboItem As New ImcComboItem
NewComboItem.FInit ShadowImageCombo, Key, Image, ImageIndex, SelImage, SelImageIndex
If Key = vbNullString Then
    If Index > 0 And Index <= PropComboItem.Count Then
        PropComboItem.Add NewComboItem, , Index
    Else
        PropComboItem.Add NewComboItem
    End If
Else
    If Index > 0 And Index <= PropComboItem.Count Then
        PropComboItem.Add NewComboItem, Key, Index
    Else
        PropComboItem.Add NewComboItem, Key
    End If
End If
If Index = 0 Then
    NewComboItem.Index = PropComboItem.Count
Else
    Dim i As Long
    For i = Index To PropComboItem.Count
        Me.Item(i).Index = i
    Next i
End If
Set Add = NewComboItem
End Function

Public Property Get Item(ByVal Index As Variant) As ImcComboItem
Attribute Item.VB_Description = "Returns a combo item given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropComboItem(Index)
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the combo items collection exists."
On Error Resume Next
PropComboItem.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of combo items in the collection."
Count = PropComboItem.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all combo items from the collection."
ShadowImageCombo.FComboItemsClear
Set PropComboItem = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a combo item from the collection."
ShadowImageCombo.FComboItemsRemove PropComboItem(Index).Index
PropComboItem.Remove Index
Dim i As Long
For i = 1 To PropComboItem.Count
    Me.Item(i).Index = i
Next i
End Sub
