VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event GetNextItem(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long, ByRef VNextItem As Variant, ByRef NoMoreItems As Boolean)
Public Event Reset(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long)
Public Event Skip(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long, ByVal SkipCount As Long, ByRef SkippedAll As Boolean)
Implements OLEGuids.IEnumVARIANTUnrestricted
Private PropControl As Long, PropIndex As Long, PropData As Long

Private Sub Class_Initialize()
Call SetVTableSubclass(Me, VTableInterfaceEnumeration)
End Sub

Private Sub Class_Terminate()
Call RemoveVTableSubclass(Me, VTableInterfaceEnumeration)
End Sub

Private Sub IEnumVARIANTUnrestricted_Next(ByVal cElt As Long, ByRef rgVar As Variant, ByVal pcEltFetched As Long)
End Sub

Private Sub IEnumVARIANTUnrestricted_Skip(ByVal cElt As Long)
End Sub

Private Sub IEnumVARIANTUnrestricted_Reset()
End Sub

Private Sub IEnumVARIANTUnrestricted_Clone(ByRef ppEnum As IEnumVARIANT)
End Sub

Public Function GetEnum(ByVal Control As Long, Optional ByVal Data As Long, Optional ByVal Index As Long = -1) As IEnumVARIANT
PropIndex = Index
PropControl = Control
PropData = Data
Set GetEnum = Me
End Function

Friend Sub GetNextItem(ByRef VNextItem As Variant, ByRef NoMoreItems As Boolean)
RaiseEvent GetNextItem(PropIndex, PropControl, PropData, VNextItem, NoMoreItems)
End Sub

Friend Sub Reset()
RaiseEvent Reset(PropIndex, PropControl, PropData)
End Sub

Friend Sub Skip(ByVal SkipCount As Long, ByRef SkippedAll As Boolean)
RaiseEvent Skip(PropIndex, PropControl, PropData, SkipCount, SkippedAll)
End Sub
