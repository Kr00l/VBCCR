VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If False Then
Private OLEDropModeNone, OLEDropModeManual
Private CCMousePointerDefault, CCMousePointerArrow, CCMousePointerCrosshair, CCMousePointerIbeam, CCMousePointerHand, CCMousePointerSizePointer, CCMousePointerSizeNESW, CCMousePointerSizeNS, CCMousePointerSizeNWSE, CCMousePointerSizeWE, CCMousePointerUpArrow, CCMousePointerHourglass, CCMousePointerNoDrop, CCMousePointerArrowHourglass, CCMousePointerArrowQuestion, CCMousePointerSizeAll, CCMousePointerArrowCD, CCMousePointerCustom
Private CCAppearanceFlat, CCAppearance3D
Private CCBorderStyleNone, CCBorderStyleSingle, CCBorderStyleThin, CCBorderStyleSunken, CCBorderStyleRaised
Private CCBackStyleTransparent, CCBackStyleOpaque
Private CCLeftRightAlignmentLeft, CCLeftRightAlignmentRight
Private CCVerticalAlignmentTop, CCVerticalAlignmentCenter, CCVerticalAlignmentBottom
Private CCIMEModeNoControl, CCIMEModeOn, CCIMEModeOff, CCIMEModeDisable, CCIMEModeHiragana, CCIMEModeKatakana, CCIMEModeKatakanaHalf, CCIMEModeAlphaFull, CCIMEModeAlpha, CCIMEModeHangulFull, CCIMEModeHangul
Private CCRightToLeftModeNoControl, CCRightToLeftModeVBAME, CCRightToLeftModeSystemLocale, CCRightToLeftModeUserLocale, CCRightToLeftModeOSLanguage
Private CCScrollOrientationHorizontal, CCScrollOrientationVertical, CCScrollOrientationBoth
#End If
Public Enum OLEDropModeConstants
OLEDropModeNone = vbOLEDropNone
OLEDropModeManual = vbOLEDropManual
End Enum
Public Enum CCMousePointerConstants
CCMousePointerDefault = 0
CCMousePointerArrow = 1
CCMousePointerCrosshair = 2
CCMousePointerIbeam = 3
CCMousePointerHand = 4
CCMousePointerSizePointer = 5
CCMousePointerSizeNESW = 6
CCMousePointerSizeNS = 7
CCMousePointerSizeNWSE = 8
CCMousePointerSizeWE = 9
CCMousePointerUpArrow = 10
CCMousePointerHourglass = 11
CCMousePointerNoDrop = 12
CCMousePointerArrowHourglass = 13
CCMousePointerArrowQuestion = 14
CCMousePointerSizeAll = 15
CCMousePointerArrowCD = 16
CCMousePointerCustom = 99
End Enum
Public Enum CCAppearanceConstants
CCAppearanceFlat = 0
CCAppearance3D = 1
End Enum
Public Enum CCBorderStyleConstants
CCBorderStyleNone = 0
CCBorderStyleSingle = 1
CCBorderStyleThin = 2
CCBorderStyleSunken = 3
CCBorderStyleRaised = 4
End Enum
Public Enum CCBackStyleConstants
CCBackStyleTransparent = 0
CCBackStyleOpaque = 1
End Enum
Public Enum CCLeftRightAlignmentConstants
CCLeftRightAlignmentLeft = 0
CCLeftRightAlignmentRight = 1
End Enum
Public Enum CCVerticalAlignmentConstants
CCVerticalAlignmentTop = 0
CCVerticalAlignmentCenter = 1
CCVerticalAlignmentBottom = 2
End Enum
Public Enum CCIMEModeConstants
CCIMEModeNoControl = 0
CCIMEModeOn = 1
CCIMEModeOff = 2
CCIMEModeDisable = 3
CCIMEModeHiragana = 4
CCIMEModeKatakana = 5
CCIMEModeKatakanaHalf = 6
CCIMEModeAlphaFull = 7
CCIMEModeAlpha = 8
CCIMEModeHangulFull = 9
CCIMEModeHangul = 10
End Enum
Public Enum CCRightToLeftModeConstants
CCRightToLeftModeNoControl = 0
CCRightToLeftModeVBAME = 1
CCRightToLeftModeSystemLocale = 2
CCRightToLeftModeUserLocale = 3
CCRightToLeftModeOSLanguage = 4
End Enum
Public Enum CCScrollOrientationConstants
CCScrollOrientationHorizontal = 0
CCScrollOrientationVertical = 1
CCScrollOrientationBoth = 2
End Enum

#If VBA7 Then
Public Function Message(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
Attribute Message.VB_Description = "Interface function to forward window messages."
#Else
Public Function Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Attribute Message.VB_Description = "Interface function to forward window messages."
#End If
End Function
