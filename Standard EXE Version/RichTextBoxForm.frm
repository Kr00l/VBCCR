VERSION 5.00
Begin VB.Form RichTextBoxForm 
   Caption         =   "RichTextBox Demo"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12780
   KeyPreview      =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin ComCtlsDemo.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   741
      FixedOrder      =   -1  'True
      InitBands       =   "RichTextBoxForm.frx":0000
      Begin ComCtlsDemo.FontCombo FontCombo2 
         Height          =   315
         Left            =   11415
         TabIndex        =   2
         Top             =   45
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Style           =   0
         RecentMax       =   3
      End
      Begin ComCtlsDemo.ToolBar ToolBar1 
         Height          =   360
         Left            =   60
         Top             =   30
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   635
         TextAlignment   =   1
         Wrappable       =   0   'False
         AllowCustomize  =   0   'False
         ButtonWidth     =   83
         InitButtons     =   "RichTextBoxForm.frx":0290
      End
      Begin ComCtlsDemo.FontCombo FontCombo1 
         Height          =   315
         Left            =   8955
         TabIndex        =   1
         Top             =   45
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         BuddyControl    =   "FontCombo2"
         Style           =   0
         FontType        =   2
         RecentMax       =   3
      End
   End
   Begin ComCtlsDemo.RichTextBox RichTextBox1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   10186
      HideSelection   =   0   'False
      MultiLine       =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   "RichTextBoxForm.frx":06C4
   End
End
Attribute VB_Name = "RichTextBoxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
#If VBA7 Then
Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal LCID As Long, ByVal LCType As Long, ByVal lpLCData As LongPtr, ByVal cchData As Long) As Long
Private Declare PtrSafe Function SetActiveWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
#Else
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
#End If
Private LocaleMeasure As Long
Private FindDialogHandle As LongPtr
Private FontComboFreezeClick As Boolean
Private RichTextBoxFreezeSelChange As Boolean
Private CommonDialogPrinter As CommonDialog
Private WithEvents CommonDialogPageSetup As CommonDialog
Attribute CommonDialogPageSetup.VB_VarHelpID = -1
Private WithEvents CommonDialogFont As CommonDialog
Attribute CommonDialogFont.VB_VarHelpID = -1
Private WithEvents CommonDialogFind As CommonDialog
Attribute CommonDialogFind.VB_VarHelpID = -1

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' This is a work around to provide accelerator key access for the tool bar.
' Forms 'KeyPreview' property must be set to true.
' A solution similar to the tab strip control is not possible,
' due to the fact that the tool bar control cannot receive focus.
ToolBar1.ContainerKeyDown KeyCode, Shift
End Sub

Private Sub CommonDialogPageSetup_Help(Handled As Boolean, ByVal Action As Integer, ByVal hDlg As Long)
MsgBox "Help button was pushed within the page setup dialog." & vbLf & "Handle of the dialog: " & hDlg, vbSystemModal + vbOKOnly
Handled = True
End Sub

Private Sub CommonDialogFind_FindNext()
Dim Options As RtfFindOptionConstants, RetVal As Long
If (CommonDialogFind.Flags And CdlFRWholeWord) = CdlFRWholeWord Then Options = RtfFindOptionWholeWord
If (CommonDialogFind.Flags And CdlFRMatchCase) = CdlFRMatchCase Then Options = Options Or RtfFindOptionMatchCase
If (CommonDialogFind.Flags And CdlFRDown) = CdlFRDown Then
    RetVal = RichTextBox1.Find(CommonDialogFind.FindWhat, RichTextBox1.SelStart + RichTextBox1.SelLength, , Options)
Else
    Options = Options Or RtfFindOptionReverse
    RetVal = RichTextBox1.Find(CommonDialogFind.FindWhat, RichTextBox1.SelStart, , Options)
End If
If RetVal = -1 Then MsgBox "Could not find '" & CommonDialogFind.FindWhat & "'.", vbInformation + vbOKOnly + vbSystemModal
End Sub

Private Sub Form_Load()
Call SetupVisualStylesFixes(Me)
Set CommonDialogPrinter = New CommonDialog
CommonDialogPrinter.PrinterDefault = False
CommonDialogPrinter.PrinterDefaultInit = False
Set CommonDialogPageSetup = New CommonDialog
Const LOCALE_USER_DEFAULT As Long = &H400
Const LOCALE_IMEASURE As Long = &HD, LOCALE_RETURN_NUMBER As Long = &H20000000
' cchData = sizeof(DWORD) / sizeof(TCHAR)
' That is, 2 for Unicode and 4 for ANSI.
GetLocaleInfo LOCALE_USER_DEFAULT, LOCALE_IMEASURE Or LOCALE_RETURN_NUMBER, VarPtr(LocaleMeasure), 2
CommonDialogPageSetup.PageLeftMargin = IIf(LocaleMeasure = 0, 2500, 1000)
CommonDialogPageSetup.PageTopMargin = IIf(LocaleMeasure = 0, 2500, 1000)
CommonDialogPageSetup.PageRightMargin = IIf(LocaleMeasure = 0, 2500, 1000)
CommonDialogPageSetup.PageBottomMargin = IIf(LocaleMeasure = 0, 2500, 1000)
CommonDialogPageSetup.Flags = IIf(LocaleMeasure = 0, CdlPSDInHundredthsOfMillimeters, CdlPSDInThousandthsOfInches) Or CdlPSDDefaultMinMargins Or CdlPSDMargins Or CdlPSDHelpButton
CommonDialogPageSetup.PrinterDefault = False
CommonDialogPageSetup.PrinterDefaultInit = False
Set CommonDialogFont = New CommonDialog
Set CommonDialogFind = New CommonDialog
CommonDialogFind.Flags = CdlFRDown
End Sub

Private Sub Form_Resize()
With Me
If .WindowState <> vbMinimized Then
    On Error Resume Next
    RichTextBox1.Move 0, ToolBar1.Height, .ScaleWidth, .ScaleHeight - ToolBar1.Height
End If
End With
End Sub

Private Sub ToolBar1_ButtonClick(ByVal Button As TbrButton)
Select Case Button.Caption
    Case "Bold"
        RichTextBox1.SelBold = Not RichTextBox1.SelBold
    Case "Italic"
        RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
    Case "Underline"
        RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
    Case "Print"
        Dim TwipsMargins(0 To 3) As Long
        With CommonDialogPageSetup
        TwipsMargins(0) = CLng(Me.ScaleX(.PageLeftMargin / IIf(LocaleMeasure = 0, 100, 1000), IIf(LocaleMeasure = 0, vbMillimeters, vbInches), vbTwips))
        TwipsMargins(1) = CLng(Me.ScaleY(.PageTopMargin / IIf(LocaleMeasure = 0, 100, 1000), IIf(LocaleMeasure = 0, vbMillimeters, vbInches), vbTwips))
        TwipsMargins(2) = CLng(Me.ScaleX(.PageRightMargin / IIf(LocaleMeasure = 0, 100, 1000), IIf(LocaleMeasure = 0, vbMillimeters, vbInches), vbTwips))
        TwipsMargins(3) = CLng(Me.ScaleY(.PageBottomMargin / IIf(LocaleMeasure = 0, 100, 1000), IIf(LocaleMeasure = 0, vbMillimeters, vbInches), vbTwips))
        End With
        With CommonDialogPrinter
        .Flags = CdlPDNoCurrentPage Or CdlPDUseDevModeCopiesAndCollate Or CdlPDReturnDC
        If RichTextBox1.GetSelType = RtfSelTypeEmpty Then .Flags = .Flags Or CdlPDNoSelection
        Select Case .ShowPrinterEx
            Case CdlPDResultPrint
                CommonDialogPageSetup.Orientation = .Orientation
                CommonDialogPageSetup.PaperSize = .PaperSize
                CommonDialogPageSetup.PaperBin = .PaperBin
                CommonDialogPageSetup.PrinterDriver = .PrinterDriver
                CommonDialogPageSetup.PrinterName = .PrinterName
                CommonDialogPageSetup.PrinterPort = .PrinterPort
                If (.Flags And CdlPDSelection) = 0 Then
                    RichTextBox1.PrintDoc .hDC, , , TwipsMargins(0), TwipsMargins(1), TwipsMargins(2), TwipsMargins(3)
                Else
                    RichTextBox1.SelPrint .hDC, , , TwipsMargins(0), TwipsMargins(1), TwipsMargins(2), TwipsMargins(3)
                End If
            Case CdlPDResultApply
                CommonDialogPageSetup.Orientation = .Orientation
                CommonDialogPageSetup.PaperSize = .PaperSize
                CommonDialogPageSetup.PaperBin = .PaperBin
                CommonDialogPageSetup.PrinterDriver = .PrinterDriver
                CommonDialogPageSetup.PrinterName = .PrinterName
                CommonDialogPageSetup.PrinterPort = .PrinterPort
        End Select
        End With
    Case "Page Setup"
        Dim Result As Boolean
        With CommonDialogPageSetup
        On Error Resume Next
        Result = .ShowPageSetup
        On Error GoTo 0
        If Result = True Then
            CommonDialogPrinter.Orientation = .Orientation
            CommonDialogPrinter.PaperSize = .PaperSize
            CommonDialogPrinter.PaperBin = .PaperBin
        End If
        End With
    Case "Font"
        With CommonDialogFont
        .HookEvents = True
        .Flags = CdlCFScreenFonts Or CdlCFEffects Or CdlCFApply Or CdlCFLimitSize
        If Not IsNull(RichTextBox1.SelFontName) Then .FontName = RichTextBox1.SelFontName Else .Flags = .Flags Or CdlCFNoFaceSel
        If Not IsNull(RichTextBox1.SelBold) Then .FontBold = RichTextBox1.SelBold Else .Flags = .Flags Or CdlCFNoStyleSel
        If Not IsNull(RichTextBox1.SelItalic) Then .FontItalic = RichTextBox1.SelItalic Else If (.Flags And CdlCFNoStyleSel) = 0 Then .Flags = .Flags Or CdlCFNoStyleSel
        If Not IsNull(RichTextBox1.SelFontSize) Then .FontSize = RichTextBox1.SelFontSize Else .Flags = .Flags Or CdlCFNoSizeSel
        If Not IsNull(RichTextBox1.SelStrikethru) Then .FontStrikethru = RichTextBox1.SelStrikethru Else If (.Flags And CdlCFEffects) = CdlCFEffects Then .Flags = .Flags And Not CdlCFEffects
        If Not IsNull(RichTextBox1.SelUnderline) Then .FontUnderline = RichTextBox1.SelUnderline Else If (.Flags And CdlCFEffects) = CdlCFEffects Then .Flags = .Flags And Not CdlCFEffects
        If Not IsNull(RichTextBox1.SelColor) Then .Color = RichTextBox1.SelColor Else If (.Flags And CdlCFEffects) = CdlCFEffects Then .Flags = .Flags And Not CdlCFEffects
        If Not IsNull(RichTextBox1.SelFontCharset) Then .FontCharset = RichTextBox1.SelFontCharset Else .Flags = .Flags Or CdlCFNoScriptSel
        .Min = 6
        .Max = 72
        If .ShowFont = True Then
            If (.Flags And CdlCFNoFaceSel) = 0 Then RichTextBox1.SelFontName = .FontName
            If (.Flags And CdlCFNoStyleSel) = 0 Then
                RichTextBox1.SelBold = .FontBold
                RichTextBox1.SelItalic = .FontItalic
            End If
            If (.Flags And CdlCFNoSizeSel) = 0 Then RichTextBox1.SelFontSize = .FontSize
            If (.Flags And CdlCFEffects) = CdlCFEffects Then
                RichTextBox1.SelStrikethru = .FontStrikethru
                RichTextBox1.SelUnderline = .FontUnderline
                If RichTextBox1.SelColor <> .Color Then RichTextBox1.SelColor = .Color
            End If
            If (.Flags And CdlCFNoScriptSel) = 0 Then RichTextBox1.SelFontCharset = .FontCharset
        End If
        End With
    Case "&Find"
        Dim RetVal As LongPtr
        RetVal = CommonDialogFind.ShowFind
        If RetVal <> NULL_PTR Then
            FindDialogHandle = RetVal
        Else
            If FindDialogHandle <> NULL_PTR Then SetActiveWindow FindDialogHandle
        End If
End Select
End Sub

Private Sub CommonDialogFont_FontApply(ByVal Flags As Long, ByVal FontName As String, ByVal FontSize As Single, ByVal FontBold As Boolean, ByVal FontItalic As Boolean, ByVal FontStrikethru As Boolean, ByVal FontUnderline As Boolean, ByVal FontCharset As Integer, ByVal RGBColor As Long, ByVal hDlg As Long)
If (Flags And CdlCFNoFaceSel) = 0 Then RichTextBox1.SelFontName = FontName
If (Flags And CdlCFNoStyleSel) = 0 Then
    RichTextBox1.SelBold = FontBold
    RichTextBox1.SelItalic = FontItalic
End If
If (Flags And CdlCFNoSizeSel) = 0 Then RichTextBox1.SelFontSize = FontSize
If (Flags And CdlCFEffects) = CdlCFEffects Then
    RichTextBox1.SelStrikethru = FontStrikethru
    RichTextBox1.SelUnderline = FontUnderline
    If RichTextBox1.SelColor <> RGBColor Then RichTextBox1.SelColor = RGBColor
End If
If (Flags And CdlCFNoScriptSel) = 0 Then RichTextBox1.SelFontCharset = FontCharset
End Sub

Private Sub FontCombo1_Click()
If FontComboFreezeClick = True Then Exit Sub
RichTextBoxFreezeSelChange = True
If FontCombo1.ListIndex > -1 Then
    RichTextBox1.SelFontName = FontCombo1.Text
    If IsNull(RichTextBox1.SelFontSize) Then
        FontCombo2.ListIndex = -1
    Else
        On Error Resume Next
        FontCombo2.Text = CStr(CLng(RichTextBox1.SelFontSize))
        On Error GoTo 0
    End If
End If
RichTextBoxFreezeSelChange = False
End Sub

Private Sub FontCombo1_CloseUp()
RichTextBox1.SetFocus
End Sub

Private Sub FontCombo2_Change()
If FontCombo2.ListIndex = -1 Then
    RichTextBoxFreezeSelChange = True
    On Error Resume Next
    RichTextBox1.SelFontSize = CLng(FontCombo2.Text)
    On Error GoTo 0
    RichTextBoxFreezeSelChange = False
End If
End Sub

Private Sub FontCombo2_Click()
If FontComboFreezeClick = True Then Exit Sub
RichTextBoxFreezeSelChange = True
If FontCombo2.ListIndex > -1 Then
    On Error Resume Next
    RichTextBox1.SelFontSize = CLng(FontCombo2.Text)
    On Error GoTo 0
End If
RichTextBoxFreezeSelChange = False
End Sub

Private Sub FontCombo2_CloseUp()
RichTextBox1.SetFocus
End Sub

Private Sub RichTextBox1_SelChange(ByVal SelType As Integer, ByVal SelStart As Long, ByVal SelEnd As Long)
If RichTextBoxFreezeSelChange = True Then Exit Sub
If (SelType And RtfSelTypeText) <> 0 Or SelType = RtfSelTypeEmpty Then
    FontComboFreezeClick = True
    If IsNull(RichTextBox1.SelFontName) Then
        FontCombo1.ListIndex = -1
        FontCombo2.ListIndex = -1
    Else
        FontCombo1.Text = RichTextBox1.SelFontName
        If IsNull(RichTextBox1.SelFontSize) Then
            FontCombo2.ListIndex = -1
        Else
            On Error Resume Next
            FontCombo2.Text = CStr(CLng(RichTextBox1.SelFontSize))
            On Error GoTo 0
        End If
    End If
    FontComboFreezeClick = False
End If
End Sub
